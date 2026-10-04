/* Slilot —— 任务窗格
 * 对话 + 工具回路：模型通过工具经 Office.js 实时读写当前演示文稿。
 * API 请求发往同源本地代理（/api/forward），由代理按所选接口格式翻译并转发到上游。
 */
"use strict";

/* ---------------- 配置 ----------------
 * 默认预置的服务商端点/模型可在面板 ⚙ 里全部修改；
 * API Key 出于安全不在代码中保存，首次使用在 ⚙ 里粘贴（存本机 localStorage）。
 */
const DEFAULT_SETTINGS = {
  upstreamBase: "https://api.minimaxi.com/anthropic",
  apiFormat: "messages",
  imageApiUrl: "https://api.minimaxi.com/v1/image_generation",
  imageModel: "image-01",
  apiKey: "",
  model: "MiniMax-M3.1-Flash-Preview",
  maxTokens: 16000,
};
const LS_KEY = "mm_ppt_settings_v4";
let settings = loadSettings();
let capSets = [];            // 宿主支持的 PowerPointApi 版本
let canvas = { w: 960, h: 540 }; // 最近一次检测到的画布尺寸(pt)
let messages = [];           // Anthropic 格式消息数组
let busy = false;
let abortCtrl = null;
let toolCount = 0;           // 当前任务已执行的工具数（状态栏隐式进度）
let runGen = 0;              // 会话代号：新会话/中止时递增，旧回路据此静默退出
let docVersion = 0;          // 文档修改版本号：截图缓存据此失效
let compactionAttempts = 0;  // 摘要压缩尝试次数（每次运行最多 2 次）
// 单位是 JSON 字符数，不是 token：中文 ≈0.7-1 token/字符，base64 ≈0.25 token/字符。
// 截图已改 JPG（~15KB）且只留 1 张，历史大头是中文文本，300K 字符 ≈ 20-30 万 token 时触发才安全。
const COMPACT_THRESHOLD = 300 * 1024;

function loadSettings() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (raw) return Object.assign({}, DEFAULT_SETTINGS, JSON.parse(raw));
  } catch (e) {}
  return Object.assign({}, DEFAULT_SETTINGS);
}
function saveSettings() {
  try { localStorage.setItem(LS_KEY, JSON.stringify(settings)); } catch (e) {}
}

/* ---------------- UI 基础 ---------------- */
const $ = (id) => document.getElementById(id);
function addMsg(kind, text) {
  const div = document.createElement("div");
  div.className = "msg " + kind;
  div.textContent = text;
  $("chat").appendChild(div);
  scrollBottom();
  return div;
}
function addToolChip(label, isErr) {
  const div = document.createElement("div");
  div.className = "toolchip" + (isErr ? " err" : "");
  div.textContent = label;
  $("chat").appendChild(div);
  scrollBottom();
  return div;
}
function scrollBottom() { $("chat").scrollTop = $("chat").scrollHeight; }
function setStatus(text) {
  const sb = $("statusbar");
  if (!text) { sb.classList.add("hidden"); return; }
  sb.textContent = text;
  sb.classList.remove("hidden");
}
function setBusy(v) {
  busy = v;
  const btn = $("sendBtn");
  btn.disabled = false;
  btn.textContent = v ? "停止" : "发送";
}

/* ---------------- 系统提示词 ---------------- */
function systemPrompt() {
  return [
    "你是内嵌在 PowerPoint 任务窗格中的 PPT 智能体（Slilot），通过工具直接操作用户当前打开的演示文稿（真实文档对象模型，修改立即生效并可在 PowerPoint 中撤销）。",
    "",
    "## 工作流程：RIP 循环（Plan → Implement → Review），任何生成/修改任务都必须完整走完三段，跳过 Review 视为未完成。",
    "",
    "### 1) Plan（规划）",
    "- 先 get_presentation_overview 了解现状（空白文档则从零规划）。",
    "- 产出页面级规划：每页的标题、核心内容（要点/表格数据）、配图需求（写好具体生图提示词）、版式布局（各元素的坐标与大小）。",
    "- 用几行向用户展示规划（每页一行），然后立即进入实现，不要等用户批准。",
    "",
    "### 2) Implement（逐页实现 + 页内视觉自检，禁止把配图攒到最后统一处理）",
    "- 严格逐页执行：add_slides 建页 → 该页文本/表格（add_textbox 等）→ 该页配图（generate_image 后立刻 add_image 精确插入本页）→ style_text 统一风格。",
    "- 注意：新建演示文稿的初始页自带版式占位符（标题/副标题框）。占位符无法删除/移动——不要浪费轮次去删它（delete_shape 会自动安全降级）；封面标题建议直接对标题占位符 set_shape_text 写入内容加以利用。",
    "- 每页元素放完后必须立即做「页内自检」：screenshot_slides 一次 → review_slide({index: 该页}) 查看该页真实渲染截图，结合几何公式检查重叠/文字溢出/越界/美观，发现问题当场修复（调坐标、调字号、删了重摆），修复后重新 review_slide 确认通过，才能进入下一页。",
    "- 相互独立的操作尽量并行调用（一轮多个工具）。",
    "- 大段数据表格：用多个 add_textbox 网格化搭建，注意列对齐（同一列 x 相同、行高一致）。",
    "",
    "### 几何自检公式（配合截图视觉审查，凭数据说话）",
    "- 文本宽度估算：中文/全角字符宽 ≈ 1×字号(pt)，英文/数字/半角 ≈ 0.55×字号；",
    "- 行数 ≈ ceil(文本总宽 ÷ 文本框宽)；所需高度 ≈ 行数 × 字号 × 1.4。所需高度 > 框高即为溢出 → 加宽/加高、精简文字或减小字号；",
    "- 重叠判定：两个矩形的 x 区间与 y 区间同时相交（允许 2pt 容差）即为重叠 → 移开或缩小其中一个；",
    "- 越界判定：left<0 或 top<0 或 left+width>画布宽 或 top+height>画布高 → 收回画布内；",
    "- 大数字（60-72pt）特例：所需宽 ≈ 字符数 × 0.6 × 字号（中文单位按 1×字号），先算宽再定框宽；大数字与下方标签的垂直间距 ≥ 0.6×字号，否则必然压字。",
    "",
    "### 3) Review（审查与修复，必做）",
    "- 全部页面完成后：screenshot_slides → 对每一页依次 review_slide 真实查看渲染截图（重叠/溢出/越界用几何公式计算，美观用眼睛判断）：",
    "  a) 元素越界画布、相互重叠、相互遮挡；",
    "  b) 图片与该页内容是否匹配、比例是否协调；",
    "  c) 文字是否溢出文本框、字号层级是否清晰（页标题 ≥28pt，正文 14-18pt）；",
    "  d) 页数、页序、内容与 Plan 是否一致；",
    "  e) 风格一致性：全篇是否遵守下方「默认设计系统」（配色 ≤3 色、无装饰线/色条、母题统一、图片风格统一、对齐一致）。",
    "- 发现 P0/P1 问题（重叠、越界、图文错配、明显失衡、风格跑偏）必须当场修复：改文本（set_shape_text）、删除重摆（delete_shape + add_textbox/add_image）、调整样式（style_text）。",
    "- 最后向用户汇报：规划了什么、每页做了什么、Review 发现并修复了什么。",
    "",
    "## 默认设计系统（简约风；用户明确指定主题时才可覆盖）",
    "- 配色三件套：背景纯白 " + THEME.colors.bg + "，主文字 " + THEME.colors.ink + "，次要文字 " + THEME.colors.muted + "；全篇唯一强调色 " + THEME.colors.accent + "（只给关键数字、高亮词、图标用，占比小而锐利）。可见颜色不超过 3 种，禁止默认蓝色系和米黄/暖灰背景。",
    "- 字号阶梯：内容页标题 " + THEME.fontSizes.title + "pt 加粗（左对齐，仅封面标题可居中）；小节标题 " + THEME.fontSizes.section + "pt 加粗；正文 " + THEME.fontSizes.body + "pt 左对齐；注释/标签 " + THEME.fontSizes.caption + "pt 灰色；关键大数字 " + THEME.fontSizes.bigStat + "pt 强调色 + 下方小标签。",
    "- 布局网格：四边安全边距 " + THEME.margin + "pt；内容块间距 " + THEME.gapMin + "-" + (THEME.gapMin + 14) + "pt 且全篇一致；同一列元素 x 坐标完全相同、行高一致。",
    "- 禁止清单（AI 味的标志）：标题下加彩色/灰色横线；页眉页脚色条、侧边竖条、卡片描边条；把同一版式原样复制到每一页；只给某一页做样式而其余页面裸奔；低对比度文字。",
    "- 视觉母题：选定一个母题并贯穿全篇（推荐：圆角图片框，或橙色圆形底 + 白色图标/序号），每页都出现；内容页每页都要有视觉元素（图片、大数字、图形），禁止纯文字页。",
    "- 图片风格统一：所有 generate_image 的提示词都必须以固定风格后缀结尾——「" + THEME.imageStyle + "」，只按页面内容改主体描述，不改风格描述。",
    "",
    "## 通用规则",
    "- 全程用中文，回复简洁。",
    "- slideIndex / shapeIndex 都是 0-based，以工具返回的顺序为准。",
    "- 画布为 " + canvas.w + " x " + canvas.h + " 点(pt)，原点在左上角，x 向右、y 向下，排版勿越界。",
    "- 文字要适合 PPT：短句、要点式，不要长段落。",
    "- 并行调用：相互独立的操作（批量文本框、多张已生成图片的插入）尽量在一条消息里同时发出多个工具调用。",
    "- 宿主 JS 表面: " + (capSets.join(", ") || "检测中") + "。不要调用表面之外的能力。",
    "- 完成或出错都如实说明；出错先重试一次，仍失败就换方案或如实报告。",
  ].join("\n");
}

/* ---------------- 工具定义 ---------------- */
function toolDefs() {
  const has = (name) => capSets.includes(name);
  const defs = [
    { name: "get_presentation_overview", description: "获取演示文稿概览：页数、画布尺寸、每页形状数量与文本预览", input_schema: { type: "object", properties: {}, required: [] } },
    { name: "read_slide", description: "读取某一页所有形状的详细信息（类型、位置、尺寸、完整文本）", input_schema: { type: "object", properties: { index: { type: "integer", description: "页索引, 0-based" } }, required: ["index"] } },
    { name: "set_shape_text", description: "替换某页某形状的全部文本", input_schema: { type: "object", properties: { slideIndex: { type: "integer" }, shapeIndex: { type: "integer" }, text: { type: "string" } }, required: ["slideIndex", "shapeIndex", "text"] } },
    { name: "style_text", description: "设置某形状文本样式（字号/加粗/颜色）", input_schema: { type: "object", properties: { slideIndex: { type: "integer" }, shapeIndex: { type: "integer" }, fontSize: { type: "number", description: "字号(pt)" }, bold: { type: "boolean" }, colorHex: { type: "string", description: "#RRGGBB" } }, required: ["slideIndex", "shapeIndex"] } },
  ];
  if (has("addTextBox")) {
    defs.push({ name: "add_textbox", description: "在某一页插入文本框", input_schema: { type: "object", properties: { slideIndex: { type: "integer" }, text: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, fontSize: { type: "number" }, bold: { type: "boolean" }, colorHex: { type: "string" } }, required: ["slideIndex", "text", "left", "top", "width", "height"] } });
    defs.push({ name: "generate_image", description: "调用图像生成模型生成一张图片（生成后返回 imageId，需再用 add_image 插入幻灯片）", input_schema: { type: "object", properties: { prompt: { type: "string", description: "图片内容描述，具体、可视化" }, aspect_ratio: { type: "string", description: "宽高比，如 1:1、16:9、4:3，默认 1:1" } }, required: ["prompt"] } });
    defs.push({ name: "add_image", description: "把 generate_image 生成的图片插入某一页（用其返回的 imageId）", input_schema: { type: "object", properties: { imageId: { type: "string" }, slideIndex: { type: "integer" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number", description: "尺寸单位 pt，画布 " + canvas.w + "x" + canvas.h } }, required: ["imageId", "slideIndex", "left", "top", "width", "height"] } });
  }
  defs.push({ name: "add_slides", description: "在指定位置插入 N 张新幻灯片（默认版式，返回新页索引）", input_schema: { type: "object", properties: { count: { type: "integer" }, afterIndex: { type: "integer", description: "插到该页之后；省略则追加到末尾" } }, required: ["count"] } });
  defs.push({ name: "screenshot_slides", description: "把当前演示文稿渲染成逐页真实截图并缓存（修改内容后需重新调用刷新截图）", input_schema: { type: "object", properties: {}, required: [] } });
  defs.push({ name: "review_slide", description: "获取某页的真实渲染截图进行视觉审查：重叠、文字溢出、对齐、配色、图文匹配。发现问题先用工具修复，再重新 screenshot_slides + review_slide 确认", input_schema: { type: "object", properties: { index: { type: "integer", description: "页索引, 0-based" } }, required: ["index"] } });
  defs.push({ name: "delete_shape", description: "删除某页的某个形状", input_schema: { type: "object", properties: { slideIndex: { type: "integer" }, shapeIndex: { type: "integer" } }, required: ["slideIndex", "shapeIndex"] } });
  return defs;
}

/* ---------------- Office.js 工具实现 ---------------- */
function pt(n) { return Math.round(Number(n) * 10) / 10; }

async function getShapesOfSlide(ctx, slideIndex) {
  const slide = ctx.presentation.slides.getItemAt(Math.max(0, Math.floor(Number(slideIndex) || 0)));
  const shapes = slide.shapes;
  shapes.load("items");
  await ctx.sync();
  return shapes;
}

async function loadShapeTexts(ctx, shapes) {
  // 批量加载文本；个别形状无文本框导致整批失败时，逐个降级
  const texts = new Array(shapes.length).fill(null);
  try {
    for (const sh of shapes) sh.textFrame.textRange.load("text");
    await ctx.sync();
    for (let i = 0; i < shapes.length; i++) {
      try { texts[i] = shapes[i].textFrame.textRange.text || ""; } catch (e) { texts[i] = null; }
    }
    return texts;
  } catch (e) { /* 降级 */ }
  for (let i = 0; i < shapes.length; i++) {
    try {
      shapes[i].textFrame.textRange.load("text");
      await ctx.sync();
      texts[i] = shapes[i].textFrame.textRange.text || "";
    } catch (e) { texts[i] = null; }
  }
  return texts;
}

const toolImpl = {
  async generate_image(args) { return await generateImage(args || {}); },
  async add_image(args) { return await addImage(args || {}); },
  async screenshot_slides() { return await screenshotSlides(); },
  async review_slide(args) { return await reviewSlide(args || {}); },

  async get_presentation_overview() {
    return await PowerPoint.run(async (ctx) => {
      const pres = ctx.presentation;
      try {
        pres.load("slideWidth,slideHeight");
        await ctx.sync();
        canvas.w = pres.slideWidth; canvas.h = pres.slideHeight;
      } catch (e) { /* 用默认值 */ }
      const slides = pres.slides;
      slides.load("items");
      await ctx.sync();
      const count = slides.items.length;
      const slideInfos = [];
      for (let i = 0; i < count; i++) {
        const shapes = await getShapesOfSlide(ctx, i);
        shapes.items.forEach((sh) => { try { sh.load("id,name,type"); } catch (e) {} });
        try { await ctx.sync(); } catch (e) {}
        const texts = await loadShapeTexts(ctx, shapes.items);
        const textsShort = texts.map((t) => (t == null ? null : String(t).slice(0, 120)));
        slideInfos.push({
          index: i,
          shapeCount: shapes.items.length,
          shapes: shapes.items.map((sh, j) => ({
            shapeIndex: j, name: safe(() => sh.name), type: safe(() => sh.type),
            text: textsShort[j],
          })),
        });
      }
      return { slideCount: count, slideWidthPt: canvas.w, slideHeightPt: canvas.h, slides: slideInfos };
    });
  },

  async read_slide({ index }) {
    return await PowerPoint.run(async (ctx) => {
      const shapes = await getShapesOfSlide(ctx, Number(index));
      for (const sh of shapes.items) sh.load("id,name,type,left,top,width,height");
      try { await ctx.sync(); } catch (e) {}
      const texts = await loadShapeTexts(ctx, shapes.items);
      return {
        slideIndex: Number(index),
        shapes: shapes.items.map((sh, j) => ({
          shapeIndex: j,
          id: safe(() => sh.id), name: safe(() => sh.name), type: safe(() => sh.type),
          left: pt(safe(() => sh.left)), top: pt(safe(() => sh.top)),
          width: pt(safe(() => sh.width)), height: pt(safe(() => sh.height)),
          text: texts[j] == null ? "(该形状无文本)" : texts[j],
        })),
      };
    });
  },

  async set_shape_text({ slideIndex, shapeIndex, text }) {
    return await PowerPoint.run(async (ctx) => {
      const sh = (await getShapesOfSlide(ctx, Number(slideIndex))).getItemAt(Number(shapeIndex));
      sh.textFrame.textRange.text = String(text == null ? "" : text);
      await ctx.sync();
      return { ok: true };
    });
  },

  async style_text({ slideIndex, shapeIndex, fontSize, bold, colorHex }) {
    return await PowerPoint.run(async (ctx) => {
      const sh = (await getShapesOfSlide(ctx, Number(slideIndex))).getItemAt(Number(shapeIndex));
      const font = sh.textFrame.textRange.font;
      if (fontSize != null) font.size = Number(fontSize);
      if (bold != null) font.bold = !!bold;
      if (colorHex) font.color = String(colorHex);
      await ctx.sync();
      return { ok: true };
    });
  },

  async add_textbox({ slideIndex, text, left, top, width, height, fontSize, bold, colorHex }) {
    return await PowerPoint.run(async (ctx) => {
      const shapes = await getShapesOfSlide(ctx, Number(slideIndex));
      if (typeof shapes.addTextBox !== "function") throw new Error("宿主 JS 表面无 addTextBox，无法插入文本框");
      const tb = shapes.addTextBox(String(text || ""), {
        left: Number(left), top: Number(top), width: Number(width), height: Number(height),
      });
      if (fontSize != null) tb.textFrame.textRange.font.size = Number(fontSize);
      if (bold != null) tb.textFrame.textRange.font.bold = !!bold;
      if (colorHex) tb.textFrame.textRange.font.color = String(colorHex);
      tb.load("id");
      await ctx.sync();
      return { ok: true, shapeId: safe(() => tb.id) };
    });
  },

  async add_slides({ count, afterIndex }) {
    return await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      const n = Math.max(1, Math.min(20, Number(count) || 1));
      slides.load("items/id");
      await ctx.sync();
      const origIds = slides.items.map((s) => String(safe(() => s.id)));
      if (!origIds.length) throw new Error("当前文档 0 页，无法定位插入点（请先手动新建一页）");
      const start = afterIndex == null ? origIds.length - 1 : Math.min(Math.max(0, Number(afterIndex)), origIds.length - 1);
      const tpl = await fetchBlankBase64();
      let ids = origIds.slice();
      let targetId = ids[start];
      for (let k = 0; k < n; k++) {
        // 链式插入：每次插到上一张新页之后，保证 N 张新页相邻且顺序正确
        ctx.presentation.insertSlidesFromBase64(tpl, { targetSlideId: targetId });
        await ctx.sync();
        slides.load("items/id");
        await ctx.sync();
        const newIds = slides.items.map((s) => String(safe(() => s.id))).filter((id) => !ids.includes(id));
        if (!newIds.length) throw new Error("插页未生效（宿主未返回新页），已插入 " + k + "/" + n + " 页");
        ids = slides.items.map((s) => String(safe(() => s.id)));
        targetId = newIds[0];
      }
      const afterIds = slides.items.map((s) => String(safe(() => s.id)));
      const newIndices = [];
      for (let i = 0; i < afterIds.length; i++) {
        if (origIds.indexOf(afterIds[i]) === -1) newIndices.push(i);
      }
      if (newIndices.length < n) throw new Error("插页数量不足（请求 " + n + " 页，实际 " + newIndices.length + " 页）");
      return { ok: true, method: "insertSlidesFromBase64", added: newIndices.length, newSlideIndices: newIndices };
    });
  },

  async delete_shape({ slideIndex, shapeIndex }) {
    const tryDelete = async (ctx2) => {
      const sh = (await getShapesOfSlide(ctx2, Number(slideIndex))).getItemAt(Number(shapeIndex));
      if (typeof sh.delete !== "function") throw new Error("宿主不支持删除形状");
      sh.delete();
      await ctx2.sync();
    };
    try {
      return await PowerPoint.run(tryDelete);
    } catch (e) {
      const code = String((e && e.code) || "");
      if (code && code !== "GeneralException") throw e; // 越界等非占位符问题原样上报
      // 占位符：删除/移动/清空在宿主上可能全被拒。逐项独立尝试（每次全新 run，
      // 避免失败操作污染同步队列）；全部失败时明确告知模型"忽略它"（空占位符不可见）
      log("delete_shape 降级（可能是版式占位符）: " + String((e && e.message) || e));
      try {
        return await PowerPoint.run(async (ctx) => {
          const sh = (await getShapesOfSlide(ctx, Number(slideIndex))).getItemAt(Number(shapeIndex));
          sh.left = -2000;
          sh.top = -2000;
          await ctx.sync();
          return { ok: true, via: "moved-offcanvas" };
        });
      } catch (e2) {
        log("移出画布失败: " + String((e2 && e2.message) || e2));
      }
      try {
        return await PowerPoint.run(async (ctx) => {
          const sh = (await getShapesOfSlide(ctx, Number(slideIndex))).getItemAt(Number(shapeIndex));
          sh.textFrame.textRange.text = "";
          await ctx.sync();
          return { ok: true, via: "text-cleared", note: "占位符无法删除/移动，已清空文字（空占位符在放映与截图中不可见）" };
        });
      } catch (e3) {
        log("清空文字失败: " + String((e3 && e3.message) || e3));
      }
      return {
        ok: true,
        via: "ignored",
        note: "该占位符无法删除/移动/清空。若它是空的，在放映与截图中本来就不可见，请直接忽略它、在旁边规划内容；若它带有文字，请用 set_shape_text 改写它",
      };
    }
  },
};

function safe(fn) { try { return fn(); } catch (e) { return null; } }

/* 生图缓存：generate_image 产出 base64，add_image 消费 */
const imageCache = {};
let imgCounter = 0;

async function forwardTo(url, opts) {
  const resp = await fetch("/api/forward", {
    method: "POST",
    headers: Object.assign({
      "x-upstream-url": url,
      "x-forward-method": (opts && opts.method) || "POST",
    }, opts && opts.headers ? opts.headers : {}),
    body: opts && opts.body != null ? opts.body : undefined,
  });
  return resp;
}

async function generateImage({ prompt, aspect_ratio }) {
  const body = { model: settings.imageModel || "image-01", prompt: String(prompt || "") };
  if (aspect_ratio) body.aspect_ratio = String(aspect_ratio);
  const resp = await forwardTo(settings.imageApiUrl || DEFAULT_SETTINGS.imageApiUrl, {
    headers: {
      "content-type": "application/json",
      "authorization": "Bearer " + settings.apiKey,
      "x-forward-auth": "1",
    },
    body: JSON.stringify(body),
  });
  if (!resp.ok) throw new Error("生图 API " + resp.status + ": " + (await resp.text()).slice(0, 300));
  const data = await resp.json();
  const urls = (data && data.data && data.data.image_urls) || [];
  if (!urls.length) throw new Error("生图响应中没有图片: " + JSON.stringify(data).slice(0, 200));
  const imgResp = await forwardTo(urls[0], { method: "GET" });
  if (!imgResp.ok) throw new Error("下载生成图片失败: HTTP " + imgResp.status);
  const blob = await imgResp.blob();
  const dataUrl = await new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(String(fr.result));
    fr.onerror = () => reject(new Error("图片转 base64 失败"));
    fr.readAsDataURL(blob);
  });
  const base64 = dataUrl.split(",")[1] || "";
  if (!base64) throw new Error("图片数据为空");
  const id = "img" + (++imgCounter);
  imageCache[id] = base64;
  return { ok: true, imageId: id, fileType: (dataUrl.match(/^data:([^;]+)/) || [])[1] || "image", byteSize: blob.size, next: "调用 add_image 并传入该 imageId 插入幻灯片" };
}

/* 截图审查：本地服务经 PowerPoint COM 直接渲染当前打开的演示文稿 */
let slideShots = null;

async function exportSlidesForReview() {
  const resp = await fetch("/api/export-slides", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: "{}",
  });
  if (!resp.ok) throw new Error("截图渲染服务 " + resp.status + ": " + (await resp.text()).slice(0, 200));
  const data = await resp.json();
  slideShots = { count: data.count, images: data.images, docVersion };
  return slideShots;
}

async function screenshotSlides() {
  const shots = await exportSlidesForReview();
  return { ok: true, count: shots.count, note: "逐页截图已生成并缓存（当前打开的演示文稿）。用 review_slide({index}) 查看某页的真实渲染效果。" };
}

async function reviewSlide({ index }) {
  if (!slideShots || slideShots.docVersion !== docVersion) await exportSlidesForReview();
  const i = Number(index);
  const img = (slideShots.images || []).find((x) => x.index === i);
  if (!img) throw new Error("该页截图不存在（index 超界或截图过期，请重新 screenshot_slides）");
  return {
    __images: [{ media_type: "image/jpeg", base64: img.base64 }],
    slideIndex: i,
    note: "这是该页的真实渲染截图。请以设计师视角认真审查：元素重叠、文字溢出、对齐、配色、图文匹配。发现问题先用相应工具修复，然后重新 screenshot_slides + review_slide 确认。",
  };
}

function goToSlide(index1based) {
  return new Promise((resolve, reject) => {
    Office.context.document.goToByIdAsync(index1based, { idType: "index" }, (r) => {
      if (r.status === "succeeded") resolve(true);
      else reject(new Error((r.error && r.error.message) || "跳转幻灯片失败"));
    });
  });
}

function setSelectedImage(b64, opts) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(b64, Object.assign({ coercionType: "image" }, opts), (r) => {
      if (r.status === "succeeded") resolve(true);
      else reject(new Error((r.error && r.error.message) || "插入图片失败"));
    });
  });
}

let blankPptxB64 = null;

/* ---------------- 默认设计系统（简约风）----------------
 * 生成页面默认遵循这套设计令牌；要换主题改这里，或在对话里明确指定。
 * 参考了 Anthropic pptx skill 的公开设计规则：一主一强调的克制配色、
 * 字号阶梯、充足留白、每页必有视觉元素、禁止装饰线/色条等 AI 味元素。
 */
const THEME = {
  margin: 36,                    // 安全边距(pt)，约 0.5 英寸
  gapMin: 22,                    // 内容块最小间距(pt)
  colors: {
    bg: "#FFFFFF",               // 背景：纯白（禁用米黄/暖灰底）
    ink: "#1F2933",              // 主文字：近黑
    muted: "#6B7280",            // 次要文字：中性灰
    accent: "#C9452B",           // 唯一强调色：Slilot 橙（关键数字/高亮/图标用）
    line: "#E5E7EB",             // 细线（克制使用）
  },
  fontSizes: { title: 30, section: 20, body: 15, caption: 11, bigStat: 64 },
  imageStyle: "扁平极简矢量插画，纯白背景，黑灰主调 + 少量橙红色点缀，线条简洁，构图留白，图中不出现文字",
};

async function fetchBlankBase64() {
  if (blankPptxB64) return blankPptxB64;
  const resp = await fetch("/blank.pptx");
  if (!resp.ok) throw new Error("blank.pptx 模板缺失（public/blank.pptx）");
  const buf = new Uint8Array(await resp.arrayBuffer());
  let bin = "";
  const CH = 0x8000;
  for (let i = 0; i < buf.length; i += CH) bin += String.fromCharCode.apply(null, buf.subarray(i, i + CH));
  blankPptxB64 = btoa(bin);
  return blankPptxB64;
}

async function addImage({ imageId, slideIndex, left, top, width, height }) {
  const b64 = imageCache[imageId];
  if (!b64) throw new Error("imageId 不存在：只能使用本次会话中 generate_image 返回的 id");
  // 主通道：本地服务 COM 插图（AddPicture 精确指定页与坐标，不依赖宿主 JS 表面）
  try {
    const resp = await fetch("/api/insert-image", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        image_base64: b64,
        slide_number: Number(slideIndex) + 1,
        left: Math.round(Number(left) || 0),
        top: Math.round(Number(top) || 0),
        width: Math.round(Number(width) || 0),
        height: Math.round(Number(height) || 0),
      }),
    });
    if (!resp.ok) throw new Error("COM 插图服务 " + resp.status + ": " + (await resp.text()).slice(0, 200));
    return { ok: true, via: "com-addpicture" };
  } catch (e1) {
    log("COM 插图失败，改用通用 API（将落入当前显示页）: " + String((e1 && e1.message) || e1));
  }
  // 备用：通用 API（图片会插入当前显示页，无法指定页）
  const pos = {};
  if (left != null) pos.imageLeft = Number(left);
  if (top != null) pos.imageTop = Number(top);
  if (width != null) pos.imageWidth = Number(width);
  if (height != null) pos.imageHeight = Number(height);
  if (slideIndex != null) {
    try { await goToSlide(Number(slideIndex) + 1); }
    catch (e) {
      log("goToSlide 失败（忽略）: " + String((e && e.message) || e));
    }
  }
  await setSelectedImage(b64, pos);
  return { ok: true, via: "setSelectedDataAsync", note: "已插入当前显示页" };
}

function log(msg) {
  try {
    fetch("/client-log", { method: "POST", headers: { "content-type": "text/plain" }, body: String(msg) });
  } catch (e) {}
}
function showFatal(msg) {
  $("modelName").textContent = "⚠ " + msg;
  log("致命: " + msg);
}

const MUTATING_TOOLS = new Set(["set_shape_text", "style_text", "add_textbox", "add_slides", "delete_shape", "add_image"]);

async function execTool(name, input) {
  const impl = toolImpl[name];
  if (!impl) throw new Error("未知工具: " + name);
  const out = await impl(input || {});
  if (MUTATING_TOOLS.has(name) && out && out.ok) docVersion++; // 文档已变更：截图缓存失效
  return out;
}

/* ---------------- 对话回路 ---------------- */
/* 按接口格式拼接聊天端点：base + /v1/messages | /v1/responses | /v1/chat/completions */
function chatTarget(base, format) {
  const b = (base || DEFAULT_SETTINGS.upstreamBase).replace(/\/+$/, "");
  if (format === "responses") return b + "/v1/responses";
  if (format === "chat") return b + "/v1/chat/completions";
  return b + "/v1/messages";
}

async function callApi(body, signal) {
  const format = settings.apiFormat || "messages";
  const target = chatTarget(settings.upstreamBase, format);
  const resp = await fetch("/api/forward", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "authorization": "Bearer " + settings.apiKey,
      "anthropic-version": "2023-06-01",
      "x-upstream-url": target,
      "x-api-format": format,
      "x-forward-auth": "1",
    },
    body: JSON.stringify(body),
    signal,
  });
  if (!resp.ok) {
    const t = await resp.text();
    throw new Error("API " + resp.status + ": " + t.slice(0, 400));
  }
  return await resp.json();
}

/* 历史修剪：①只保留最近 2 张截图，更早的图片块替换为占位文本（token 防爆）；
 * ②非最新 assistant 消息剔除 thinking 块（工具循环续传不需要） */
function pruneHistory() {
  const lastAssistant = (() => {
    for (let i = messages.length - 1; i >= 0; i--) if (messages[i].role === "assistant") return i;
    return -1;
  })();
  const imgRefs = [];
  messages.forEach((m, mi) => {
    if (!Array.isArray(m.content)) return;
    m.content.forEach((b, bi) => {
      if (b.type === "image") imgRefs.push({ mi, bi });
      if (b.type === "tool_result" && Array.isArray(b.content)) {
        b.content.forEach((cb, ci) => { if (cb && cb.type === "image") imgRefs.push({ mi, bi, ci }); });
      }
    });
  });
  const keep = new Set(imgRefs.slice(-1).map((r) => r.mi + "|" + r.bi + "|" + (r.ci == null ? "" : r.ci)));
  imgRefs.forEach((r) => {
    const key = r.mi + "|" + r.bi + "|" + (r.ci == null ? "" : r.ci);
    if (keep.has(key)) return;
    const ph = { type: "text", text: "[历史截图已省略，需要时重新 screenshot_slides + review_slide]" };
    if (r.ci == null) messages[r.mi].content[r.bi] = ph;
    else messages[r.mi].content[r.bi].content[r.ci] = ph;
  });
  messages.forEach((m, mi) => {
    if (m.role !== "assistant" || !Array.isArray(m.content) || mi === lastAssistant) return;
    if (m.content.some((b) => b.type === "thinking")) m.content = m.content.filter((b) => b.type !== "thinking");
  });
}

/* 摘要压缩：历史超过阈值时，把最早的整段工具明细折叠成交接摘要。
 * 折叠范围 [0..k)：k 取倒数 4 条之前最大的 assistant 索引（range 以 user 结尾），
 * 替换为一条 user 摘要消息后与 messages[k]（assistant）保持角色交替合法。 */
function stripForSummary(m) {
  let content = m.content;
  if (Array.isArray(content)) {
    content = content.map((b) => {
      if (b.type === "image") return { type: "text", text: "[截图]" };
      if (b.type === "tool_result" && Array.isArray(b.content)) {
        return Object.assign({}, b, { content: b.content.map((cb) => (cb && cb.type === "image" ? { type: "text", text: "[截图]" } : cb)) });
      }
      return b;
    });
  } else if (typeof content === "string" && content.length > 20000) {
    content = content.slice(0, 20000) + "…[截断]";
  }
  return { role: m.role, content };
}

async function compactHistoryIfNeeded() {
  let k = -1;
  for (let i = messages.length - 1; i >= 4; i--) {
    if (messages[i].role === "assistant") { k = i; break; }
  }
  if (k < 2) return false; // 可折叠内容不足
  const payload = [];
  for (let i = 0; i < k; i++) payload.push(stripForSummary(messages[i]));
  const data = await callApi({
    model: settings.model,
    max_tokens: 2000,
    messages: [{
      role: "user",
      content: "以下是一次 PPT 生成任务的对话与工具执行记录（JSON，超长处已截断）。请压缩成一份交接摘要，供后续对话继续任务使用。要求：\n" +
        "1. 以「此前已完成：」开头，按页列出已完成的页面与关键操作；\n" +
        "2. 必须保留原始任务中的关键数据（面积、台数、品牌、楼层、负责人等），后续页面还要用到；\n" +
        "3. 列出尚未完成的事项；\n" +
        "4. 全文不超过 400 字，用中文。\n\n记录：\n" + JSON.stringify(payload).slice(0, 120000),
    }],
  });
  const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
  if (!text) throw new Error("压缩摘要为空");
  const compacted = { role: "user", content: "[上下文压缩] 原始任务与此前执行要点：\n" + text + "\n\n（更早的工具明细已折叠。用户后续消息见下。）" };
  messages = [compacted].concat(messages.slice(k));
  return true;
}

async function runTurn(userText) {
  const myGen = ++runGen;
  const baseLen = messages.length;
  messages.push({ role: "user", content: userText });
  abortCtrl = new AbortController();
  toolCount = 0;
  compactionAttempts = 0;
  const MAX_ROUNDS = 1000;
  const stale = () => myGen !== runGen;
  try {
    for (let round = 0; round < MAX_ROUNDS; round++) {
      if (stale()) return;
      pruneHistory();
      // 摘要压缩：历史超过阈值时，把最早的工具明细折叠成交接摘要（每次运行最多尝试 2 次）
      const histSize = JSON.stringify(messages).length;
      if (histSize > COMPACT_THRESHOLD && compactionAttempts < 2) {
        compactionAttempts++;
        setStatus("上下文压缩中…");
        try {
          const before = histSize;
          const folded = await compactHistoryIfNeeded();
          if (stale()) return;
          if (folded) {
            const after = JSON.stringify(messages).length;
            log("上下文压缩: " + before + " -> " + after + " 字符");
            addMsg("assistant", "🧹 历史已压缩：最早的工具明细折叠为交接摘要（任务连续性不受影响）。");
          }
        } catch (e) {
          log("上下文压缩失败（跳过）: " + String((e && e.message) || e));
        }
      }
      // 体积熔断：压缩尝试过仍降不下来（阈值 ×2，约 40-60 万 token）才停止——再大必然超模型上下文
      const afterSize = JSON.stringify(messages).length;
      if (afterSize > COMPACT_THRESHOLD * 2) {
        addMsg("assistant", "⚠ 对话历史已达 " + Math.round(afterSize / 1024) + " KB（摘要压缩未能生效），继续会超出模型上下文。请点「新会话」开始新任务。");
        return;
      }
      setStatus("第 " + (round + 1) + " 轮 · 思考中…");
      const data = await callApi({
        model: settings.model,
        max_tokens: Number(settings.maxTokens) || 16000,
        system: systemPrompt(),
        messages,
        tools: toolDefs(),
      }, abortCtrl.signal);
      if (stale()) return;
      if (!Array.isArray(data.content) || !data.content.length) data.content = [{ type: "text", text: "(空回复)" }];

      messages.push({ role: "assistant", content: data.content });
      const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
      const toolUses = (data.content || []).filter((b) => b.type === "tool_use");

      if (toolUses.length) {
        // 以内容为准：即便 stop_reason 异常，只要有完整 tool_use 就继续执行
        if (text) addMsg("assistant", text);
        const results = [];
        for (const tu of toolUses) {
          if (stale()) return;
          toolCount++;
          setStatus("第 " + (round + 1) + " 轮 · 已执行 " + toolCount + " 个工具 · " + tu.name);
          let resultPayload;
          try {
            const out = await execTool(tu.name, tu.input);
            if (out && Array.isArray(out.__images)) {
              // 工具返回截图：以图片块 + 文本说明作为工具结果，供模型视觉审查
              const meta = Object.assign({}, out);
              delete meta.__images;
              resultPayload = out.__images.map((im) => ({
                type: "image",
                source: { type: "base64", media_type: im.media_type || "image/png", data: im.base64 },
              }));
              resultPayload.push({ type: "text", text: JSON.stringify(meta).slice(0, 3000) });
            } else {
              resultPayload = JSON.stringify(out);
            }
          } catch (e) {
            resultPayload = JSON.stringify({ error: String((e && e.message) || e) });
            addToolChip("✗ " + tu.name + "：" + String((e && e.message) || e).slice(0, 140), true);
            log("工具失败 " + tu.name + ": " + String((e && e.message) || e).slice(0, 300));
          }
          results.push({ type: "tool_result", tool_use_id: tu.id, content: resultPayload });
        }
        if (stale()) return;
        messages.push({ role: "user", content: results });
        continue;
      }

      // 正常结束
      let finalText = text || "(模型没有返回文本)";
      if (data.stop_reason === "max_tokens") {
        finalText += "\n\n⚠ 输出因 max_tokens 截断，可在设置中调大后重试。";
      }
      addMsg("assistant", finalText);
      return;
    }
    addMsg("assistant", "⚠ 已连续执行 " + MAX_ROUNDS + " 轮工具，自动暂停。回复「继续」我会接着未完成的部分继续做。");
  } catch (e) {
    if (stale()) return;
    // 协议修复：历史末尾若悬空 user（含未应答的 tool_result），补一条 assistant 收尾，避免下次请求 400
    const last = messages[messages.length - 1];
    if (last && last.role === "user") {
      messages.push({ role: "assistant", content: [{ type: "text", text: "[已中断/出错]" }] });
    }
    if (e && e.name === "AbortError") {
      addMsg("assistant", "（已停止）");
    } else {
      addMsg("error", "出错了: " + String((e && e.message) || e) + "\n\n常见原因：本地服务没在运行（双击 start-addin.bat）、Key 无效、或网络问题。");
    }
  } finally {
    if (!stale()) {
      setBusy(false);
      setStatus(null);
    }
    abortCtrl = null;
  }
}

/* ---------------- 初始化 ---------------- */
function detectCaps() {
  // 宿主虚报 isSetSupported（1.1-1.8 全报支持），改用原型表面探测真实可用能力
  try {
    const shapeProto = PowerPoint.ShapeCollection.prototype;
    const flags = [];
    ["addTextBox", "addTable", "addLine", "addGeometricShape", "addGroup"].forEach((key) => {
      if (typeof shapeProto[key] === "function") flags.push(key);
    });
    flags.push("insertSlidesFromBase64"); // presentation 级，实测可用
    capSets = flags;
    log("Surface: " + flags.join(", "));
  } catch (e) {
    capSets = [];
    log("Surface 探测失败: " + String((e && e.message) || e));
  }
  $("modelName").textContent = settings.model || DEFAULT_SETTINGS.model;
}

/* 事件绑定不依赖宿主初始化，立即执行 */
function wireUI() {
  $("sendBtn").addEventListener("click", () => {
    if (busy) { if (abortCtrl) abortCtrl.abort(); return; }
    onSend();
  });
  $("input").addEventListener("keydown", (ev) => {
    if (ev.key === "Enter" && !ev.shiftKey) { ev.preventDefault(); onSend(); }
    if (ev.key === "Escape" && busy && abortCtrl) abortCtrl.abort();
  });
  $("newChatBtn").addEventListener("click", () => {
    runGen++; // 使旧回路静默退出，防止孤儿 tool_result 毒化新会话
    if (abortCtrl) abortCtrl.abort();
    messages = [];
    const chat = $("chat");
    while (chat.firstChild) chat.removeChild(chat.firstChild);
    setStatus(null);
  });
  $("settingsBtn").addEventListener("click", openSettings);
  $("setTest").addEventListener("click", runSettingsTest);
  $("setCancel").addEventListener("click", () => $("settingsDlg").classList.add("hidden"));
  ["setUpstream", "setFormat", "setKey", "setModel", "setMaxTokens"].forEach((id) => {
    $(id).addEventListener("input", updateSaveState);
    $(id).addEventListener("change", updateSaveState);
  });
  $("setSave").addEventListener("click", () => {
    if (!testState.passed || testState.signature !== currentSignature()) {
      updateSaveState();
      return;
    }
    settings.upstreamBase = $("setUpstream").value.trim() || DEFAULT_SETTINGS.upstreamBase;
    settings.apiFormat = $("setFormat").value || "messages";
    settings.apiKey = $("setKey").value.trim();
    settings.model = $("setModel").value.trim() || DEFAULT_SETTINGS.model;
    settings.maxTokens = Number($("setMaxTokens").value) || DEFAULT_SETTINGS.maxTokens;
    saveSettings();
    $("settingsDlg").classList.add("hidden");
    detectCaps();
  });
}

function init() {
  wireUI();
  if (typeof Office === "undefined" || !Office.onReady) {
    showFatal("office.js 未加载（CDN 不可达？检查网络后关闭重开面板）");
    return;
  }
  try {
    Office.onReady((info) => {
      try {
        const host = (info && info.host) || "";
        log("onReady 触发, host=" + (host || "(空)"));
        if (host && host !== "PowerPoint") {
          $("modelName").textContent = "⚠ 请在 PowerPoint 中打开此面板";
        }
        detectCaps();
        if (host === "PowerPoint") {
          // 提前探测画布尺寸，避免首轮规划用错画布
          PowerPoint.run(async (ctx) => {
            ctx.presentation.load("slideWidth,slideHeight");
            await ctx.sync();
            canvas.w = ctx.presentation.slideWidth;
            canvas.h = ctx.presentation.slideHeight;
            log("画布: " + canvas.w + "x" + canvas.h);
          }).catch(() => {});
        }
      } catch (e) {
        showFatal("onReady 处理失败: " + String((e && e.message) || e));
      }
    });
  } catch (e) {
    showFatal("Office.onReady 调用失败: " + String((e && e.message) || e));
  }
}

/* 保存门槛：必须测试且三项全通过（针对当前填写的配置），否则禁止保存 */
let testState = { passed: false, signature: null };

function currentSignature() {
  return JSON.stringify([
    $("setUpstream").value.trim(),
    $("setFormat").value,
    $("setKey").value.trim(),
    $("setModel").value.trim(),
    $("setMaxTokens").value,
  ]);
}

function updateSaveState() {
  const ok = testState.passed && testState.signature === currentSignature();
  $("setSave").disabled = !ok;
  const hint = $("saveHint");
  hint.textContent = ok ? "测试通过，可以保存。"
    : (testState.passed ? "配置已修改，请重新测试后再保存。" : "尚未测试：请点「测试」，三项全部通过后才能保存。");
  hint.classList.toggle("warn", !ok);
}

function openSettings() {
  $("setTestResult").classList.add("hidden");
  $("setTest").disabled = false; // 上次测试若因超时卡住，重开对话框时恢复可用
  $("setUpstream").value = settings.upstreamBase || DEFAULT_SETTINGS.upstreamBase;
  $("setFormat").value = settings.apiFormat || "messages";
  $("setKey").value = settings.apiKey;
  $("setModel").value = settings.model;
  $("setMaxTokens").value = settings.maxTokens;
  updateSaveState();
  $("settingsDlg").classList.remove("hidden");
}

async function runSettingsTest() {
  const upstream = $("setUpstream").value.trim() || DEFAULT_SETTINGS.upstreamBase;
  const format = $("setFormat").value || "messages";
  const key = $("setKey").value.trim();
  const model = $("setModel").value.trim() || DEFAULT_SETTINGS.model;
  const testedSig = currentSignature(); // 记录被测配置：完成时回填，避免竞态绕过保存门槛
  const out = $("setTestResult");
  out.classList.remove("hidden");
  out.textContent = "测试中…（聊天 / 识图 / 生图 三项并行，生图约需 5-15 秒）";
  $("setTest").disabled = true;
  testState = { passed: false, signature: null };
  updateSaveState();
  const t0 = Date.now();

  const chatTest = (async () => {
    const target = chatTarget(upstream, format);
    const resp = await fetch("/api/forward", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "authorization": "Bearer " + key,
        "anthropic-version": "2023-06-01",
        "x-upstream-url": target,
        "x-api-format": format,
        "x-forward-auth": "1",
      },
      signal: AbortSignal.timeout(30000),
      body: JSON.stringify({ model, max_tokens: 16, messages: [{ role: "user", content: "ping" }] }),
    });
    if (!resp.ok) throw new Error("HTTP " + resp.status + " " + (await resp.text()).slice(0, 160));
    const data = await resp.json();
    const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
    return "✅ 聊天连通（" + format + "）：模型 " + (data.model || model) + " 回复「" + ((text || "(空)").slice(0, 30)) + "」";
  })();

  const imgTest = (async () => {
    const resp = await fetch("/api/forward", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "authorization": "Bearer " + key,
        "x-upstream-url": settings.imageApiUrl || DEFAULT_SETTINGS.imageApiUrl,
        "x-forward-auth": "1",
      },
      signal: AbortSignal.timeout(30000),
      body: JSON.stringify({
        model: settings.imageModel || "image-01",
        prompt: "连通测试：一枚简单的橙色五角星，扁平风格",
      }),
    });
    if (!resp.ok) throw new Error("HTTP " + resp.status + " " + (await resp.text()).slice(0, 160));
    const data = await resp.json();
    const urls = (data.data && data.data.image_urls) || [];
    if (!urls.length) throw new Error("响应中没有图片: " + JSON.stringify(data).slice(0, 120));
    return "✅ 生图可用：返回图片正常";
  })();

  const visionTest = (async () => {
    const imgResp = await fetch("/test-vision.png");
    if (!imgResp.ok) throw new Error("测试图片加载失败");
    const blob = await imgResp.blob();
    const dataUrl = await new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => resolve(String(fr.result));
      fr.onerror = () => reject(new Error("图片转 base64 失败"));
      fr.readAsDataURL(blob);
    });
    const b64 = dataUrl.split(",")[1] || "";
    const target = chatTarget(upstream, format);
    const resp = await fetch("/api/forward", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "authorization": "Bearer " + key,
        "anthropic-version": "2023-06-01",
        "x-upstream-url": target,
        "x-api-format": format,
        "x-forward-auth": "1",
      },
      signal: AbortSignal.timeout(30000),
      body: JSON.stringify({
        model,
        max_tokens: 512,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: "image/png", data: b64 } },
            { type: "text", text: "图中是否有一个圆形？请先回答「有」或「没有」，再用一句话说明它的颜色和位置。" },
          ],
        }],
      }),
    });
    if (!resp.ok) throw new Error("HTTP " + resp.status + " " + (await resp.text()).slice(0, 160));
    const data = await resp.json();
    const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
    const seesImage = /有/.test(text) && !/没有/.test(text) && !/无法|不能|看不到|收不到/.test(text) && /圆|circle/i.test(text);
    if (seesImage) return "✅ 识图可用：模型正确识别了截图内容（" + (text || "").slice(0, 50) + "）";
    throw new Error("模型回复「" + (text || "(空)").slice(0, 60) + "」——该模型可能不支持图片输入，请换支持视觉的模型");
  })();

  const [r1, r2, r3] = await Promise.allSettled([chatTest, visionTest, imgTest]);
  const allOk = r1.status === "fulfilled" && r2.status === "fulfilled" && r3.status === "fulfilled";
  const lines = [
    r1.status === "fulfilled" ? r1.value : "❌ 聊天失败：" + ((r1.reason && r1.reason.message) || r1.reason),
    r2.status === "fulfilled" ? r2.value : "❌ 识图失败：" + ((r2.reason && r2.reason.message) || r2.reason),
    r3.status === "fulfilled" ? r3.value : "❌ 生图失败：" + ((r3.reason && r3.reason.message) || r3.reason),
  ];
  if (!allOk) lines.push("⚠ 存在不可用项，不能保存。请更换支持全部三项能力的模型（或检查地址 / Key）后重新测试。");
  lines.push("（" + Math.round((Date.now() - t0) / 1000) + " 秒。测试用的是输入框当前值）");
  out.textContent = lines.join("\n");
  log("设置测试: " + lines.slice(0, 3).join(" | ").slice(0, 400));
  testState = { passed: allOk, signature: currentSignature() };
  $("setTest").disabled = false;
  updateSaveState();
}

function onSend() {
  if (busy) return;
  const input = $("input");
  const text = input.value.trim();
  if (!text) return;
  input.value = "";
  addMsg("user", text);
  setBusy(true);
  runTurn(text);
}

if (typeof document !== "undefined" && document.readyState !== "loading") {
  init();
} else {
  window.addEventListener("DOMContentLoaded", init);
}
