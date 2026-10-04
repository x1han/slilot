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
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
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
    "### 2) Implement（逐页实现，禁止把配图攒到最后统一处理）",
    "- 严格逐页执行：add_slides 建页 → 该页文本/表格（add_textbox 等）→ 该页配图（generate_image 后立刻 add_image）→ style_text 统一风格。",
    "- 一页完全做完再做下一页；相互独立的操作尽量并行调用（一轮多个工具）。",
    "- 大段数据表格：用多个 add_textbox 网格化搭建，注意列对齐（同一列 x 相同、行高一致）。",
    "",
    "### 3) Review（审查与修复，必做）",
    "- 全部页面完成后，用 get_presentation_overview + read_slide 逐页复查：",
    "  a) 元素越界画布、相互重叠、相互遮挡；",
    "  b) 图片与该页内容是否匹配、比例是否协调；",
    "  c) 文字是否溢出文本框、字号层级是否清晰（页标题 ≥28pt，正文 14-18pt）；",
    "  d) 页数、页序、内容与 Plan 是否一致。",
    "- 发现 P0/P1 问题（重叠、越界、图文错配、明显失衡）必须当场修复：改文本（set_shape_text）、删除重摆（delete_shape + add_textbox/add_image）、调整样式（style_text）。",
    "- 最后向用户汇报：规划了什么、每页做了什么、Review 发现并修复了什么。",
    "",
    "## 通用规则",
    "- 全程用中文，回复简洁。",
    "- slideIndex / shapeIndex 都是 0-based，以工具返回的顺序为准。",
    "- 画布为 " + canvas.w + " x " + canvas.h + " 点(pt)，原点在左上角，x 向右、y 向下，排版勿越界。",
    "- 文字要适合 PPT：短句、要点式，不要长段落。",
    "- 并行调用：相互独立的操作（批量文本框、多张已生成图片的插入）尽量在一条消息里同时发出多个工具调用。",
    "- 宿主支持的 API 集: " + (capSets.join(", ") || "检测中") + "。不要调用不可用工具。",
    "- 完成或出错都如实说明；出错先重试一次，仍失败就换方案或如实报告。",
  ].join("\n");
}

/* ---------------- 工具定义 ---------------- */
function toolDefs() {
  const has = (set) => capSets.includes(set);
  const defs = [
    { name: "get_presentation_overview", description: "获取演示文稿概览：页数、画布尺寸、每页形状数量与文本预览", input_schema: { type: "object", properties: {}, required: [] } },
    { name: "read_slide", description: "读取某一页所有形状的详细信息（类型、位置、尺寸、完整文本）", input_schema: { type: "object", properties: { index: { type: "integer", description: "页索引, 0-based" } }, required: ["index"] } },
  ];
  if (has("PowerPointApi 1.3")) {
    defs.push({ name: "set_shape_text", description: "替换某页某形状的全部文本", input_schema: { type: "object", properties: { slideIndex: { type: "integer" }, shapeIndex: { type: "integer" }, text: { type: "string" } }, required: ["slideIndex", "shapeIndex", "text"] } });
    defs.push({ name: "style_text", description: "设置某形状文本样式（字号/加粗/颜色）", input_schema: { type: "object", properties: { slideIndex: { type: "integer" }, shapeIndex: { type: "integer" }, fontSize: { type: "number", description: "字号(pt)" }, bold: { type: "boolean" }, colorHex: { type: "string", description: "#RRGGBB" } }, required: ["slideIndex", "shapeIndex"] } });
  }
  if (has("PowerPointApi 1.4")) {
    defs.push({ name: "add_textbox", description: "在某一页插入文本框", input_schema: { type: "object", properties: { slideIndex: { type: "integer" }, text: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, fontSize: { type: "number" }, bold: { type: "boolean" }, colorHex: { type: "string" } }, required: ["slideIndex", "text", "left", "top", "width", "height"] } });
    defs.push({ name: "generate_image", description: "调用图像生成模型生成一张图片（生成后返回 imageId，需再用 add_image 插入幻灯片）", input_schema: { type: "object", properties: { prompt: { type: "string", description: "图片内容描述，具体、可视化" }, aspect_ratio: { type: "string", description: "宽高比，如 1:1、16:9、4:3，默认 1:1" } }, required: ["prompt"] } });
    defs.push({ name: "add_image", description: "把 generate_image 生成的图片插入某一页（用其返回的 imageId）", input_schema: { type: "object", properties: { imageId: { type: "string" }, slideIndex: { type: "integer" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, description: "尺寸单位为 pt，画布 " + canvas.w + "x" + canvas.h }, required: ["imageId", "slideIndex", "left", "top", "width", "height"] } });
  }
  defs.push({ name: "add_slides", description: "在指定位置插入 N 张新幻灯片（默认版式，返回新页索引）", input_schema: { type: "object", properties: { count: { type: "integer" }, afterIndex: { type: "integer", description: "插到该页之后；省略则追加到末尾" } }, required: ["count"] } });
  defs.push({ name: "delete_shape", description: "删除某页的某个形状", input_schema: { type: "object", properties: { slideIndex: { type: "integer" }, shapeIndex: { type: "integer" } }, required: ["slideIndex", "shapeIndex"] } });
  return defs;
}

/* ---------------- Office.js 工具实现 ---------------- */
function pt(n) { return Math.round(Number(n) * 10) / 10; }

async function getShapesOfSlide(ctx, slideIndex) {
  const slide = ctx.presentation.slides.getItemAt(slideIndex);
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
        const metas = shapes.items.map((sh) => { sh.load("id,name,type"); return sh; });
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
    if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.4")) {
      throw new Error("宿主 PowerPointApi < 1.4，无法插入文本框");
    }
    return await PowerPoint.run(async (ctx) => {
      const shapes = await getShapesOfSlide(ctx, Number(slideIndex));
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
      const start = afterIndex == null ? null : Number(afterIndex);
      slides.load("items/id");
      await ctx.sync();
      const beforeIds = slides.items.map((s) => String(safe(() => s.id)));
      let usedInsert = typeof slides.add !== "function";

      if (!usedInsert) {
        // 主通道：slides.add（PowerPointApi 1.5+）
        try {
          for (let k = 0; k < n; k++) {
            const s = start == null ? slides.add() : slides.add({ index: start + 1 + k });
            s.load("id");
            await ctx.sync();
          }
        } catch (e) {
          log("slides.add 失败，改用 insertSlidesFromBase64: " + String((e && e.message) || e));
          usedInsert = true;
        }
      }
      if (usedInsert) {
        // 备用通道：插入空白模板页（insertSlidesFromBase64，PowerPointApi 1.3 起可用）
        const tpl = await fetchBlankBase64();
        for (let k = 0; k < n; k++) {
          const targetId = start == null ? beforeIds[beforeIds.length - 1] : beforeIds[Math.min(start + k, beforeIds.length - 1)];
          ctx.presentation.insertSlidesFromBase64(tpl, { targetSlideId: targetId });
          await ctx.sync();
        }
      }
      slides.load("items/id");
      await ctx.sync();
      const afterIds = slides.items.map((s) => String(safe(() => s.id)));
      const newIndices = [];
      for (let i = 0; i < afterIds.length; i++) {
        if (beforeIds.indexOf(afterIds[i]) === -1) newIndices.push(i);
      }
      return { ok: true, method: usedInsert ? "insertSlidesFromBase64" : "slides.add", added: newIndices.length || n, newSlideIndices: newIndices };
    });
  },

  async delete_shape({ slideIndex, shapeIndex }) {
    return await PowerPoint.run(async (ctx) => {
      const sh = (await getShapesOfSlide(ctx, Number(slideIndex))).getItemAt(Number(shapeIndex));
      if (typeof sh.delete !== "function") throw new Error("宿主不支持删除形状");
      sh.delete();
      await ctx.sync();
      return { ok: true };
    });
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
  const pos = {};
  if (left != null) pos.imageLeft = Number(left);
  if (top != null) pos.imageTop = Number(top);
  if (width != null) pos.imageWidth = Number(width);
  if (height != null) pos.imageHeight = Number(height);
  // 主通道：通用 API（Office 2013 起全平台可用）——先跳到目标页，再按坐标插入
  let jumpFailed = false;
  if (slideIndex != null) {
    try { await goToSlide(Number(slideIndex) + 1); }
    catch (e) {
      // 跳页失败不致命：图片将插入当前显示页
      jumpFailed = true;
      log("goToSlide 失败（忽略）: " + String((e && e.message) || e));
    }
  }
  try {
    await setSelectedImage(b64, pos);
    return { ok: true, via: "setSelectedDataAsync", jumpFailed, note: jumpFailed ? "未能跳页，图片插在了当前显示页" : undefined };
  } catch (e1) {
    log("setSelectedDataAsync 插图失败，改用富 API addImage: " + String((e1 && e1.message) || e1));
  }
  // 备用：富 API（需要 PowerPointApi 1.4 且宿主 JS 完整）
  return await PowerPoint.run(async (ctx) => {
    const shapes = await getShapesOfSlide(ctx, Number(slideIndex));
    const opts = {};
    if (left != null) opts.left = Number(left);
    if (top != null) opts.top = Number(top);
    if (width != null) opts.width = Number(width);
    if (height != null) opts.height = Number(height);
    const sh = shapes.addImage(b64, opts);
    sh.load("id");
    await ctx.sync();
    return { ok: true, via: "addImage", shapeId: safe(() => sh.id) };
  });
}

function log(msg) {
  try {
    fetch("/client-log", { method: "POST", headers: { "content-type": "text/plain" }, body: String(msg) });
  } catch (e) {}
}
function showFatal(msg) {
  $("capBadge").textContent = "⚠ " + msg;
  log("致命: " + msg);
}

async function execTool(name, input) {
  const impl = toolImpl[name];
  if (!impl) throw new Error("未知工具: " + name);
  return await impl(input || {});
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

async function runTurn(userText) {
  messages.push({ role: "user", content: userText });
  abortCtrl = new AbortController();
  const MAX_ROUNDS = 1000;
  try {
    for (let round = 0; round < MAX_ROUNDS; round++) {
      setStatus("第 " + (round + 1) + " 轮 · 思考中…");
      const data = await callApi({
        model: settings.model,
        max_tokens: Number(settings.maxTokens) || 16000,
        system: systemPrompt(),
        messages,
        tools: toolDefs(),
      }, abortCtrl.signal);

      messages.push({ role: "assistant", content: data.content });
      const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");

      if (data.stop_reason === "tool_use") {
        if (text) addMsg("assistant", text);
        const toolUses = (data.content || []).filter((b) => b.type === "tool_use");
        const results = [];
        for (const tu of toolUses) {
          setStatus("第 " + (round + 1) + " 轮 · 执行: " + tu.name);
          const chip = addToolChip("🔧 " + tu.name + " " + JSON.stringify(tu.input || {}).slice(0, 160), false);
          let resultStr;
          try {
            const out = await execTool(tu.name, tu.input);
            resultStr = JSON.stringify(out);
            chip.textContent = "🔧 " + tu.name + " ✓";
          } catch (e) {
            resultStr = JSON.stringify({ error: String((e && e.message) || e) });
            chip.textContent = "🔧 " + tu.name + " ✗ " + String((e && e.message) || e).slice(0, 120);
            chip.classList.add("err");
            log("工具失败 " + tu.name + ": " + String((e && e.message) || e).slice(0, 300));
          }
          results.push({ type: "tool_result", tool_use_id: tu.id, content: resultStr.slice(0, 60000) });
        }
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
    addMsg("assistant", "⚠ 已连续执行 " + MAX_ROUNDS + " 轮工具，自动暂停以防空转。回复「继续」我会接着未完成的部分继续做。");
  } catch (e) {
    if (e && e.name === "AbortError") {
      addMsg("assistant", "（已停止）");
    } else {
      addMsg("error", "出错了: " + String((e && e.message) || e) + "\n\n常见原因：本地服务没在运行（双击 start-addin.bat）、Key 无效、或网络问题。");
    }
  } finally {
    setBusy(false);
    setStatus(null);
    abortCtrl = null;
  }
}

/* ---------------- 初始化 ---------------- */
function detectCaps() {
  try {
    const reqs = Office.context && Office.context.requirements;
    if (!reqs) throw new Error("无宿主能力上下文（需在 PowerPoint 中打开）");
    const sets = [];
    ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8"].forEach((v) => {
      if (reqs.isSetSupported("PowerPointApi", v)) sets.push("PowerPointApi " + v);
    });
    capSets = sets;
    $("capBadge").textContent = "宿主: PowerPoint ✓  |  能力: " + (sets.join(", ") || "基本") +
      "  |  模型: " + settings.model;
    log("能力检测: " + (sets.join(", ") || "基本"));
    // 黑匣子：记录实际可用的 JS 表面（方法原型），用于诊断"能力集已报但方法缺失"类问题
    try {
      const shapeKeys = Object.getOwnPropertyNames(PowerPoint.ShapeCollection.prototype).filter((k) => k[0] !== "_");
      const slideKeys = Object.getOwnPropertyNames(PowerPoint.SlideCollection.prototype).filter((k) => k[0] !== "_");
      log("Surface ShapeCollection: " + shapeKeys.join(","));
      log("Surface SlideCollection: " + slideKeys.join(","));
    } catch (e) {
      log("Surface 探测失败: " + String((e && e.message) || e));
    }
    try {
      const res = performance.getEntriesByType("resource").filter((r) => /\.js/i.test(r.name)).map((r) => r.name.split("/").pop());
      log("JS资源: " + res.join(", "));
    } catch (e) {}
  } catch (e) {
    $("capBadge").textContent = "能力检测失败: " + String((e && e.message) || e);
    log("能力检测失败: " + String((e && e.message) || e));
  }
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
    if (abortCtrl) abortCtrl.abort();
    messages = [];
    const chat = $("chat");
    while (chat.children.length > 1) chat.removeChild(chat.lastChild);
    setStatus(null);
  });
  Array.prototype.forEach.call(document.querySelectorAll(".chip"), (c) => {
    c.addEventListener("click", () => { $("input").value = c.textContent; $("input").focus(); });
  });
  $("settingsBtn").addEventListener("click", openSettings);
  $("setTest").addEventListener("click", runSettingsTest);
  $("setCancel").addEventListener("click", () => $("settingsDlg").classList.add("hidden"));
  $("setSave").addEventListener("click", () => {
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
          $("capBadge").textContent = "⚠ 请在 PowerPoint 中打开此面板（当前宿主: " + host + "）";
        }
        detectCaps();
      } catch (e) {
        showFatal("onReady 处理失败: " + String((e && e.message) || e));
      }
    });
  } catch (e) {
    showFatal("Office.onReady 调用失败: " + String((e && e.message) || e));
  }
}

function openSettings() {
  $("setTestResult").classList.add("hidden");
  $("setUpstream").value = settings.upstreamBase || DEFAULT_SETTINGS.upstreamBase;
  $("setFormat").value = settings.apiFormat || "messages";
  $("setKey").value = settings.apiKey;
  $("setModel").value = settings.model;
  $("setMaxTokens").value = settings.maxTokens;
  $("settingsDlg").classList.remove("hidden");
}

async function runSettingsTest() {
  const upstream = $("setUpstream").value.trim() || DEFAULT_SETTINGS.upstreamBase;
  const format = $("setFormat").value || "messages";
  const key = $("setKey").value.trim();
  const model = $("setModel").value.trim() || DEFAULT_SETTINGS.model;
  const out = $("setTestResult");
  out.classList.remove("hidden");
  out.textContent = "测试中…（生图约需 5-15 秒，两项并行）";
  $("setTest").disabled = true;
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

  const [r1, r2] = await Promise.allSettled([chatTest, imgTest]);
  const lines = [
    r1.status === "fulfilled" ? r1.value : "❌ 聊天失败：" + ((r1.reason && r1.reason.message) || r1.reason),
    r2.status === "fulfilled" ? r2.value : "❌ 生图失败：" + ((r2.reason && r2.reason.message) || r2.reason),
    "（" + Math.round((Date.now() - t0) / 1000) + " 秒。测试用的是输入框当前值，确认无误后点「保存」）",
  ];
  out.textContent = lines.join("\n");
  log("设置测试: " + lines.slice(0, 2).join(" | ").slice(0, 300));
  $("setTest").disabled = false;
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
