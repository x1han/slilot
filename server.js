// Slilot - 本地服务
// 1) 托管加载项静态页面 (public/)
// 2) 通用反向代理 /api/forward：
//    - x-upstream-url      目标地址（必须 https）
//    - x-api-format        messages（透传）| responses | chat（格式翻译）
//    - x-forward-auth=1    附带客户端的 Authorization / anthropic-version 头
//    同源代理规避 CORS；API key 由页面随请求携带，不落盘。
"use strict";

const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { Readable } = require("stream");
const { execFile } = require("child_process");

const PORT = 3010;
const ROOT = path.join(__dirname, "public");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".json": "application/json; charset=utf-8",
  ".ico": "image/x-icon",
  ".svg": "image/svg+xml",
};

/* ---------- 格式翻译层 ----------
 * 面板内部始终说 Anthropic 方言；按所选接口格式翻译请求与响应：
 *   messages  -> 原样透传
 *   responses -> OpenAI /v1/responses
 *   chat      -> OpenAI /v1/chat/completions
 */
function toolResultText(tr) {
  if (typeof tr.content === "string") return tr.content;
  if (Array.isArray(tr.content)) return tr.content.map((b) => (b && b.text) || "").join("\n");
  return JSON.stringify(tr.content);
}

function anthropicToOpenAIChat(body) {
  const msgs = [];
  if (body.system) {
    msgs.push({ role: "system", content: typeof body.system === "string" ? body.system : JSON.stringify(body.system) });
  }
  for (const m of body.messages || []) {
    if (m.role === "user") {
      if (typeof m.content === "string") { msgs.push({ role: "user", content: m.content }); continue; }
      const blocks = m.content || [];
      for (const tr of blocks.filter((b) => b.type === "tool_result")) {
        msgs.push({ role: "tool", tool_call_id: tr.tool_use_id, content: toolResultText(tr) || "(empty)" });
        const imgs = Array.isArray(tr.content) ? tr.content.filter((b) => b && b.type === "image") : [];
        if (imgs.length) {
          const parts = imgs.map((b) => ({ type: "image_url", image_url: { url: "data:" + ((b.source && b.source.media_type) || "image/png") + ";base64," + ((b.source && b.source.data) || "") } }));
          parts.push({ type: "text", text: "（上方工具返回的截图）" });
          msgs.push({ role: "user", content: parts });
        }
      }
      const parts = [];
      for (const b of blocks) {
        if (b.type === "text") parts.push({ type: "text", text: b.text });
        else if (b.type === "image") parts.push({ type: "image_url", image_url: { url: "data:" + ((b.source && b.source.media_type) || "image/png") + ";base64," + ((b.source && b.source.data) || "") } });
      }
      if (parts.length === 1 && parts[0].type === "text") msgs.push({ role: "user", content: parts[0].text });
      else if (parts.length) msgs.push({ role: "user", content: parts });
    } else if (m.role === "assistant") {
      if (typeof m.content === "string") { msgs.push({ role: "assistant", content: m.content }); continue; }
      const blocks = m.content || [];
      const text = blocks.filter((b) => b.type === "text").map((b) => b.text).join("\n");
      const toolCalls = blocks.filter((b) => b.type === "tool_use").map((b) => ({
        id: b.id,
        type: "function",
        function: { name: b.name, arguments: JSON.stringify(b.input || {}) },
      }));
      const am = { role: "assistant", content: text || (toolCalls.length ? null : "") };
      if (toolCalls.length) am.tool_calls = toolCalls;
      msgs.push(am);
    }
  }
  const out = { model: body.model, max_tokens: body.max_tokens, messages: msgs };
  if (body.tools && body.tools.length) {
    out.tools = body.tools.map((t) => ({
      type: "function",
      function: { name: t.name, description: t.description, parameters: t.input_schema },
    }));
  }
  return out;
}

function openAIChatToAnthropic(data) {
  const choice = (data.choices && data.choices[0]) || {};
  const msg = choice.message || {};
  const content = [];
  if (msg.content) content.push({ type: "text", text: String(msg.content) });
  for (const tc of msg.tool_calls || []) {
    let input = {};
    try { input = JSON.parse((tc.function && tc.function.arguments) || "{}"); } catch (e) { input = { _raw: tc.function && tc.function.arguments }; }
    content.push({ type: "tool_use", id: tc.id || "call_" + Math.random().toString(36).slice(2), name: tc.function && tc.function.name, input });
  }
  if (!content.length) content.push({ type: "text", text: "(empty response)" });
  return {
    id: data.id || "",
    role: "assistant",
    model: data.model || "",
    content,
    stop_reason: choice.finish_reason === "tool_calls" ? "tool_use" : (choice.finish_reason === "length" ? "max_tokens" : "end_turn"),
    usage: data.usage ? { input_tokens: data.usage.prompt_tokens || 0, output_tokens: data.usage.completion_tokens || 0 } : undefined,
  };
}

function anthropicToResponses(body) {
  const input = [];
  for (const m of body.messages || []) {
    if (m.role === "user") {
      if (typeof m.content === "string") { input.push({ role: "user", content: m.content }); continue; }
      const blocks = m.content || [];
      for (const tr of blocks.filter((b) => b.type === "tool_result")) {
        input.push({ type: "function_call_output", call_id: tr.tool_use_id, output: toolResultText(tr) || "(empty)" });
        const imgs = Array.isArray(tr.content) ? tr.content.filter((b) => b && b.type === "image") : [];
        if (imgs.length) {
          const parts = imgs.map((b) => ({ type: "input_image", image_url: "data:" + ((b.source && b.source.media_type) || "image/png") + ";base64," + ((b.source && b.source.data) || "") }));
          parts.push({ type: "input_text", text: "（上方工具返回的截图）" });
          input.push({ role: "user", content: parts });
        }
      }
      const parts = [];
      for (const b of blocks) {
        if (b.type === "text") parts.push({ type: "input_text", text: b.text });
        else if (b.type === "image") parts.push({ type: "input_image", image_url: "data:" + ((b.source && b.source.media_type) || "image/png") + ";base64," + ((b.source && b.source.data) || "") });
      }
      if (parts.length === 1 && parts[0].type === "input_text") input.push({ role: "user", content: parts[0].text });
      else if (parts.length) input.push({ role: "user", content: parts });
    } else if (m.role === "assistant") {
      if (typeof m.content === "string") { input.push({ role: "assistant", content: m.content }); continue; }
      const blocks = m.content || [];
      const text = blocks.filter((b) => b.type === "text").map((b) => b.text).join("\n");
      if (text) input.push({ role: "assistant", content: text });
      for (const b of blocks.filter((b) => b.type === "tool_use")) {
        input.push({ type: "function_call", call_id: b.id, name: b.name, arguments: JSON.stringify(b.input || {}) });
      }
    }
  }
  const out = { model: body.model, max_output_tokens: body.max_tokens, input, store: false };
  if (body.system) out.instructions = typeof body.system === "string" ? body.system : JSON.stringify(body.system);
  if (body.tools && body.tools.length) {
    out.tools = body.tools.map((t) => ({ type: "function", name: t.name, description: t.description, parameters: t.input_schema }));
  }
  return out;
}

function responsesToAnthropic(data) {
  const content = [];
  let hasCall = false;
  for (const item of data.output || []) {
    if (item.type === "message") {
      const text = (item.content || []).filter((c) => c.type === "output_text").map((c) => c.text).join("");
      if (text) content.push({ type: "text", text });
    } else if (item.type === "function_call") {
      hasCall = true;
      let input = {};
      try { input = JSON.parse(item.arguments || "{}"); } catch (e) { input = { _raw: item.arguments }; }
      content.push({ type: "tool_use", id: item.call_id, name: item.name, input });
    }
    // reasoning 等其他块忽略
  }
  if (!content.length) content.push({ type: "text", text: "(empty response)" });
  return {
    id: data.id || "",
    role: "assistant",
    model: data.model || "",
    content,
    stop_reason: hasCall ? "tool_use" : (data.status === "incomplete" ? "max_tokens" : "end_turn"),
    usage: data.usage ? { input_tokens: data.usage.input_tokens || 0, output_tokens: data.usage.output_tokens || 0 } : undefined,
  };
}

async function handleForward(req, res) {
  const target = req.headers["x-upstream-url"];
  if (!target || !/^https:\/\//i.test(target)) {
    res.writeHead(400, { "content-type": "application/json; charset=utf-8" });
    return res.end(JSON.stringify({ error: "缺少或非法的 x-upstream-url（必须 https）" }));
  }
  const format = (req.headers["x-api-format"] || "messages").toLowerCase();
  const method = (req.headers["x-forward-method"] || "POST").toUpperCase();
  const chunks = [];
  if (method !== "GET" && method !== "HEAD") {
    for await (const c of req) chunks.push(c);
  }
  let bodyBuffer = chunks.length ? Buffer.concat(chunks) : null;
  // 请求翻译
  if (bodyBuffer && format !== "messages") {
    try {
      const anth = JSON.parse(bodyBuffer.toString("utf8"));
      const translated = format === "chat" ? anthropicToOpenAIChat(anth) : anthropicToResponses(anth);
      bodyBuffer = Buffer.from(JSON.stringify(translated));
    } catch (e) {
      res.writeHead(400, { "content-type": "application/json; charset=utf-8" });
      return res.end(JSON.stringify({ error: "请求翻译失败: " + String((e && e.message) || e) }));
    }
  }
  const headers = {
    "accept": req.headers["accept"] || "*/*",
  };
  if (bodyBuffer) {
    headers["content-type"] = req.headers["content-type"] || "application/json";
  }
  // 鉴权头：客户端显式声明时才转发（避免干扰对象存储签名 URL 等）
  if (req.headers["x-forward-auth"] === "1") {
    headers["authorization"] = req.headers["authorization"] || "";
    if (req.headers["anthropic-version"]) headers["anthropic-version"] = req.headers["anthropic-version"];
  }
  const upstream = await fetch(target, {
    method,
    headers,
    body: bodyBuffer || undefined,
    signal: AbortSignal.timeout(300000), // 上游挂起时 5 分钟自动放弃，避免请求永久悬挂
  });
  // 响应翻译（仅成功且非 messages 格式；错误体原样透传便于排错）
  if (format !== "messages" && upstream.ok) {
    const text = await upstream.text();
    let out;
    try {
      const data = JSON.parse(text);
      out = format === "chat" ? openAIChatToAnthropic(data) : responsesToAnthropic(data);
    } catch (e) {
      out = { raw: text.slice(0, 500) };
    }
    res.writeHead(upstream.status, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
    return res.end(JSON.stringify(out));
  }
  res.writeHead(upstream.status, {
    "content-type": upstream.headers.get("content-type") || "application/octet-stream",
  });
  if (upstream.body) {
    // 上游断流 / 客户端取消时不允许拖垮整个服务
    const nodeStream = Readable.fromWeb(upstream.body);
    nodeStream.on("error", () => {
      serverLog("上游响应流中断: " + String(target).slice(0, 120));
      try { res.destroy(); } catch (e) {}
    });
    res.on("error", () => { try { nodeStream.destroy(); } catch (e) {} });
    nodeStream.pipe(res);
  } else {
    res.end();
  }
}

/* ---------- 服务端图片下载：生图 CDN 二进制不再经 WebView 的流式代理管道 ----------
 * 此前生图结果的 CDN 下载走 /api/forward 流式转发，中途断流时 res.destroy() 会让
 * WebView 报出无任何信息的 "Failed to fetch"。改为服务端整体下载后返回 base64，
 * 失败时有明确 JSON 错误，且小响应可被客户端安全重试。
 * TLS 说明：部分图片 CDN 的证书链在浏览器可用（AIA 自动补全）但 Node 校验失败，
 * 表现为 "fetch failed"。因此先按标准校验下载，失败才回退到跳过校验的 https.get——
 * 该通道只用于公开的生成图片，不含任何凭据，风险可接受。 */
function httpsGetLoose(urlStr, redirects) {
  return new Promise((resolve, reject) => {
    let u;
    try { u = new URL(urlStr); } catch (e) { return reject(new Error("非法 URL")); }
    if (u.protocol !== "https:") return reject(new Error("仅支持 https"));
    const req = https.get({
      hostname: u.hostname,
      path: u.pathname + u.search,
      headers: { "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Slilot/1.0" },
      rejectUnauthorized: false,
      timeout: 30000,
    }, (r) => {
      if (r.statusCode >= 300 && r.statusCode < 400 && r.headers.location && (redirects || 0) < 4) {
        r.resume();
        const next = new URL(r.headers.location, u).toString();
        r.on("end", () => httpsGetLoose(next, (redirects || 0) + 1).then(resolve, reject));
        return;
      }
      const chunks = [];
      r.on("data", (c) => chunks.push(c));
      r.on("end", () => resolve({ status: r.statusCode, buf: Buffer.concat(chunks), mime: r.headers["content-type"] || "image/png" }));
      r.on("error", reject);
    });
    req.on("timeout", () => req.destroy(new Error("下载超时")));
    req.on("error", reject);
  });
}

/* 从 PNG/JPEG 字节头解析像素尺寸（零依赖） */
function imageDims(buf) {
  try {
    if (buf.length > 24 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) {
      return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
    }
    if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) {
      let off = 2;
      while (off + 9 < buf.length) {
        if (buf[off] !== 0xff) { off++; continue; }
        const m = buf[off + 1];
        if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
          return { height: buf.readUInt16BE(off + 5), width: buf.readUInt16BE(off + 7) };
        }
        if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { off += 2; continue; }
        off += 2 + buf.readUInt16BE(off + 2);
      }
    }
  } catch (e) {}
  return null;
}

async function handleFetchImage(req, res) {
  const chunks = [];
  for await (const c of req) chunks.push(c);
  let body;
  try { body = JSON.parse(Buffer.concat(chunks).toString("utf8")); } catch (e) {}
  const url = body && body.url;
  if (!url || !/^https:\/\//i.test(url)) {
    res.writeHead(400, { "content-type": "application/json; charset=utf-8" });
    return res.end(JSON.stringify({ error: "缺少或非法的 url（必须 https）" }));
  }
  let buf = null, mime = null, firstCause = "";
  try {
    const upstream = await fetch(url, {
      headers: { "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Slilot/1.0" },
      signal: AbortSignal.timeout(60000),
    });
    if (!upstream.ok) throw new Error("HTTP " + upstream.status);
    buf = Buffer.from(await upstream.arrayBuffer());
    mime = upstream.headers.get("content-type");
  } catch (e1) {
    const cause = e1 && e1.cause ? String(e1.cause.code || e1.cause.message) : "";
    firstCause = cause || String((e1 && e1.message) || e1);
    serverLog("fetch-image 标准通道失败 " + String(url).slice(0, 120) + ": " + String((e1 && e1.message) || e1) + (cause ? " cause: " + cause : "") + "，尝试宽松 TLS 重试");
    try {
      const r2 = await httpsGetLoose(String(url));
      if (r2.status !== 200) throw new Error("HTTP " + r2.status);
      buf = r2.buf;
      mime = r2.mime;
    } catch (e2) {
      serverLog("fetch-image 宽松通道也失败 " + String(url).slice(0, 120) + ": " + String((e2 && e2.message) || e2));
      res.writeHead(502, { "content-type": "application/json; charset=utf-8" });
      return res.end(JSON.stringify({ error: "图片下载失败: " + String((e2 && e2.message) || e2) + "（首次原因: " + firstCause + "）" }));
    }
  }
  const dims = imageDims(buf);
  res.writeHead(200, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
  res.end(JSON.stringify({
    base64: buf.toString("base64"),
    mime: mime || "image/png",
    width: dims ? dims.width : undefined,
    height: dims ? dims.height : undefined,
  }));
}

/* ---------- 幻灯片截图渲染（PowerPoint COM，附着当前活动演示文稿）---------- */
let exportLock = Promise.resolve();

function runExportScript(outDir) {
  const script = path.join(__dirname, "scripts", "export-slides.ps1");
  return new Promise((resolve, reject) => {
    execFile(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script, "-OutDir", outDir],
      { timeout: 120000, windowsHide: true },
      (err, stdout, stderr) => {
        if (err) {
          return reject(new Error("COM 渲染失败: " + String(err.message || err).slice(0, 200) + " " + String(stdout || "") + String(stderr || "").slice(0, 300)));
        }
        resolve();
      }
    );
  });
}

async function doExportSlides() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "slilot-export-"));
  try {
    await runExportScript(tmp);
    // 不依赖 PS 侧 manifest（避免编码交接问题），直接按文件名枚举排序
    const files = fs.readdirSync(tmp)
      .filter((f) => /^slide-\d+\.(?:png|jpg)$/i.test(f))
      .sort((a, b) => (parseInt(a.match(/\d+/), 10) - parseInt(b.match(/\d+/), 10)))
      .map((f) => path.join(tmp, f));
    const images = files.map((p, i) => ({ index: i, base64: fs.readFileSync(p).toString("base64") }));
    return { ok: true, count: images.length, images };
  } finally {
    try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (e) {}
  }
}

let exportPending = 0;

async function handleExportSlides(req, res) {
  if (exportPending >= 2) {
    res.writeHead(429, { "content-type": "application/json; charset=utf-8" });
    return res.end(JSON.stringify({ error: "已有截图任务在执行，请稍后再试" }));
  }
  exportPending++;
  // 串行执行，避免多个 COM 实例互相干扰
  const run = exportLock.then(() => doExportSlides()).finally(() => { exportPending--; });
  exportLock = run.catch(() => {});
  const out = await run;
  res.writeHead(200, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
  res.end(JSON.stringify(out));
}

/* ---------- 图片插入（PowerPoint COM，精确页 + 坐标）---------- */
async function handleInsertImage(req, res) {
  const chunks = [];
  for await (const c of req) chunks.push(c);
  let body;
  try { body = JSON.parse(Buffer.concat(chunks).toString("utf8")); } catch (e) {}
  const imgB64 = body && body.image_base64;
  const slideNumber = Math.floor(Number(body && body.slide_number) || 0);
  if (!imgB64 || slideNumber < 1) {
    res.writeHead(400, { "content-type": "application/json; charset=utf-8" });
    return res.end(JSON.stringify({ error: "缺少 image_base64 或 slide_number 非法" }));
  }
  const mediaType = (body.media_type || "image/png").toLowerCase();
  const ext = mediaType.includes("jpeg") ? ".jpg" : ".png";
  const left = Number(body.left) || 0;
  const top = Number(body.top) || 0;
  const width = Number(body.width) || 0;
  const height = Number(body.height) || 0;

  const run = exportLock.then(() => new Promise((resolve, reject) => {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "slilot-img-"));
    const imagePath = path.join(tmp, "image" + ext);
    try {
      fs.writeFileSync(imagePath, Buffer.from(imgB64, "base64"));
      execFile(
        "powershell.exe",
        ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
         path.join(__dirname, "scripts", "insert-image.ps1"),
         "-ImagePath", imagePath, "-SlideNumber", String(slideNumber),
         "-Left", String(left), "-Top", String(top), "-Width", String(width), "-Height", String(height)],
        { timeout: 60000, windowsHide: true },
        (err, stdout, stderr) => {
          try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (e) {}
          if (err) {
            return reject(new Error(String((err.message || err)).slice(0, 200) + " " + String(stdout || "") + String(stderr || "").slice(0, 300)));
          }
          resolve(String(stdout || ""));
        }
      );
    } catch (e) {
      try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (e2) {}
      reject(e);
    }
  }));
  exportLock = run.catch(() => {});
  const out = await run;
  // 解析脚本输出的实际放置矩形（脚本按原图比例 contain 缩放，可能与请求框不同）
  const placed = { left, top, width: width || undefined, height: height || undefined };
  const m = /PLACED left=([\d.-]+) top=([\d.-]+) width=([\d.-]+) height=([\d.-]+) native=(\d+)x(\d+)px scale=([\d.]+)/.exec(out);
  if (m) {
    placed.left = Number(m[1]); placed.top = Number(m[2]);
    placed.width = Number(m[3]); placed.height = Number(m[4]);
    placed.native = m[5] + "x" + m[6];
    placed.scale = Number(m[7]);
  }
  res.writeHead(200, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
  res.end(JSON.stringify({ ok: true, placed }));
}

function handleStatic(req, res, pathname) {
  let p = decodeURIComponent(pathname);
  if (p === "/") p = "/taskpane.html";
  const file = path.normalize(path.join(ROOT, p));
  const rel = path.relative(ROOT, file);
  if (rel === "" || rel.startsWith("..") || path.isAbsolute(rel)) {
    res.writeHead(403);
    return res.end("forbidden");
  }
  fs.readFile(file, (err, data) => {
    if (err) {
      res.writeHead(404, { "content-type": "text/plain; charset=utf-8" });
      return res.end("404 not found: " + p);
    }
    res.writeHead(200, { "content-type": MIME[path.extname(file).toLowerCase()] || "application/octet-stream", "cache-control": "no-store" });
    res.end(data);
  });
}

const handler = async (req, res) => {
  const url = new URL(req.url, "http://localhost");
  try {
    // 跨站防护：浏览器发起的跨源请求一律 403（同源面板与本机无 Origin 的工具调用不受影响）
    const origin = req.headers.origin;
    if (origin && origin !== "https://localhost:3010") {
      res.writeHead(403, { "content-type": "application/json; charset=utf-8" });
      return res.end(JSON.stringify({ error: "cross-origin forbidden" }));
    }
    if (url.pathname === "/api/forward") {
      // 通用转发：目标地址由请求头 x-upstream-url 指定（仅允许 https）
      await handleForward(req, res);
    } else if (url.pathname === "/api/export-slides") {
      // 幻灯片截图：pptx_base64 -> 逐页 PNG base64（PowerPoint COM 渲染）
      await handleExportSlides(req, res);
    } else if (url.pathname === "/api/fetch-image") {
      // 服务端下载图片（生图结果的 CDN 地址），返回 base64 + mime
      await handleFetchImage(req, res);
    } else if (url.pathname === "/api/insert-image") {
      // 精确插图：指定页 + 坐标（PowerPoint COM AddPicture）
      await handleInsertImage(req, res);
    } else if (url.pathname === "/client-log") {
      // 面板"黑匣子"：接收页面 JS 错误与状态日志
      const chunks = [];
      for await (const c of req) chunks.push(c);
      const line = new Date().toISOString() + "  " + Buffer.concat(chunks).toString("utf8").slice(0, 2000) + "\n";
      fs.mkdirSync(path.join(__dirname, "logs"), { recursive: true });
      fs.appendFileSync(path.join(__dirname, "logs", "client-log.txt"), line);
      res.writeHead(204);
      res.end();
    } else if (url.pathname === "/healthz") {
      res.writeHead(200, { "content-type": "text/plain" });
      res.end("ok");
    } else {
      handleStatic(req, res, url.pathname);
    }
  } catch (e) {
    const cause = e && e.cause ? String(e.cause.code || e.cause.message) : "";
    serverLog("请求处理失败 " + req.url + ": " + String((e && e.message) || e) + (cause ? " cause: " + cause : ""));
    if (!res.headersSent) res.writeHead(502, { "content-type": "application/json; charset=utf-8" });
    try { res.end(JSON.stringify({ error: String((e && e.message) || e) + (cause ? "（" + cause + "）" : "") })); } catch (_) {}
  }
};

let server;
try {
  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  const tls = {
    key: fs.readFileSync(path.join(certDir, "localhost.key")),
    cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
  };
  server = https.createServer(tls, handler);
} catch (e) {
  console.error("[Slilot] 缺少 localhost 开发证书，无法启动 HTTPS。");
  console.error("请先运行: npx office-addin-dev-certs install");
  console.error(String((e && e.message) || e));
  process.exit(1);
}

server.listen(PORT, "127.0.0.1", () => {
  console.log(`[Slilot] 服务已启动: https://localhost:${PORT}`);
});

// Node 默认 5 秒就掐断空闲 keep-alive 连接，而 WebView 的连接池会复用更久的连接，
// 撞上掐断窗口就表现为偶发 "Failed to fetch"。拉长到 65 秒基本消除该竞态。
server.keepAliveTimeout = 65000;
server.headersTimeout = 66000;

server.on("error", (e) => {
  console.error("[Slilot] 启动失败:", e.message);
  process.exit(1);
});

// 兜底：任何未捕获异常只记日志，不允许拖垮本地服务（面板的所有请求都依赖它）
function serverLog(line) {
  try {
    fs.mkdirSync(path.join(__dirname, "logs"), { recursive: true });
    fs.appendFileSync(path.join(__dirname, "logs", "server.log"), new Date().toISOString() + "  " + line + "\n");
  } catch (e) {}
}
process.on("uncaughtException", (e) => serverLog("uncaughtException: " + String((e && e.stack) || e)));
process.on("unhandledRejection", (e) => serverLog("unhandledRejection: " + String((e && e.stack) || e)));
