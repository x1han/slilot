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
    Readable.fromWeb(upstream.body).pipe(res);
  } else {
    res.end();
  }
}

/* ---------- 幻灯片截图渲染（PowerPoint COM，附着当前活动演示文稿）---------- */
let exportLock = Promise.resolve();

function runExportScript(outDir) {
  const script = path.join(__dirname, "scripts", "export-slides.ps1");
  return new Promise((resolve, reject) => {
    execFile(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script, "-OutDir", outDir],
      { timeout: 120000 },
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
    const manifestTxt = fs.readFileSync(path.join(tmp, "manifest.txt"), "utf8");
    const files = manifestTxt.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
    const images = files.map((p, i) => ({ index: i, base64: fs.readFileSync(p).toString("base64") }));
    return { ok: true, count: images.length, images };
  } finally {
    try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (e) {}
  }
}

async function handleExportSlides(req, res) {
  // 串行执行，避免多个 COM 实例互相干扰
  const run = exportLock.then(() => doExportSlides());
  exportLock = run.catch(() => {});
  const out = await run;
  res.writeHead(200, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
  res.end(JSON.stringify(out));
}

function handleStatic(req, res, pathname) {
  let p = decodeURIComponent(pathname);
  if (p === "/") p = "/taskpane.html";
  const file = path.normalize(path.join(ROOT, p));
  if (!file.startsWith(ROOT)) {
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
    if (url.pathname === "/api/forward") {
      // 通用转发：目标地址由请求头 x-upstream-url 指定（仅允许 https）
      await handleForward(req, res);
    } else if (url.pathname === "/api/export-slides") {
      // 幻灯片截图：pptx_base64 -> 逐页 PNG base64（PowerPoint COM 渲染）
      await handleExportSlides(req, res);
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
    if (!res.headersSent) res.writeHead(502, { "content-type": "application/json; charset=utf-8" });
    try { res.end(JSON.stringify({ error: String((e && e.message) || e) })); } catch (_) {}
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

server.on("error", (e) => {
  console.error("[Slilot] 启动失败:", e.message);
  process.exit(1);
});
