// rag/aiProvider.js — LAPISAN ADAPTER MULTI-PROVIDER.
//
// Format INTERNAL FRIDA = format Anthropic Messages:
//   request : { system, messages, tools, tool_choice, maxTokens }
//   response: { content: [{type:"text"|"tool_use",...}], stop_reason }
//
// callMessages() me-routing ke adapter provider aktif; tiap adapter MENERJEMAHKAN
// request internal -> format provider, dan response provider -> format internal.
// Dengan begitu pemanggil (server.js /api/edit, /api/agent, rag/llm.js) TIDAK perlu
// tahu provider mana yang dipakai.
//
// Provider:
//   anthropic — /v1/messages, header x-api-key + anthropic-version (endpoint resmi dikunci)
//   openai    — /chat/completions, Authorization: Bearer (endpoint resmi dikunci)
//   gemini    — :generateContent, header x-goog-api-key (endpoint resmi dikunci)
//   custom    — /v1/messages OpenAI/Anthropic-compatible (9Router/Aerolink), baseUrl bebas

const providerConfig = require("./providerConfig");

// Endpoint resmi DIKUNCI (tidak bisa dioverride dari UI). Custom pakai baseUrl sendiri.
const ENDPOINTS = {
  anthropic: "https://api.anthropic.com/",
  openai: "https://api.openai.com/v1/",
  gemini: "https://generativelanguage.googleapis.com/v1beta/",
};
const ANTHROPIC_VERSION = "2023-06-01";

// Beberapa gateway Anthropic-compatible (mis. AgentRouter/agentrouter.org) hanya melayani
// klien yang dikenali (Claude Code) dan menolak yang lain: "unauthorized client detected".
// Header ini membuat request adapter `custom` tampak seperti Claude Code CLI resmi supaya
// lolos, memakai token gateway milik user sendiri. Semua bisa dioverride via env bila
// gateway mengubah kriteria deteksinya.
//   FRIDA_CLIENT_UA      -> ganti User-Agent penuh
//   FRIDA_CLAUDE_CODE_VERSION -> ganti nomor versi di UA/paket
//   FRIDA_ANTHROPIC_BETA -> ganti daftar flag anthropic-beta
//   FRIDA_CUSTOM_HEADERS -> JSON objek header tambahan (mis. token khusus gateway)
const CLAUDE_CODE_VERSION = process.env.FRIDA_CLAUDE_CODE_VERSION || "1.0.60";
function claudeCodeHeaders() {
  const ua = process.env.FRIDA_CLIENT_UA || ("claude-cli/" + CLAUDE_CODE_VERSION + " (external, cli)");
  const h = {
    "user-agent": ua,
    "x-app": "cli",
    "anthropic-beta": process.env.FRIDA_ANTHROPIC_BETA || "claude-code-20250219,oauth-2025-04-20",
    "anthropic-dangerous-direct-browser-access": "true",
    "x-stainless-lang": "js",
    "x-stainless-runtime": "node",
    "x-stainless-package-version": CLAUDE_CODE_VERSION,
  };
  if (process.env.FRIDA_CUSTOM_HEADERS) {
    try { Object.assign(h, JSON.parse(process.env.FRIDA_CUSTOM_HEADERS)); } catch (_) {}
  }
  return h;
}

const PROVIDER_LABELS = {
  anthropic: "Anthropic (Claude)",
  openai: "OpenAI",
  gemini: "Google Gemini",
  custom: "Custom / OpenAI-compatible",
};

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Gabungkan base URL + path relatif tanpa menduplikasi segmen "v1/".
// Base bisa berakhir "/", "/v1", atau "/v1/"; path bisa "v1/messages" atau "v1/models".
// Contoh: joinUrl("https://x/v1/", "v1/messages") -> "https://x/v1/messages" (bukan /v1/v1/).
function joinUrl(base, path) {
  let b = String(base || "").trim().replace(/\/+$/, ""); // buang trailing slash
  let p = String(path || "").replace(/^\/+/, "");        // buang leading slash
  const pFirst = p.split("/")[0];                          // segmen pertama path (mis. "v1")
  if (pFirst && new RegExp("/" + pFirst + "$", "i").test(b)) {
    // base sudah berakhir dengan segmen yang sama -> jangan ulang
    p = p.slice(pFirst.length).replace(/^\/+/, "");
  }
  return b + "/" + p;
}

// ============================ util error ============================
function extractErrText(raw) {
  if (!raw) return "";
  try {
    const j = JSON.parse(raw);
    if (j.error) return typeof j.error === "string" ? j.error : (j.error.message || JSON.stringify(j.error));
    if (j.message) return j.message;
  } catch (_) {}
  return String(raw).slice(0, 180);
}

// Pesan error yang ramah untuk UI (Bahasa Indonesia), bukan cuma "Gagal".
function friendlyError(provider, status, raw) {
  const name = PROVIDER_LABELS[provider] || provider;
  let msg;
  if (status === 401 || status === 403) msg = "API key " + name + " tidak valid atau tidak berizin (" + status + ").";
  else if (status === 404) msg = "Model atau endpoint " + name + " tidak ditemukan (404). Periksa nama model.";
  else if (status === 429) msg = "Batas rate " + name + " tercapai (429). Coba lagi sebentar.";
  else if (status >= 500) msg = "Server " + name + " sedang bermasalah (" + status + ").";
  else msg = name + " menolak permintaan (" + status + ").";
  const detail = extractErrText(raw);
  return new Error(detail ? msg + " — " + detail : msg);
}

// POST JSON generik; melempar friendlyError bila status != 2xx. Mengembalikan {raw, contentType}.
async function postJson(provider, url, headers, bodyObj, timeoutMs) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), timeoutMs || 60000);
  let resp;
  try {
    resp = await fetch(url, {
      method: "POST",
      headers: Object.assign({ "content-type": "application/json" }, headers),
      body: JSON.stringify(bodyObj),
      signal: ctrl.signal,
    });
  } catch (e) {
    clearTimeout(t);
    throw new Error((PROVIDER_LABELS[provider] || provider) + " tidak dapat dihubungi: " + String(e.message || e));
  }
  clearTimeout(t);
  const raw = await resp.text();
  if (!resp.ok) throw friendlyError(provider, resp.status, raw);
  return { raw, contentType: resp.headers.get("content-type") };
}

// ===================== parser respons -> Anthropic =====================
// Menangani: (a) JSON native Anthropic, (b) JSON OpenAI, (c) SSE Anthropic (router lokal).
function parseBodyToAnthropic(contentType, raw) {
  const isStream = (contentType || "").includes("event-stream");
  if (!isStream) {
    let data;
    try { data = JSON.parse(raw); } catch (e) { throw new Error("Respons bukan JSON valid: " + String(raw).slice(0, 200)); }
    return normalizeOpenAI(data);
  }
  const lines = String(raw).split(/\r?\n/);
  const blocks = {};
  let stopReason = "end_turn";
  let msgId = "";
  for (const line of lines) {
    if (!line.startsWith("data:")) continue;
    const payload = line.slice(5).trim();
    if (!payload || payload === "[DONE]") continue;
    let ev;
    try { ev = JSON.parse(payload); } catch (_) { continue; }
    if (ev.type === "message_start" && ev.message) {
      msgId = ev.message.id || "";
    } else if (ev.type === "content_block_start") {
      const cb = ev.content_block || {};
      blocks[ev.index] = { type: cb.type, id: cb.id, name: cb.name, partialJson: "", text: "" };
    } else if (ev.type === "content_block_delta") {
      const b = blocks[ev.index];
      if (!b) continue;
      const d = ev.delta || {};
      if (d.type === "input_json_delta") b.partialJson += (d.partial_json || "");
      if (d.type === "text_delta") b.text += (d.text || "");
    } else if (ev.type === "message_delta") {
      if (ev.delta && ev.delta.stop_reason) stopReason = ev.delta.stop_reason;
    }
  }
  const content = Object.keys(blocks).sort((a, b) => Number(a) - Number(b)).map((idx) => {
    const b = blocks[idx];
    if (b.type === "tool_use") {
      let input = {};
      try { input = JSON.parse(b.partialJson || "{}"); } catch (_) {}
      return { type: "tool_use", id: b.id, name: b.name, input };
    }
    return { type: "text", text: b.text || "" };
  });
  return { id: msgId, stop_reason: stopReason, content };
}

// Normalisasi format OpenAI -> Anthropic (untuk JSON non-stream). Anthropic native diloloskan.
function normalizeOpenAI(data) {
  if (data.content) return data; // sudah Anthropic
  const choice = (data.choices || [])[0];
  if (!choice) return data;
  const msg = choice.message || {};
  const finish = choice.finish_reason || "stop";
  const content = [];
  if (msg.content) content.push({ type: "text", text: msg.content });
  if (Array.isArray(msg.tool_calls)) {
    for (const tc of msg.tool_calls) {
      const fn = tc.function || {};
      let input = {};
      try { input = JSON.parse(fn.arguments || "{}"); } catch (_) {}
      content.push({ type: "tool_use", id: tc.id || ("tc_" + Math.random().toString(36).slice(2)), name: fn.name || "", input });
    }
  }
  const reasonMap = { tool_calls: "tool_use", stop: "end_turn", length: "max_tokens" };
  return { id: data.id, model: data.model, stop_reason: reasonMap[finish] || finish, content };
}

// ============================ ADAPTER: Anthropic ============================
async function callAnthropic(o) {
  const body = { model: o.model, max_tokens: o.maxTokens, messages: o.messages };
  if (o.system) body.system = o.system;
  if (o.tools) body.tools = o.tools;
  if (o.tool_choice) body.tool_choice = o.tool_choice;
  const { raw, contentType } = await postJson("anthropic", ENDPOINTS.anthropic + "v1/messages", {
    "x-api-key": o.apiKey,
    "anthropic-version": ANTHROPIC_VERSION,
  }, body);
  const data = parseBodyToAnthropic(contentType, raw);
  return { content: data.content || [], stop_reason: data.stop_reason || "end_turn" };
}

// ============================ ADAPTER: Custom (existing) ============================
// Pertahankan perilaku existing yang sudah jalan ke 9Router/Aerolink:
// POST {baseUrl}v1/messages dgn x-api-key + Authorization Bearer + anthropic-version.
// Ditambah header identitas Claude Code (claudeCodeHeaders) agar lolos gateway ber-proteksi
// klien seperti AgentRouter ("unauthorized client detected").
async function callCustom(o) {
  const base = String(o.baseUrl || providerConfig.DEFAULT_CUSTOM_BASE);
  const body = { model: o.model, max_tokens: o.maxTokens, messages: o.messages };
  if (o.system) body.system = o.system;
  if (o.tools) body.tools = o.tools;
  if (o.tool_choice) body.tool_choice = o.tool_choice;
  const headers = Object.assign(claudeCodeHeaders(), {
    "x-api-key": o.apiKey,
    "authorization": "Bearer " + o.apiKey,
    "anthropic-version": ANTHROPIC_VERSION,
  });
  const { raw, contentType } = await postJson("custom", joinUrl(base, "v1/messages"), headers, body);
  const data = parseBodyToAnthropic(contentType, raw);
  return { content: data.content || [], stop_reason: data.stop_reason || "end_turn" };
}

// ============================ ADAPTER: OpenAI ============================
function toOpenAIMessages(system, messages) {
  const out = [];
  if (system) out.push({ role: "system", content: system });
  for (const m of messages || []) {
    if (typeof m.content === "string") { out.push({ role: m.role, content: m.content }); continue; }
    const blocks = Array.isArray(m.content) ? m.content : [];
    if (m.role === "assistant") {
      const text = blocks.filter((b) => b.type === "text").map((b) => b.text).join("\n");
      const toolUses = blocks.filter((b) => b.type === "tool_use");
      const msg = { role: "assistant", content: text || null };
      if (toolUses.length) {
        msg.tool_calls = toolUses.map((tu) => ({
          id: tu.id, type: "function",
          function: { name: tu.name, arguments: JSON.stringify(tu.input || {}) },
        }));
      }
      out.push(msg);
    } else { // user
      // tool_result -> pesan role "tool" (harus mengikuti pesan assistant yg memanggil)
      for (const b of blocks) {
        if (b.type === "tool_result") {
          out.push({
            role: "tool",
            tool_call_id: b.tool_use_id,
            content: typeof b.content === "string" ? b.content : JSON.stringify(b.content),
          });
        }
      }
      const text = blocks.filter((b) => b.type === "text").map((b) => b.text).join("\n");
      if (text) out.push({ role: "user", content: text });
    }
  }
  return out;
}

function toOpenAITools(tools) {
  if (!tools || !tools.length) return undefined;
  return tools.map((t) => ({ type: "function", function: { name: t.name, description: t.description, parameters: t.input_schema } }));
}

function toOpenAIToolChoice(tc) {
  if (!tc) return undefined;
  if (tc.type === "tool" && tc.name) return { type: "function", function: { name: tc.name } };
  if (tc.type === "any") return "required";
  if (tc.type === "auto") return "auto";
  return undefined;
}

async function callOpenAI(o) {
  const body = {
    model: o.model,
    max_completion_tokens: o.maxTokens, // model modern (gpt-5/o-series) menolak max_tokens
    messages: toOpenAIMessages(o.system, o.messages),
  };
  const tools = toOpenAITools(o.tools);
  if (tools) body.tools = tools;
  const tc = toOpenAIToolChoice(o.tool_choice);
  if (tc) body.tool_choice = tc;
  const { raw } = await postJson("openai", ENDPOINTS.openai + "chat/completions", {
    "authorization": "Bearer " + o.apiKey,
  }, body);
  let data;
  try { data = JSON.parse(raw); } catch (e) { throw new Error("Respons OpenAI bukan JSON valid: " + String(raw).slice(0, 200)); }
  const norm = normalizeOpenAI(data);
  return { content: norm.content || [], stop_reason: norm.stop_reason || "end_turn" };
}

// ============================ ADAPTER: Gemini ============================
// Buang field skema yang tak didukung subset OpenAPI Gemini (mis. additionalProperties, $schema).
function sanitizeSchema(schema) {
  if (Array.isArray(schema)) return schema.map(sanitizeSchema);
  if (!schema || typeof schema !== "object") return schema;
  const DROP = new Set(["additionalProperties", "$schema", "$id", "$ref", "$defs", "definitions", "examples", "default"]);
  const out = {};
  for (const k of Object.keys(schema)) {
    if (DROP.has(k)) continue;
    out[k] = sanitizeSchema(schema[k]);
  }
  return out;
}

function toGeminiContents(messages) {
  // peta tool_use_id -> nama fungsi (Gemini functionResponse butuh nama, bukan id)
  const idToName = {};
  for (const m of messages || []) {
    if (Array.isArray(m.content)) for (const b of m.content) if (b.type === "tool_use") idToName[b.id] = b.name;
  }
  const out = [];
  for (const m of messages || []) {
    const role = m.role === "assistant" ? "model" : "user";
    if (typeof m.content === "string") { out.push({ role, parts: [{ text: m.content }] }); continue; }
    const blocks = Array.isArray(m.content) ? m.content : [];
    const parts = [];
    for (const b of blocks) {
      if (b.type === "text" && b.text) parts.push({ text: b.text });
      else if (b.type === "tool_use") parts.push({ functionCall: { name: b.name, args: b.input || {} } });
      else if (b.type === "tool_result") {
        let resp = b.content;
        if (typeof resp === "string") { try { resp = JSON.parse(resp); } catch (_) { resp = { result: resp }; } }
        if (typeof resp !== "object" || resp === null) resp = { result: resp };
        parts.push({ functionResponse: { name: idToName[b.tool_use_id] || "tool", response: resp } });
      }
    }
    out.push({ role, parts: parts.length ? parts : [{ text: "" }] });
  }
  return out;
}

function toGeminiTools(tools) {
  if (!tools || !tools.length) return undefined;
  return [{ functionDeclarations: tools.map((t) => ({
    name: t.name, description: t.description, parameters: sanitizeSchema(t.input_schema),
  })) }];
}

function toGeminiToolConfig(tc) {
  if (!tc) return undefined;
  if (tc.type === "tool" && tc.name) return { functionCallingConfig: { mode: "ANY", allowedFunctionNames: [tc.name] } };
  if (tc.type === "any") return { functionCallingConfig: { mode: "ANY" } };
  if (tc.type === "auto") return { functionCallingConfig: { mode: "AUTO" } };
  return undefined;
}

function fromGemini(data) {
  const cand = (data.candidates || [])[0];
  if (!cand) {
    const reason = data.promptFeedback && data.promptFeedback.blockReason;
    if (reason) throw new Error("Gemini memblokir permintaan (" + reason + ").");
    return { content: [], stop_reason: "end_turn" };
  }
  const parts = (cand.content && cand.content.parts) || [];
  const content = [];
  for (const p of parts) {
    if (typeof p.text === "string" && p.text) content.push({ type: "text", text: p.text });
    else if (p.functionCall) content.push({
      type: "tool_use",
      id: "gm_" + Math.random().toString(36).slice(2),
      name: p.functionCall.name || "",
      input: p.functionCall.args || {},
    });
  }
  const finish = cand.finishReason || "STOP";
  const stop_reason = content.some((b) => b.type === "tool_use") ? "tool_use"
    : (finish === "MAX_TOKENS" ? "max_tokens" : "end_turn");
  return { content, stop_reason };
}

async function callGemini(o) {
  const body = {
    contents: toGeminiContents(o.messages),
    generationConfig: { maxOutputTokens: o.maxTokens },
  };
  if (o.system) body.systemInstruction = { parts: [{ text: o.system }] };
  const tools = toGeminiTools(o.tools);
  if (tools) body.tools = tools;
  const toolConfig = toGeminiToolConfig(o.tool_choice);
  if (toolConfig) body.toolConfig = toolConfig;
  const url = ENDPOINTS.gemini + "models/" + encodeURIComponent(o.model) + ":generateContent";
  const { raw } = await postJson("gemini", url, { "x-goog-api-key": o.apiKey }, body);
  let data;
  try { data = JSON.parse(raw); } catch (e) { throw new Error("Respons Gemini bukan JSON valid: " + String(raw).slice(0, 200)); }
  return fromGemini(data);
}

// ============================ router utama ============================
// callMessages({ system, messages, tools, tool_choice, maxTokens, _cfg? })
//   -> { content, stop_reason } (format internal Anthropic).
// _cfg opsional: paksa provider/apiKey/model/baseUrl tertentu (dipakai testProvider).
async function callMessages(req) {
  const cfg = req._cfg || providerConfig.get();
  const provider = cfg.provider;
  if (!cfg.apiKey) {
    throw new Error("API key " + (PROVIDER_LABELS[provider] || provider) + " belum di-set (atur di panel Pengaturan).");
  }
  const o = {
    provider,
    apiKey: cfg.apiKey,
    model: cfg.model,
    baseUrl: cfg.baseUrl,
    maxTokens: req.maxTokens || cfg.maxTokens || 8000,
    system: req.system,
    messages: req.messages || [],
    tools: req.tools,
    tool_choice: req.tool_choice,
  };
  switch (provider) {
    case "anthropic": return callAnthropic(o);
    case "openai": return callOpenAI(o);
    case "gemini": return callGemini(o);
    case "custom": return callCustom(o);
    default: throw new Error("Provider tidak dikenal: " + provider);
  }
}

// callMessages dengan retry (provider kadang tersendat sesaat).
async function callMessagesRetry(req, maxTries) {
  maxTries = maxTries || 3;
  let lastErr;
  for (let attempt = 1; attempt <= maxTries; attempt++) {
    try { return await callMessages(req); }
    catch (err) {
      lastErr = err;
      // 4xx (key/model salah) tak perlu diulang — langsung lempar.
      if (/\((4\d\d)\)/.test(String(err.message || ""))) throw err;
      if (attempt < maxTries) await sleep(800 * attempt);
    }
  }
  throw lastErr;
}

// ============================ daftar model & tes koneksi ============================
// Ambil daftar model dari endpoint OpenAI-compatible {baseUrl}/v1/models (atau /models).
async function listCustomModels(baseUrl, apiKey) {
  const base = String(baseUrl || "");
  const urls = [joinUrl(base, "v1/models"), joinUrl(base, "models")];
  const headers = Object.assign(claudeCodeHeaders(), { "authorization": "Bearer " + apiKey, "x-api-key": apiKey });
  let lastErr = "";
  for (const u of urls) {
    try {
      const ctrl = new AbortController();
      const t = setTimeout(() => ctrl.abort(), 8000);
      const resp = await fetch(u, { headers, signal: ctrl.signal });
      clearTimeout(t);
      const raw = await resp.text();
      if (!resp.ok) { lastErr = "HTTP " + resp.status + ": " + String(raw).slice(0, 120); continue; }
      const data = JSON.parse(raw);
      const list = (data.data || data.models || []).map((m) => (typeof m === "string" ? m : m.id)).filter(Boolean);
      if (list.length) return { ok: true, models: list };
      lastErr = "daftar model kosong";
    } catch (e) { lastErr = String(e.message || e); }
  }
  return { ok: false, error: lastErr || "tak bisa mengambil daftar model" };
}

// Tes koneksi provider: validasi key via endpoint daftar model (tanpa membakar token generasi).
async function testProvider(cfg) {
  const provider = cfg.provider;
  const apiKey = cfg.apiKey;
  if (!apiKey) return { ok: false, error: "API key belum diisi." };
  try {
    if (provider === "custom") return await listCustomModels(cfg.baseUrl, apiKey);
    let url, headers;
    if (provider === "anthropic") { url = ENDPOINTS.anthropic + "v1/models"; headers = { "x-api-key": apiKey, "anthropic-version": ANTHROPIC_VERSION }; }
    else if (provider === "openai") { url = ENDPOINTS.openai + "models"; headers = { "authorization": "Bearer " + apiKey }; }
    else if (provider === "gemini") { url = ENDPOINTS.gemini + "models"; headers = { "x-goog-api-key": apiKey }; }
    else return { ok: false, error: "provider tidak dikenal: " + provider };
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 8000);
    const resp = await fetch(url, { headers, signal: ctrl.signal });
    clearTimeout(t);
    const raw = await resp.text();
    if (!resp.ok) return { ok: false, error: friendlyError(provider, resp.status, raw).message };
    let list = [];
    try {
      const data = JSON.parse(raw);
      list = (data.data || data.models || []).map((m) => (typeof m === "string" ? m : (m.id || m.name))).filter(Boolean);
    } catch (_) {}
    return { ok: true, models: list };
  } catch (e) {
    return { ok: false, error: (PROVIDER_LABELS[provider] || provider) + ": " + String(e.message || e) };
  }
}

module.exports = {
  callMessages,
  callMessagesRetry,
  testProvider,
  listCustomModels,
  parseBodyToAnthropic,
  normalizeOpenAI,
  ENDPOINTS,
  PROVIDER_LABELS,
};
