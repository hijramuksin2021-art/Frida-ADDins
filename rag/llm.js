// rag/llm.js — pemanggil model (Claude via aerolink) untuk tugas RAG di server.
// Dipakai generate_paragraph_from_source (generasi grounded dgn output terstruktur
// lewat forced tool). Membaca konfigurasi dari env saat dipanggil (.env sudah dimuat
// server saat start).

// Muat .env sekali bila env belum terisi (mis. modul dipakai di luar server.js).
let _envLoaded = false;
function ensureEnv() {
  if (_envLoaded || process.env.AERO_API_KEY) { _envLoaded = true; return; }
  try {
    const fs = require("fs"), path = require("path");
    const p = path.join(__dirname, "..", ".env");
    if (fs.existsSync(p)) {
      for (const line of fs.readFileSync(p, "utf8").split(/\r?\n/)) {
        const m = line.match(/^\s*([\w.-]+)\s*=\s*(.*)\s*$/);
        if (!m || line.trim().startsWith("#")) continue;
        let v = m[2].trim();
        if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) v = v.slice(1, -1);
        if (!(m[1] in process.env)) process.env[m[1]] = v;
      }
    }
  } catch (_) {}
  _envLoaded = true;
}

function cfg() {
  ensureEnv();
  return {
    apiKey: process.env.AERO_API_KEY,
    baseUrl: (process.env.AERO_BASE_URL || "https://capi.aerolink.lat/").replace(/\/?$/, "/"),
    model: process.env.FRIDA_MODEL || "claude-opus-4-8",
    maxTokens: Number(process.env.FRIDA_MAX_TOKENS || 8000),
  };
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// --- Parser SSE → Anthropic response object ---
// Router (9router/LiteLLM/aerolink) bisa mengembalikan:
//   (a) SSE stream (Content-Type: text/event-stream) — selalu dari router lokal
//   (b) JSON biasa (Content-Type: application/json) — aerolink asli
// Kedua kasus dirakit menjadi format Anthropic standar: { content[], stop_reason }
function parseBodyToAnthropic(contentType, raw) {
  const isStream = (contentType || "").includes("event-stream");

  if (!isStream) {
    // JSON biasa: Anthropic native atau OpenAI format
    let data;
    try { data = JSON.parse(raw); } catch (e) { throw new Error("Respons bukan JSON valid: " + raw.slice(0, 200)); }
    return normalizeOpenAI(data);
  }

  // SSE stream: rakit dari event-data lines
  // Format Anthropic SSE: message_start → content_block_start/delta/stop → message_delta → message_stop
  const lines = raw.split(/\r?\n/);
  const blocks = {}; // index → {type, id, name, partialJson}
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

  // Rakit content array
  const content = Object.keys(blocks).sort((a, b) => Number(a) - Number(b)).map(idx => {
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

// Normalisasi format OpenAI → Anthropic (untuk non-stream JSON)
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

// callModel({system, messages, tools, tool_choice, maxTokens}) -> data (Anthropic response)
async function callModelOnce(opts) {
  const c = cfg();
  if (!c.apiKey) throw new Error("AERO_API_KEY belum di-set");
  const resp = await fetch(c.baseUrl + "v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": c.apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: c.model,
      max_tokens: opts.maxTokens || c.maxTokens,
      system: opts.system,
      tools: opts.tools,
      tool_choice: opts.tool_choice,
      messages: opts.messages,
    }),
  });
  const raw = await resp.text();
  if (!resp.ok) throw new Error("Provider " + resp.status + ": " + raw.slice(0, 300));
  return parseBodyToAnthropic(resp.headers.get("content-type"), raw);
}

async function callModel(opts) {
  let lastErr;
  for (let i = 1; i <= 3; i++) {
    try { return await callModelOnce(opts); }
    catch (e) { lastErr = e; if (i < 3) await sleep(600 * i); }
  }
  throw lastErr;
}

// Ambil input dari blok tool_use pertama (abaikan nama, provider kadang me-rename).
function firstToolInput(data) {
  const b = (data.content || []).find((x) => x.type === "tool_use" && x.input);
  return b ? b.input : null;
}

module.exports = { callModel, firstToolInput };

