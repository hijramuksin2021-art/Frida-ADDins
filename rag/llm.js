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
  return JSON.parse(raw);
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
