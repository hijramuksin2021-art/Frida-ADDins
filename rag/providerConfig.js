// rag/providerConfig.js — konfigurasi provider LLM yang bisa diubah RUNTIME.
// Sumber: env (.env) -> dioverride file persist provider.local.json -> dioverride set() dari UI.
// Dipakai server.js (callOnce/callAgentOnce) DAN rag/llm.js, jadi satu sumber kebenaran.
// API key tetap di server (in-memory + file lokal gitignored), tak pernah ke dokumen.

const fs = require("fs");
const path = require("path");

const FILE = path.join(__dirname, "..", "provider.local.json");

// Muat .env bila env belum terisi (untuk pemakaian standalone/test).
let _envLoaded = false;
function ensureEnv() {
  if (_envLoaded || process.env.AERO_API_KEY) { _envLoaded = true; return; }
  try {
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

function normBase(u) { return String(u || "").trim().replace(/\/?$/, "/"); }

function fromEnv() {
  ensureEnv();
  return {
    apiKey: process.env.AERO_API_KEY || "",
    baseUrl: normBase(process.env.AERO_BASE_URL || "https://capi.aerolink.lat/"),
    model: process.env.FRIDA_MODEL || "claude-opus-4-8",
    maxTokens: Number(process.env.FRIDA_MAX_TOKENS || 8000),
  };
}

function loadPersisted() {
  try { return JSON.parse(fs.readFileSync(FILE, "utf8")); } catch (_) { return null; }
}

let active = null;
function init() {
  const env = fromEnv();
  const p = loadPersisted();
  active = Object.assign({}, env, p || {});
  active.baseUrl = normBase(active.baseUrl);
  active.maxTokens = Number(active.maxTokens) || 8000;
}

function get() { if (!active) init(); return active; }

function set(patch) {
  if (!active) init();
  patch = patch || {};
  if (patch.baseUrl !== undefined) active.baseUrl = normBase(patch.baseUrl);
  if (patch.apiKey !== undefined && patch.apiKey !== "") active.apiKey = patch.apiKey;
  if (patch.model !== undefined && patch.model !== "") active.model = patch.model;
  if (patch.maxTokens) active.maxTokens = Number(patch.maxTokens);
  persist();
  return status();
}

function persist() {
  try {
    fs.writeFileSync(FILE, JSON.stringify(
      { baseUrl: active.baseUrl, apiKey: active.apiKey, model: active.model }, null, 2));
  } catch (_) {}
}

// status TANPA membocorkan key.
function status() {
  const a = get();
  return {
    baseUrl: a.baseUrl,
    model: a.model,
    maxTokens: a.maxTokens,
    hasKey: !!a.apiKey,
    keyHint: a.apiKey ? a.apiKey.slice(0, 4) + "…" + a.apiKey.slice(-2) : null,
  };
}

module.exports = { get, set, status, init };
