// rag/providerConfig.js — konfigurasi provider LLM MULTI-PROVIDER yang bisa diubah RUNTIME.
// Sumber: env (.env) -> dioverride file persist provider.local.json -> dioverride set() dari UI.
// Dipakai server.js, rag/aiProvider.js, dan rag/llm.js, jadi satu sumber kebenaran.
// API key tetap di server (in-memory + file lokal gitignored), tak pernah ke dokumen.
//
// Struktur baru (multi-provider):
//   {
//     activeProvider: "custom",
//     providers: {
//       anthropic: { apiKey, model },
//       openai:    { apiKey, model },
//       gemini:    { apiKey, model },
//       custom:    { apiKey, baseUrl, model }   // hanya custom yang punya baseUrl
//     },
//     maxTokens: 8000
//   }
// Backward-compat: provider.local.json lama berformat flat { baseUrl, apiKey, model }
// otomatis dipetakan ke providers.custom + activeProvider="custom" (setting lama tak hilang).

const fs = require("fs");
const path = require("path");

const FILE = path.join(__dirname, "..", "provider.local.json");

// Daftar provider yang dikenal. 'custom' = OpenAI-compatible (9Router/Aerolink/proxy).
const PROVIDERS = ["anthropic", "openai", "gemini", "custom"];

// Model default per provider (dipakai bila belum ada pilihan tersimpan).
const DEFAULT_MODELS = {
  anthropic: "claude-opus-4-8",
  openai: "gpt-5",
  gemini: "gemini-2.5-pro",
  custom: "",
};

// Base URL default untuk custom saja. Provider resmi endpoint-nya dikunci di aiProvider.js.
const DEFAULT_CUSTOM_BASE = "https://capi.aerolink.lat/";

// Muat .env bila env belum terisi (untuk pemakaian standalone/test).
let _envLoaded = false;
function ensureEnv() {
  if (_envLoaded) return;
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

// Konfigurasi awal dari environment. AERO_* -> custom (kompatibel lama);
// ANTHROPIC_API_KEY / OPENAI_API_KEY / GEMINI_API_KEY -> provider resmi masing-masing.
function fromEnv() {
  ensureEnv();
  return {
    activeProvider: process.env.FRIDA_PROVIDER || "custom",
    maxTokens: Number(process.env.FRIDA_MAX_TOKENS || 8000),
    providers: {
      anthropic: {
        apiKey: process.env.ANTHROPIC_API_KEY || "",
        model: process.env.FRIDA_ANTHROPIC_MODEL || DEFAULT_MODELS.anthropic,
      },
      openai: {
        apiKey: process.env.OPENAI_API_KEY || "",
        model: process.env.FRIDA_OPENAI_MODEL || DEFAULT_MODELS.openai,
      },
      gemini: {
        apiKey: process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY || "",
        model: process.env.FRIDA_GEMINI_MODEL || DEFAULT_MODELS.gemini,
      },
      custom: {
        apiKey: process.env.AERO_API_KEY || "",
        baseUrl: normBase(process.env.AERO_BASE_URL || DEFAULT_CUSTOM_BASE),
        model: process.env.FRIDA_MODEL || DEFAULT_MODELS.custom,
      },
    },
  };
}

function loadPersisted() {
  try { return JSON.parse(fs.readFileSync(FILE, "utf8")); } catch (_) { return null; }
}

// Ubah file lama (flat) -> struktur multi-provider (hanya isi bagian custom).
function migrateLegacy(p) {
  if (!p || typeof p !== "object") return null;
  if (p.providers) return null; // sudah format baru
  if (p.apiKey === undefined && p.baseUrl === undefined && p.model === undefined) return null;
  return {
    activeProvider: "custom",
    providers: {
      custom: {
        apiKey: p.apiKey || "",
        baseUrl: normBase(p.baseUrl || DEFAULT_CUSTOM_BASE),
        model: p.model || "",
      },
    },
  };
}

// Gabungkan konfigurasi provider (base <- override) tanpa menimpa dengan nilai kosong.
function mergeProvider(base, over) {
  const out = Object.assign({}, base);
  if (!over) return out;
  if (over.apiKey) out.apiKey = over.apiKey;                 // key kosong tak menimpa
  if (over.model) out.model = over.model;
  if (over.baseUrl !== undefined && over.baseUrl !== "") out.baseUrl = normBase(over.baseUrl);
  return out;
}

let active = null;

function init() {
  const env = fromEnv();
  let persisted = loadPersisted();
  const migrated = migrateLegacy(persisted);
  if (migrated) persisted = migrated;

  active = { activeProvider: env.activeProvider, maxTokens: env.maxTokens, providers: {} };
  for (const id of PROVIDERS) {
    active.providers[id] = mergeProvider(env.providers[id], persisted && persisted.providers && persisted.providers[id]);
  }
  if (persisted && PROVIDERS.includes(persisted.activeProvider)) {
    active.activeProvider = persisted.activeProvider;
  }
  active.maxTokens = Number((persisted && persisted.maxTokens) || env.maxTokens) || 8000;
  // pastikan custom selalu punya baseUrl
  if (!active.providers.custom.baseUrl) active.providers.custom.baseUrl = DEFAULT_CUSTOM_BASE;
}

function ensure() { if (!active) init(); return active; }

// Provider aktif ter-resolusi -> dipakai adapter. baseUrl hanya relevan utk custom.
function get() {
  const a = ensure();
  const id = a.activeProvider;
  const pc = a.providers[id] || {};
  return {
    provider: id,
    apiKey: pc.apiKey || "",
    model: pc.model || DEFAULT_MODELS[id] || "",
    baseUrl: id === "custom" ? normBase(pc.baseUrl || DEFAULT_CUSTOM_BASE) : undefined,
    maxTokens: a.maxTokens || 8000,
  };
}

// Config ter-resolusi untuk provider tertentu (termasuk apiKey — hanya dipakai di server).
function getProvider(id) {
  const a = ensure();
  if (!PROVIDERS.includes(id)) throw new Error("Provider tidak dikenal: " + id);
  const pc = a.providers[id] || {};
  return {
    provider: id,
    apiKey: pc.apiKey || "",
    model: pc.model || DEFAULT_MODELS[id] || "",
    baseUrl: id === "custom" ? normBase(pc.baseUrl || DEFAULT_CUSTOM_BASE) : undefined,
    maxTokens: a.maxTokens || 8000,
  };
}

function getActive() { return ensure().activeProvider; }

function setActive(id) {
  ensure();
  if (!PROVIDERS.includes(id)) throw new Error("Provider tidak dikenal: " + id);
  active.activeProvider = id;
  persist();
  return status();
}

// Simpan konfigurasi satu provider (key kosong tak menimpa). baseUrl hanya utk custom.
function setProvider(id, patch) {
  ensure();
  if (!PROVIDERS.includes(id)) throw new Error("Provider tidak dikenal: " + id);
  patch = patch || {};
  const cur = active.providers[id] || {};
  active.providers[id] = mergeProvider(cur, {
    apiKey: patch.apiKey,
    model: patch.model,
    baseUrl: id === "custom" ? patch.baseUrl : undefined,
  });
  if (patch.maxTokens) active.maxTokens = Number(patch.maxTokens) || active.maxTokens;
  persist();
  return status();
}

function persist() {
  try { fs.writeFileSync(FILE, JSON.stringify(active, null, 2)); } catch (_) {}
}

function keyHint(k) { return k ? k.slice(0, 4) + "…" + k.slice(-2) : null; }

// status TANPA membocorkan key (per-provider hasKey + hint).
function status() {
  const a = ensure();
  const providers = {};
  for (const id of PROVIDERS) {
    const pc = a.providers[id] || {};
    providers[id] = {
      model: pc.model || DEFAULT_MODELS[id] || "",
      hasKey: !!pc.apiKey,
      keyHint: keyHint(pc.apiKey),
    };
    if (id === "custom") providers[id].baseUrl = pc.baseUrl || DEFAULT_CUSTOM_BASE;
  }
  return {
    activeProvider: a.activeProvider,
    maxTokens: a.maxTokens || 8000,
    providers,
  };
}

module.exports = {
  get, getProvider, getActive, setActive, setProvider, status, init,
  PROVIDERS, DEFAULT_MODELS, DEFAULT_CUSTOM_BASE,
};
