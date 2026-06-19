// rag/embeddings.js — abstraksi penyedia embeddings ("API key all provider").
// Mendukung provider apa pun yang OpenAI-compatible: POST {baseUrl}/embeddings
// dengan header Authorization: Bearer {apiKey}, body { model, input:[...] }.
// Config via .env (lihat .env.example). Provider 'local' (Xenova) menyusul di R1.
//
// Dipakai mulai R1 (chunk → embed). Di R0 hanya disediakan + bisa dites koneksinya.

// Model lokal multilingual (ID+EN) — terbukti cross-lingual baik, 384-dim.
const LOCAL_MODEL = "Xenova/paraphrase-multilingual-MiniLM-L12-v2";
const LOCAL_DIM = 384;

function readConfig(env) {
  env = env || process.env;
  // Default LOCAL: provider chat (aerolink) tidak menyediakan model embeddings,
  // jadi default ke lokal (tanpa key). Set EMBED_BASE_URL utk pakai provider remote.
  return {
    provider: env.EMBED_PROVIDER || (env.EMBED_BASE_URL ? "openai" : "local"),
    baseUrl: (env.EMBED_BASE_URL || "").replace(/\/+$/, ""),
    apiKey: env.EMBED_API_KEY || "",
    model: env.EMBED_MODEL || LOCAL_MODEL,
    dim: Number(env.EMBED_DIM || 0) || (env.EMBED_BASE_URL ? null : LOCAL_DIM),
  };
}

// ---- Embeddings LOKAL via @xenova/transformers (singleton, ESM dynamic import) ----
let _extractorPromise = null;
function getExtractor() {
  if (!_extractorPromise) {
    _extractorPromise = (async () => {
      const { pipeline } = await import("@xenova/transformers");
      return pipeline("feature-extraction", LOCAL_MODEL);
    })();
  }
  return _extractorPromise;
}
async function embedLocal(texts) {
  const extractor = await getExtractor();
  const out = await extractor(texts, { pooling: "mean", normalize: true });
  return out.tolist(); // number[][], sudah ternormalisasi -> cosine = dot product
}

// Status untuk UI/diagnostik (tanpa membocorkan key).
function status(env) {
  const c = readConfig(env);
  return {
    provider: c.provider,
    model: c.model,
    baseUrl: c.baseUrl || null,
    dim: c.dim,
    configured: c.provider === "local" ? true : Boolean(c.baseUrl && c.apiKey),
  };
}

// embed(texts[]) -> number[][]. Throw bila provider belum siap.
async function embed(texts, env) {
  const c = readConfig(env);
  const input = Array.isArray(texts) ? texts : [texts];
  if (!input.length) return [];

  if (c.provider === "local") {
    return embedLocal(input);
  }

  if (!c.baseUrl || !c.apiKey) {
    throw new Error("Embeddings remote belum dikonfigurasi. Set EMBED_BASE_URL, EMBED_API_KEY, EMBED_MODEL di .env.");
  }

  const resp = await fetch(c.baseUrl + "/embeddings", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "authorization": "Bearer " + c.apiKey,
    },
    body: JSON.stringify({ model: c.model, input }),
  });
  const raw = await resp.text();
  if (!resp.ok) throw new Error("Provider embeddings " + resp.status + ": " + raw.slice(0, 300));
  const data = JSON.parse(raw);
  // format OpenAI: { data:[{embedding:[...]}], ... }
  const vecs = (data.data || []).map((d) => d.embedding);
  if (!vecs.length || !Array.isArray(vecs[0])) {
    throw new Error("Respons embeddings tidak dikenal (tak ada data[].embedding).");
  }
  return vecs;
}

function dim(env) { return readConfig(env).dim; }

module.exports = { readConfig, status, embed, dim, LOCAL_MODEL, LOCAL_DIM };
