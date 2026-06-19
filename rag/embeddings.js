// rag/embeddings.js — abstraksi penyedia embeddings ("API key all provider").
// Mendukung provider apa pun yang OpenAI-compatible: POST {baseUrl}/embeddings
// dengan header Authorization: Bearer {apiKey}, body { model, input:[...] }.
// Config via .env (lihat .env.example). Provider 'local' (Xenova) menyusul di R1.
//
// Dipakai mulai R1 (chunk → embed). Di R0 hanya disediakan + bisa dites koneksinya.

function readConfig(env) {
  env = env || process.env;
  return {
    provider: env.EMBED_PROVIDER || (env.EMBED_BASE_URL ? "openai" : "local"),
    baseUrl: (env.EMBED_BASE_URL || "").replace(/\/+$/, ""),
    apiKey: env.EMBED_API_KEY || "",
    model: env.EMBED_MODEL || "text-embedding-3-large",
    dim: Number(env.EMBED_DIM || 0) || null,
  };
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

  if (c.provider === "local") {
    // Implementasi lokal (@xenova/transformers) ditambahkan di R1.
    throw new Error("Embeddings lokal belum aktif (akan ditambahkan di R1). " +
      "Untuk sekarang set EMBED_BASE_URL + EMBED_API_KEY + EMBED_MODEL di .env.");
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

module.exports = { readConfig, status, embed };
