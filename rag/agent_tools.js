// rag/agent_tools.js — eksekusi tool RAG di SISI SERVER (runtime:"server").
// Berbeda dari tool Word (client): tool ini tak menyentuh Office.js, jadi dijalankan
// langsung di server dalam loop agentic. Mengembalikan objek hasil (jadi tool_result).

const store = require("./store");
const vectors = require("./vectors");
const embeddings = require("./embeddings");
const { generate_paragraph_from_source } = require("./generate");
const { resolveSourceTool } = require("./aliases");               // R4
const { summarize_source, compare_sources } = require("./summarize"); // R4

const MAX_CHUNK_CHARS = 700; // batasi teks per hit agar hemat token

async function search_uploaded_sources(input) {
  const query = (input && input.query) || "";
  if (!query) return { error: "query kosong" };
  const all = store.list();
  if (!all.length) return { hits: [], note: "Belum ada sumber terunggah." };

  const docIds = (input.document_ids && input.document_ids.length)
    ? input.document_ids
    : all.map((d) => d.id);

  let qvec;
  try { [qvec] = await embeddings.embed([query]); }
  catch (e) { return { error: "embeddings gagal: " + (e.message || e) }; }

  const hits = vectors.search(qvec, docIds, { k: input.k || 6 });
  const titles = {};
  all.forEach((d) => (titles[d.id] = d.title));

  if (!hits.length) {
    return { hits: [], note: "Tidak ada kutipan relevan di sumber (bukti tak cukup). " +
      "Jika sumber belum diindeks, minta pengguna klik 'Indeks ulang sumber'." };
  }
  return {
    hits: hits.map((h) => ({
      source_id: h.document_id,
      title: titles[h.document_id] || null,
      chunk_id: h.chunk_id,
      score: Number(h.score.toFixed(3)),
      text: (h.text || "").slice(0, MAX_CHUNK_CHARS),
    })),
  };
}

// R4: resolve_source — nama natural → document_id (tanpa embedding; scoring kata kunci)
async function resolve_source(input) {
  return resolveSourceTool(input || {});
}

// R4: summarize_source — ringkas 1 sumber via LLM (grounded ke teks sumber)
async function summarize_source_tool(input) {
  return summarize_source(input || {});
}

// R4: compare_sources — banding ≥2 sumber via LLM (grounded ke teks sumber)
async function compare_sources_tool(input) {
  return compare_sources(input || {});
}

const SERVER_TOOLS = {
  search_uploaded_sources,
  generate_paragraph_from_source,
  resolve_source,           // R4
  summarize_source,         // R4
  compare_sources,          // R4
};

// Eksekusi satu tool server berdasarkan nama (sudah dikanonikkan oleh pemanggil).
async function executeServerTool(name, input) {
  const fn = SERVER_TOOLS[name];
  if (!fn) return { error: "tool server tidak dikenal: " + name };
  try { return await fn(input || {}); }
  catch (e) { return { error: String(e.message || e) }; }
}

module.exports = { executeServerTool, SERVER_TOOLS };

