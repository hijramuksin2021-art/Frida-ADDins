// rag/agent_tools.js — eksekusi tool RAG di SISI SERVER (runtime:"server").
// Berbeda dari tool Word (client): tool ini tak menyentuh Office.js, jadi dijalankan
// langsung di server dalam loop agentic. Mengembalikan objek hasil (jadi tool_result).

const store = require("./store");
const { generate_paragraph_from_source } = require("./generate");
const { resolveSourceTool } = require("./aliases");               // R4
const { summarize_source, compare_sources } = require("./summarize"); // R4

const MAX_CHUNK_CHARS = 700; // batasi teks per hit agar hemat token

async function search_uploaded_sources(input) {
  const workspace = input.workspace;
  const docs = store.list(workspace);
  if (!docs.length) return { error: "Belum ada dokumen diupload di workspace ini." };

  const MAX_TOTAL_CHARS = 150000; // batas aman gabungan semua dokumen (~35-40rb token)
  let combined = [];
  let totalChars = 0;
  const skipped = [];

  const docIds = (input.document_ids && input.document_ids.length)
    ? input.document_ids
    : docs.map((d) => d.id);

  for (const d of docs) {
    if (!docIds.includes(d.id)) continue;
    const full = store.get(d.id);
    if (!full || !full.text) continue;
    if (totalChars + full.text.length > MAX_TOTAL_CHARS) {
      skipped.push(d.filename || d.title);
      continue;
    }
    combined.push({
      title: d.title || d.filename,
      text: full.text,
      document_id: d.id
    });
    totalChars += full.text.length;
  }

  if (!combined.length) {
    return { error: "Dokumen terlalu besar untuk disertakan langsung (>150rb karakter). " +
      "Fitur pencarian sebagian dokumen belum aktif untuk kasus ini." };
  }

  return {
    documents: combined,
    note: skipped.length
      ? "Catatan: " + skipped.length + " dokumen lain (" + skipped.join(", ") + ") tidak " +
        "disertakan karena keterbatasan ukuran gabungan. Sebutkan ke user kalau relevan."
      : undefined,
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

