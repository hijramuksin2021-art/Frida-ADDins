// rag/summarize.js — summarize_source + compare_sources (server tools, pakai LLM via callModel).
// Prinsip: HANYA gunakan teks dari sumber — tidak menambah fakta di luar passage.
// Teks panjang dipotong agar hemat token (MAX_CHARS per sumber).

const store = require("./store");
const { callModel, firstToolInput } = require("./llm");

const MAX_CHARS = 4000; // karakter maks teks per sumber yang dikirim ke LLM
const MAX_CHARS_COMPARE = 2500; // per sumber saat compare (lebih banyak sumber)

function truncate(text, max) {
  if (!text) return "";
  return text.length > max ? text.slice(0, max) + "…[dipotong]" : text;
}

// Tool forced untuk memaksa output terstruktur dari LLM (anti-hallucination; jangan percaya teks bebas).
const SUMMARIZE_TOOL = {
  name: "submit_summary",
  description: "Kirim ringkasan sumber.",
  input_schema: {
    type: "object",
    properties: {
      summary: { type: "string", description: "Ringkasan (hanya dari teks, bukan pengetahuan luar)." },
    },
    required: ["summary"],
  },
};

const COMPARE_TOOL = {
  name: "submit_comparison",
  description: "Kirim hasil perbandingan sumber.",
  input_schema: {
    type: "object",
    properties: {
      comparison: { type: "string", description: "Paragraf perbandingan (HANYA dari teks passage)." },
      similarities: { type: "array", items: { type: "string" }, description: "Poin kesamaan utama." },
      differences: { type: "array", items: { type: "string" }, description: "Poin perbedaan utama." },
    },
    required: ["comparison"],
  },
};

// ---- summarize_source ----
// Input: { source_id, aspect?, max_sentences? }
// Output: { source_id, title, summary } | { error }
async function summarize_source(input) {
  input = input || {};
  const id = input.source_id;
  if (!id) return { error: "source_id wajib diisi" };

  const doc = store.get(id);
  if (!doc) return { error: "sumber tidak ditemukan: " + id };

  const text = truncate(doc.text || "", MAX_CHARS);
  if (!text.trim()) {
    return { error: "Teks sumber kosong. Klik 'Indeks ulang sumber' terlebih dahulu." };
  }

  const sentLimit = input.max_sentences ? "Maksimal " + input.max_sentences + " kalimat." : "3–5 kalimat.";
  const focus = input.aspect ? " Fokus pada: " + input.aspect + "." : "";

  const system = [
    "Anda meringkas sumber ilmiah secara akurat.",
    "ATURAN: (1) HANYA gunakan informasi dari teks yang diberikan — jangan tambah fakta/angka/nama di luar teks.",
    "(2) Pertahankan bahasa teks asli (Indonesia/Inggris). (3) " + sentLimit,
    "Balas via tool submit_summary.",
  ].join(" ");

  const userContent =
    "Buat ringkasan sumber berikut." + focus +
    "\n\nJUDUL: " + (doc.title || doc.filename) +
    "\n\nTEKS:\n" + text;

  let data;
  try {
    data = await callModel({
      system,
      messages: [{ role: "user", content: userContent }],
      tools: [SUMMARIZE_TOOL],
      tool_choice: { type: "tool", name: "submit_summary" },
      maxTokens: 800,
    });
  } catch (e) {
    return { error: "generasi ringkasan gagal: " + (e.message || e) };
  }

  const out = firstToolInput(data) || {};
  if (!out.summary || !out.summary.trim()) {
    return { error: "LLM tidak menghasilkan ringkasan. Coba lagi." };
  }

  return {
    source_id: id,
    title: doc.title || doc.filename,
    year: doc.year || null,
    aspect: input.aspect || null,
    summary: out.summary.trim(),
  };
}

// ---- compare_sources ----
// Input: { source_ids:string[], aspect? }
// Output: { source_ids, titles, aspect, comparison, similarities, differences } | { error }
async function compare_sources(input) {
  input = input || {};
  const ids = input.source_ids;
  if (!Array.isArray(ids) || ids.length < 2) {
    return { error: "compare_sources membutuhkan minimal 2 source_id dalam source_ids[]" };
  }

  const passages = [];
  const missing = [];
  for (const id of ids.slice(0, 5)) { // maks 5 sumber agar tidak meledak token
    const doc = store.get(id);
    if (!doc) { missing.push(id); continue; }
    const text = truncate(doc.text || "", MAX_CHARS_COMPARE);
    if (!text.trim()) { missing.push(id); continue; }
    passages.push({ id, title: doc.title || doc.filename, year: doc.year, text });
  }

  if (passages.length < 2) {
    return {
      error: "Kurang dari 2 sumber valid untuk dibandingkan." +
        (missing.length ? " Tidak ditemukan/teks kosong: " + missing.join(", ") : ""),
    };
  }

  const focus = input.aspect ? " Fokus perbandingan pada: " + input.aspect + "." : "";
  const system = [
    "Anda membandingkan sumber-sumber ilmiah secara akurat.",
    "ATURAN: (1) HANYA gunakan informasi dari passages yang diberikan — jangan tambah fakta di luar passages.",
    "(2) Pertahankan bahasa passages. (3) Sertakan nama sumber saat menyebut temuan spesifik.",
    "Balas via tool submit_comparison.",
  ].join(" ");

  const srcList = passages
    .map((p, i) =>
      "SUMBER " + (i + 1) + " [" + p.id + "] \"" + p.title + "\"" +
      (p.year ? " (" + p.year + ")" : "") + ":\n" + p.text
    )
    .join("\n\n---\n\n");

  const userContent =
    "Bandingkan sumber-sumber berikut." + focus +
    "\n\nGunakan nama/judul sumber saat menyebut temuan.\n\n" + srcList;

  let data;
  try {
    data = await callModel({
      system,
      messages: [{ role: "user", content: userContent }],
      tools: [COMPARE_TOOL],
      tool_choice: { type: "tool", name: "submit_comparison" },
      maxTokens: 1500,
    });
  } catch (e) {
    return { error: "generasi perbandingan gagal: " + (e.message || e) };
  }

  const out = firstToolInput(data) || {};
  if (!out.comparison || !out.comparison.trim()) {
    return { error: "LLM tidak menghasilkan perbandingan. Coba lagi." };
  }

  return {
    source_ids: passages.map((p) => p.id),
    titles: passages.map((p) => p.title),
    aspect: input.aspect || null,
    comparison: out.comparison.trim(),
    similarities: out.similarities || [],
    differences: out.differences || [],
    note: missing.length ? "Sumber tidak ditemukan/kosong: " + missing.join(", ") : null,
  };
}

module.exports = { summarize_source, compare_sources };
