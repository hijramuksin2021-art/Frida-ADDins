// rag/generate.js — generasi paragraf akademik yang GROUNDED ke sumber (R2).
// Alur: search → GATE (bukti cukup?) → generasi terbatas-passage → VERIFIKASI sitasi → atribusi.
// Anti-halusinasi: model hanya boleh memakai passages; sitasi yang tak ada di passages ditandai.

const store = require("./store");
const vectors = require("./vectors");
const embeddings = require("./embeddings");
const { callModel, firstToolInput } = require("./llm");
const { verifyCitations } = require("./citations");

const GATE_SCORE = 0.3; // skor cosine minimal agar dianggap "ada bukti"

// GATE "ada bukti" = SKOR retrieval (deterministik di server). Tugas model HANYA menulis dari
// passages. Tidak ada opsi 'needsMoreEvidence' bagi model (dulu membuatnya selalu menolak).
const GROUNDED_SYSTEM = [
  "Tulis SATU paragraf akademik yang menjawab instruksi, MENGGUNAKAN isi PASSAGES yang diberikan.",
  "Aturan: (a) hanya nyatakan yang didukung passages — jangan menambah fakta/angka/nama/tahun/",
  "sitasi di luar passages; (b) PERTAHANKAN sitasi in-text yang ada di passages (mis. '(Nair, 2012)')",
  "— itu sitasi warisan, ikut sumber aslinya, jangan diganti dengan nama dokumen; (c) pertahankan",
  "BAHASA passages (Indonesia/Inggris). Balas via tool submit_grounded_paragraph dengan 'paragraph' terisi.",
].join("\n");

const SUBMIT_TOOL = {
  name: "submit_grounded_paragraph",
  description: "Kirim paragraf akademik hasil, berdasarkan passages.",
  input_schema: {
    type: "object",
    properties: {
      paragraph: { type: "string", description: "Paragraf hasil (wajib, berbahasa sama dengan passages)." },
      usedSourceIds: { type: "array", items: { type: "string" }, description: "source_id passage yang dipakai." },
    },
    required: ["paragraph"],
  },
};

function buildUserContent(instruction, passages) {
  return "INSTRUKSI PENGGUNA:\n" + instruction +
    "\n\nPASSAGES (HANYA gunakan ini sebagai sumber fakta & sitasi):\n" +
    passages.map((p) => "[" + p.n + "] (source_id=" + p.source_id + ")\n" + p.text).join("\n\n");
}

// Buang sitasi parenthetical yang DITANDAI (tak terverifikasi) dari paragraf.
function stripFlaggedParenthetical(paragraph, flagged) {
  let out = paragraph;
  flagged.forEach((raw) => {
    if (raw.startsWith("(")) {
      out = out.split(" " + raw).join("").split(raw).join("");
    }
  });
  return out.replace(/\s{2,}/g, " ").replace(/\s+([.,])/g, "$1").trim();
}

async function generate_paragraph_from_source(input) {
  input = input || {};
  const instruction = input.instruction || input.source_query || "";
  if (!instruction) return { error: "instruction kosong" };

  const all = store.list();
  if (!all.length) return { needsMoreEvidence: true, note: "Belum ada sumber terunggah." };

  const docIds = (input.document_ids && input.document_ids.length)
    ? input.document_ids : all.map((d) => d.id);
  const query = input.source_query || instruction;

  let qvec;
  try { [qvec] = await embeddings.embed([query]); }
  catch (e) { return { error: "embeddings gagal: " + (e.message || e) }; }

  const hits = vectors.search(qvec, docIds, { k: input.k || 6, minScore: GATE_SCORE });
  if (!hits.length) {
    return { needsMoreEvidence: true,
      note: "Bukti di sumber tak cukup untuk menulis ini (tak ada kutipan relevan di atas ambang). " +
            "Saya tidak akan mengarang. Coba persempit/ubah permintaan, atau pastikan sumber sudah diindeks." };
  }

  const passages = hits.map((h, i) => ({ n: i + 1, source_id: h.source_id, text: h.text }));

  let data;
  try {
    data = await callModel({
      system: GROUNDED_SYSTEM,
      messages: [{ role: "user", content: buildUserContent(instruction, passages) }],
      tools: [SUBMIT_TOOL],
      tool_choice: { type: "tool", name: "submit_grounded_paragraph" },
      maxTokens: 1500,
    });
  } catch (e) { return { error: "generasi gagal: " + (e.message || e) }; }

  const out = firstToolInput(data) || {};
  if (!out.paragraph || !out.paragraph.trim()) {
    return { needsMoreEvidence: true,
      note: "Paragraf tidak dihasilkan dari sumber. Coba ubah/persempit permintaan." };
  }

  // VERIFIKASI sitasi terhadap chunk sumber.
  const v = verifyCitations(out.paragraph, hits);
  const cleaned = v.flagged.length ? stripFlaggedParenthetical(out.paragraph, v.flagged) : out.paragraph;

  const titles = {};
  all.forEach((d) => (titles[d.id] = { title: d.title, year: d.year, doi: d.doi }));
  // Sumber = source_id dari chunk yang BENAR-BENAR diambil (andal; jangan percaya nomor dari model).
  const usedIds = [...new Set(hits.map((h) => h.source_id))];

  return {
    needsMoreEvidence: false,
    paragraph: cleaned,                 // paragraf siap pakai (sitasi tak-terverifikasi dibuang)
    paragraphRaw: out.paragraph,
    verifiedCitations: v.verified,      // sitasi WARISAN yang sah (ada di sumber)
    flaggedCitations: v.flagged,        // tak ditemukan di sumber -> dibuang dari `paragraph`
    primarySource: usedIds.map((id) => ({ source_id: id, title: (titles[id] || {}).title || null,
                                          year: (titles[id] || {}).year || null })),
    note: v.flagged.length
      ? ("Catatan: " + v.flagged.length + " sitasi tak terverifikasi di sumber dan SUDAH DIBUANG: " +
         v.flagged.join(", ") + ". Klaim tanpa sitasi in-text berasal dari dokumen sumber (sitasi primer diformat di tahap berikutnya).")
      : "Sitasi in-text terverifikasi ada di sumber. Klaim lain berasal dari dokumen sumber (sitasi primer).",
  };
}

module.exports = { generate_paragraph_from_source };
