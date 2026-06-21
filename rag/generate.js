// rag/generate.js — generasi paragraf akademik yang GROUNDED ke sumber (R2).
// Alur: search → GATE (bukti cukup?) → generasi terbatas-passage → VERIFIKASI sitasi → atribusi.
// Anti-halusinasi: model hanya boleh memakai passages; sitasi yang tak ada di passages ditandai.

const store = require("./store");
const vectors = require("./vectors");
const embeddings = require("./embeddings");
const { callModel, firstToolInput } = require("./llm");
const { verifyCitations } = require("./citations");
const { verifyQuotes } = require("./quotecheck");
const { verifyFaithfulness } = require("./faithfulness");
const { logGeneration } = require("./analytics");
const guidelineConfig = require("./guidelineConfig");

const GATE_SCORE = Number(process.env.GATE_SCORE_MIN || 0.3); // skor cosine minimal agar dianggap "ada bukti"

// GATE "ada bukti" = SKOR retrieval (deterministik di server). Tugas model HANYA menulis dari
// passages. Tidak ada opsi 'needsMoreEvidence' bagi model (dulu membuatnya selalu menolak).
const GROUNDED_SYSTEM_BASE = [
  "Tulis SATU paragraf akademik yang menjawab instruksi, MENGGUNAKAN isi PASSAGES yang diberikan.",
  "Aturan: (a) hanya nyatakan yang didukung passages — jangan menambah fakta/angka/nama/tahun/",
  "sitasi di luar passages; (b) PERTAHANKAN sitasi in-text yang ada di passages (mis. '(Nair, 2012)')",
  "— itu sitasi warisan, ikut sumber aslinya, jangan diganti dengan nama dokumen; (c) pertahankan",
  "BAHASA passages (Indonesia/Inggris). Balas via tool submit_grounded_paragraph dengan 'paragraph' terisi.",
].join("\n");

function getGroundedSystem() {
  let prompt = GROUNDED_SYSTEM_BASE;
  const gl = guidelineConfig.getActiveGuideline();
  if (gl) {
    prompt += "\n\nATURAN PANDUAN PENULISAN AKTIF (" + gl.nama + "):\n";
    if (gl.format_umum && gl.format_umum.aturan_teks_khusus) {
      prompt += "- Istilah asing: " + gl.format_umum.aturan_teks_khusus.istilah_asing_dan_lokal + "\n";
      prompt += "- Angka: " + gl.format_umum.aturan_teks_khusus.angka_kurang_dari_10_dalam_kalimat + "\n";
    }
    if (gl.aturan_plagiarisme) {
      prompt += "- Hindari plagiarisme (batas " + gl.aturan_plagiarisme.batas_maksimal + "). " + gl.aturan_plagiarisme.definisi + ".\n";
    }
    if (gl.sitasi && gl.sitasi.gaya) {
      prompt += "- Gaya Sitasi Utama: " + gl.sitasi.gaya + "\n";
    }
  }
  return prompt;
}

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
    const allHits = vectors.search(qvec, docIds, { k: 1 });
    const maxScore = allHits.length ? allHits[0].score : 0;
    logGeneration({ accepted: false, reason: "insufficient_evidence", maxScore, query });
    return { needsMoreEvidence: true,
      note: "Bukti di sumber tak cukup untuk menulis ini (tak ada kutipan relevan di atas ambang). " +
            "Saya tidak akan mengarang. Coba persempit/ubah permintaan, atau pastikan sumber sudah diindeks." };
  }

  const passages = hits.map((h, i) => ({ n: i + 1, source_id: h.source_id, text: h.text }));

  let data;
  try {
    data = await callModel({
      system: getGroundedSystem(),
      messages: [{ role: "user", content: buildUserContent(instruction, passages) }],
      tools: [SUBMIT_TOOL],
      tool_choice: { type: "tool", name: "submit_grounded_paragraph" },
      maxTokens: 1500,
    });
  } catch (e) {
    logGeneration({ accepted: false, reason: "llm_error", maxScore: hits[0].score, query });
    return { error: "generasi gagal: " + (e.message || e) };
  }

  const out = firstToolInput(data) || {};
  if (!out.paragraph || !out.paragraph.trim()) {
    logGeneration({ accepted: false, reason: "empty_generation", maxScore: hits[0].score, query });
    return { needsMoreEvidence: true,
      note: "Paragraf tidak dihasilkan dari sumber. Coba ubah/persempit permintaan." };
  }

  // VERIFIKASI sitasi terhadap chunk sumber.
  const v = verifyCitations(out.paragraph, hits);
  const cleaned = v.flagged.length ? stripFlaggedParenthetical(out.paragraph, v.flagged) : out.paragraph;

  // R6: VERIFIKASI kutipan (quote-check)
  const quoteCheck = verifyQuotes(cleaned, hits);
  let quoteCleaned = cleaned;
  if (quoteCheck.flagged.length > 0) {
    // Buang kutipan yang tidak terverifikasi dari paragraf
    quoteCheck.flagged.forEach((q) => {
      quoteCleaned = quoteCleaned.replace(q.raw, "[kutipan dihapus]");
    });
  }

  // R6: VERIFIKASI faithfulness (konsistensi dengan sumber)
  const faithfulness = await verifyFaithfulness(quoteCleaned, hits);
  if (!faithfulness.overall_faithful && faithfulness.contradictions.length > 0) {
    logGeneration({ accepted: false, reason: "faithfulness_contradiction", maxScore: hits[0].score, query, contradictions: faithfulness.contradictions });
    return {
      needsMoreEvidence: true,
      note: "Paragraf mengandung kontradiksi dengan sumber dan ditolak. " +
            faithfulness.contradictions.map(c => `"${c.sentence}": ${c.reason}`).join("; "),
      contradictions: faithfulness.contradictions
    };
  }

  const titles = {};
  all.forEach((d) => (titles[d.id] = { title: d.title, year: d.year, doi: d.doi }));
  // Sumber = source_id dari chunk yang BENAR-BENAR diambil (andal; jangan percaya nomor dari model).
  const usedIds = [...new Set(hits.map((h) => h.source_id))];

  logGeneration({
    accepted: true,
    maxScore: hits[0].score,
    query,
    verifiedCitations: v.verified.length,
    verifiedQuotes: quoteCheck.verified.length,
    faithfulnessFindings: faithfulness.all_findings ? faithfulness.all_findings.length : 0
  });

  let note = "";
  if (v.flagged.length) {
    note += "Catatan: " + v.flagged.length + " sitasi tak terverifikasi di sumber dan SUDAH DIBUANG: " + v.flagged.join(", ") + ". ";
  }
  if (quoteCheck.flagged.length) {
    note += "Catatan: " + quoteCheck.flagged.length + " kutipan tak terverifikasi di sumber dan dibuang: " + quoteCheck.flagged.map(q => q.raw).join(", ") + ". ";
  }
  if (!note) {
    note = "Sitasi in-text dan kutipan terverifikasi ada di sumber. Klaim lain berasal dari dokumen sumber (sitasi primer).";
  }

  return {
    needsMoreEvidence: false,
    paragraph: quoteCleaned,            // paragraf siap pakai (sitasi & kutipan tak-terverifikasi dibuang)
    paragraphRaw: out.paragraph,
    verifiedCitations: v.verified,      // sitasi WARISAN yang sah (ada di sumber)
    flaggedCitations: v.flagged,        // tak ditemukan di sumber -> dibuang dari `paragraph`
    verifiedQuotes: quoteCheck.verified.map(q => q.text),
    flaggedQuotes: quoteCheck.flagged.map(q => q.text),
    faithfulnessChecked: true,
    primarySource: usedIds.map((id) => ({ source_id: id, title: (titles[id] || {}).title || null,
                                          year: (titles[id] || {}).year || null })),
    note: note.trim()
  };
}

module.exports = { generate_paragraph_from_source };
