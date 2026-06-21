// rag/aliases.js — resolve sumber dari nama alami (tanpa embedding; scoring kata kunci).
// "jurnal Hijra" / "Nair 2012" / "paper agroforestri" → {doc, score}[]
//
// Scoring per token query:
//   +3  token cocok persis dgn kata kunci metadata (judul/penulis/filename/tahun)
//   +1  partial: kata kunci metadata mengandung token (mis. "agro" → "agroforestri")
// Bonus: +2 bila tahun cocok; +1 untuk setiap token judul yang cocok vs token query.
// Kembalikan array terurut skor turun; skor 0 disaring.

const store = require("./store");

// Ekstrak set kata kunci dari satu ringkasan dokumen (index entry).
function docKeywords(doc) {
  const kw = new Set();
  const tokenize = (s) =>
    String(s || "")
      .toLowerCase()
      .split(/[\s\-_.,;:()\[\]\/]+/)
      .filter((w) => w.length > 1);

  tokenize(doc.title).forEach((w) => kw.add(w));
  tokenize(doc.filename).forEach((w) => kw.add(w));
  if (doc.year) kw.add(String(doc.year));

  const csl = doc.csl || {};
  tokenize(csl.container).forEach((w) => kw.add(w));
  (csl.author || []).forEach((a) => {
    tokenize(a.family).forEach((w) => kw.add(w));
    tokenize(a.given).forEach((w) => kw.add(w));
  });
  return kw;
}

// Hitung skor kecocokan satu dokumen terhadap array token query.
function scoreDoc(doc, queryTokens) {
  const kw = docKeywords(doc);
  let score = 0;

  for (const qt of queryTokens) {
    if (kw.has(qt)) {
      score += 3; // cocok persis
    } else {
      for (const k of kw) {
        if (k.includes(qt) && k !== qt) { score += 1; break; }
      }
    }
  }

  // Bonus untuk tahun eksplisit
  if (doc.year && queryTokens.includes(String(doc.year))) score += 2;

  // Bonus: penulis pertama cocok (sangat kuat)
  const csl = doc.csl || {};
  const firstAuthorFamily = ((csl.author || [])[0] || {}).family || "";
  if (
    firstAuthorFamily &&
    queryTokens.some((qt) =>
      firstAuthorFamily.toLowerCase().includes(qt) ||
      qt.includes(firstAuthorFamily.toLowerCase())
    )
  ) {
    score += 2;
  }

  return score;
}

// Tokenisasi query pengguna.
function tokenize(query) {
  return String(query || "")
    .toLowerCase()
    .split(/[\s\-_.,;:()\[\]\/]+/)
    .filter((w) => w.length > 1);
}

/**
 * Resolve sumber dari query nama alami.
 * @param {string} query  — mis. "jurnal Hijra", "Nair 2012", "paper agroforestri"
 * @param {string|null} workspace
 * @returns {{ doc, score }[]} urut skor turun, hanya skor > 0
 */
function resolveSource(query, workspace) {
  const tokens = tokenize(query);
  if (!tokens.length) return [];

  const docs = store.list(workspace || null);
  return docs
    .map((doc) => ({ doc, score: scoreDoc(doc, tokens) }))
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score);
}

/**
 * Cari kandidat terbaik dan kembalikan sebagai tool result.
 * maxResults: maks berapa sumber dikembalikan (default 3).
 */
function resolveSourceTool({ query, workspace, maxResults }) {
  if (!query) return { error: "query wajib diisi" };
  const results = resolveSource(query, workspace);
  if (!results.length) {
    return {
      matches: [],
      note: "Tidak ada sumber yang cocok dengan '" + query + "'. Cek ejaan atau unggah sumber terlebih dahulu.",
    };
  }
  const max = Math.max(1, Number(maxResults) || 3);
  return {
    matches: results.slice(0, max).map((r) => ({
      source_id: r.doc.id,
      title: r.doc.title || r.doc.filename,
      year: r.doc.year || null,
      confidence: r.doc.confidence || null,
      score: r.score,
    })),
    best_id: results[0].doc.id,
    best_title: results[0].doc.title || results[0].doc.filename,
  };
}

module.exports = { resolveSource, resolveSourceTool, tokenize, docKeywords };
