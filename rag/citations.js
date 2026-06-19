// rag/citations.js — ekstraksi & VERIFIKASI sitasi dalam teks (inti anti-halusinasi R2).
//
// Konsep sitasi primer vs sekunder (warisan):
//   - PRIMER  : klaim adalah pernyataan dokumen yang diunggah itu sendiri -> sitasi = dokumen itu.
//   - WARISAN : klaim dikutip dokumen dari sumber lain (ada sitasi in-text di teks sumber, mis.
//               "(Nair, 2012)") -> sitasi harus mengikuti sumber asli (Nair 2012), bukan dokumen.
//
// Verifikasi: setiap sitasi in-text di paragraf yang DIBUAT model harus benar-benar MUNCUL di
// chunk sumber (nama penulis + tahun ada di teks sumber). Jika tidak -> ditandai (kemungkinan
// dikarang) agar tidak dipercaya.

// Pisahkan surname pertama dari token penulis.
function firstSurname(s) {
  const tok = String(s || "").trim().split(/[\s,]+/)[0] || "";
  return tok.replace(/[^A-Za-zÀ-ÿ'`-]/g, "");
}

function escapeRe(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

// Ekstrak sitasi in-text gaya umum (parenthetical & naratif), EN + ID (dkk., dan).
// -> [{ raw, author, year }]
function extractCitations(text) {
  const out = [];
  const seen = new Set();
  const push = (raw, authorTok, year) => {
    const author = firstSurname(authorTok);
    if (!author || !year) return;
    const key = author.toLowerCase() + ":" + year;
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ raw: raw.trim(), author, year });
  };

  // (Author, 2012) / (Author & Other, 2012) / (Author et al., 2012) / (Author dkk., 2012)
  const paren = /\(([A-ZÀ-Ý][^()\d]{1,60}?),?\s*(\d{4}[a-z]?)\)/g;
  let m;
  while ((m = paren.exec(text))) push(m[0], m[1], m[2]);

  // Author (2012) / Author et al. (2012) / Author dkk. (2012)  [naratif]
  const narr = /([A-ZÀ-Ý][A-Za-zÀ-ÿ'`-]+(?:\s+(?:et al\.|dkk\.))?)\s*\((\d{4}[a-z]?)\)/g;
  while ((m = narr.exec(text))) push(m[0], m[1], m[2]);

  return out;
}

// Apakah sitasi {author, year} BENAR-BENAR ada di salah satu chunk (nama+tahun di teks sama)?
function citationInChunks(cite, chunks) {
  const re = new RegExp("\\b" + escapeRe(cite.author) + "", "i");
  return chunks.some((c) => {
    const t = (c && c.text) || c || "";
    return t.indexOf(cite.year) >= 0 && re.test(t);
  });
}

// Verifikasi semua sitasi pada paragraf terhadap chunk sumber.
// -> { citations:[...], verified:[raw], flagged:[raw] }
function verifyCitations(paragraph, chunks) {
  const cites = extractCitations(paragraph);
  const verified = [], flagged = [];
  cites.forEach((ct) => {
    (citationInChunks(ct, chunks) ? verified : flagged).push(ct.raw);
  });
  return { citations: cites, verified, flagged };
}

module.exports = { extractCitations, citationInChunks, verifyCitations, firstSurname };
