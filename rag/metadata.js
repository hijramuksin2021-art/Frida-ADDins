// rag/metadata.js — bangun metadata CSL untuk dokumen (sumber sitasi).
// Prioritas: DOI->Crossref (resmi). Tanpa DOI: tebakan lokal (confidence rendah) yang
// WAJIB dikonfirmasi/dikoreksi pengguna lewat UI sebelum dipakai menyitir.

const crossref = require("./crossref");

function titleCase(s) {
  return String(s || "").trim().replace(/\s+/g, " ")
    .split(" ").map((w) => w ? w[0].toUpperCase() + w.slice(1).toLowerCase() : w).join(" ");
}

// Tebak penulis dari nama file, mis. "Skripsi nurbaya bahrun.pdf" -> {family:Bahrun, given:Nurbaya}
function authorFromFilename(filename) {
  let s = String(filename || "").replace(/\.[^.]+$/, "");
  s = s.replace(/\b(skripsi|tesis|thesis|disertasi|jurnal|journal|paper|artikel|laporan|final|fix|revisi|docx?|pdf)\b/gi, " ");
  s = s.replace(/[_\-]+/g, " ").replace(/\d+/g, " ").replace(/\s+/g, " ").trim();
  const parts = s.split(" ").filter((w) => w.length > 1);
  if (parts.length < 2 || parts.length > 4) return null; // terlalu sedikit/banyak -> tak yakin
  const family = titleCase(parts[parts.length - 1]);
  const given = titleCase(parts.slice(0, -1).join(" "));
  return [{ family, given }];
}

function guessType(filename, text) {
  const hay = (String(filename || "") + " " + String(text || "").slice(0, 500)).toLowerCase();
  if (/\b(skripsi|tesis|thesis|disertasi|dissertation)\b/.test(hay)) return "thesis";
  if (/\b(prosiding|proceedings|conference|seminar)\b/.test(hay)) return "paper-conference";
  return "article-journal";
}

function guessInstitution(text) {
  const m = String(text || "").slice(0, 4000)
    .match(/\b(Universitas|University|Institut|Politeknik|Sekolah Tinggi)\s+[A-Z][A-Za-z'\s]{2,40}/);
  return m ? m[0].replace(/\s+/g, " ").trim() : null;
}

// build({ filename, text, parsedMeta }) -> { csl, confidence }
// parsedMeta = { title, year, doi } dari parse.js
async function build({ filename, text, parsedMeta }) {
  parsedMeta = parsedMeta || {};

  // 1) DOI -> Crossref (resmi)
  if (parsedMeta.doi) {
    const cr = await crossref.fetchByDoi(parsedMeta.doi);
    if (cr && cr.title) return { csl: cr, confidence: "high" };
  }

  // 2) tebakan lokal
  const type = guessType(filename, text);
  const csl = {
    type,
    title: parsedMeta.title || (filename || "").replace(/\.[^.]+$/, ""),
    author: authorFromFilename(filename) || [],
    issued: { year: parsedMeta.year || null },
    container: null,
    institution: type === "thesis" ? guessInstitution(text) : null,
    DOI: parsedMeta.doi || null,
    _source: "guess",
  };
  // confidence: ada penulis & tahun -> medium; selain itu low
  const confidence = (csl.author.length && csl.issued.year) ? "medium" : "low";
  return { csl, confidence };
}

module.exports = { build, authorFromFilename, guessType };
