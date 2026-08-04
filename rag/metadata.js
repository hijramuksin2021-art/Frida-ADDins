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

// Tebak penulis dari isi teks (baris 2-6)
function guessAuthorFromText(text) {
  const lines = String(text || "").split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  for (let i = 1; i < Math.min(8, lines.length); i++) {
    const l = lines[i];
    if (l.length > 150 || /^(abstract|abstrak|http|doi|www\.)/i.test(l) || /universitas|institut|fakultas|departemen|program|studi/i.test(l)) continue;
    const names = l.split(/,|\bdan\b|&/i).map(n => n.replace(/[\d*]/g, "").trim()).filter(n => n.length > 2);
    if (names.length > 0 && names.length <= 5) {
      const authors = [];
      let valid = true;
      for (const n of names) {
        const words = n.split(/\s+/);
        if (words.length === 0 || words.length > 4) valid = false;
        if (valid) {
          authors.push({
            family: titleCase(words[words.length - 1]),
            given: titleCase(words.slice(0, -1).join(" "))
          });
        }
      }
      if (valid && authors.length > 0) return authors;
    }
  }
  return null;
}

// build({ filename, text, parsedMeta }) -> { csl, confidence }
// parsedMeta = { title, year, doi } dari parse.js
async function build({ filename, text, parsedMeta }) {
  parsedMeta = parsedMeta || {};
  const type = guessType(filename, text);

  // 1) DOI -> Crossref (resmi) (kecuali untuk skripsi/tesis yang rawan salah-ambil DOI dari daftar pustaka)
  if (parsedMeta.doi && type !== "thesis") {
    const cr = await crossref.fetchByDoi(parsedMeta.doi);
    if (cr && cr.title) return { csl: cr, confidence: "high" };
  }

  // 2) tebakan lokal
  let confidence = "medium";
  let author = guessAuthorFromText(text);
  if (!author) {
    author = authorFromFilename(filename) || [];
    confidence = "low"; // Jika jatuh ke fallback nama file, selalu low
  }
  if (!parsedMeta.year) confidence = "low";

  const csl = {
    type,
    title: parsedMeta.title || (filename || "").replace(/\.[^.]+$/, ""),
    author: author,
    issued: { year: parsedMeta.year || null },
    container: null,
    institution: type === "thesis" ? guessInstitution(text) : null,
    DOI: parsedMeta.doi || null,
    _source: "guess",
  };
  return { csl, confidence };
}

module.exports = { build, authorFromFilename, guessType };
