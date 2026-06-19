// rag/parse.js — ekstraksi teks dari PDF/DOCX/TXT (server, Node).
// Murni JS (tanpa native build): pdf-parse v2 (PDFParse) + mammoth.
// Mengembalikan { text, pages, meta } — meta hanya tebakan dasar (R0);
// metadata akurat (DOI/Crossref/CSL) menyusul di R3.

const mammoth = require("mammoth");
const { PDFParse } = require("pdf-parse");

// Buang penanda halaman yang disisipkan pdf-parse: "\n\n-- N of M --"
function stripPageMarkers(s) {
  return String(s || "").replace(/\n*-- \d+ of \d+ --\n*/g, "\n").trim();
}

async function parsePdf(buffer) {
  const p = new PDFParse({ data: buffer });
  const r = await p.getText();
  try { await p.destroy(); } catch (_) {}
  return {
    text: stripPageMarkers(r.text),
    pages: r.total || (Array.isArray(r.pages) ? r.pages.length : null),
  };
}

async function parseDocx(buffer) {
  const r = await mammoth.extractRawText({ buffer });
  return { text: String(r.value || "").trim(), pages: null };
}

function parseTxt(buffer) {
  return { text: buffer.toString("utf8").trim(), pages: null };
}

// Tebakan judul: baris non-kosong pertama yang "masuk akal" sbg judul.
function guessTitle(text, fallback) {
  const lines = String(text || "").split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  for (const l of lines.slice(0, 8)) {
    if (l.length >= 8 && l.length <= 200 && !/^(abstract|abstrak|http|doi|www\.)/i.test(l)) {
      return l;
    }
  }
  return fallback;
}

// Tebakan tahun: 19xx/20xx pertama di 2000 char awal.
function guessYear(text) {
  const m = String(text || "").slice(0, 2000).match(/\b(19|20)\d{2}\b/);
  return m ? Number(m[0]) : null;
}

// Tebakan DOI (utk verifikasi Crossref di R3).
function findDoi(text) {
  const m = String(text || "").match(/\b10\.\d{4,9}\/[^\s"<>]+/i);
  return m ? m[0].replace(/[.,;)]+$/, "") : null;
}

const EXT = {
  pdf: "application/pdf",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  txt: "text/plain",
};

async function parseByExt(ext, buffer, filename) {
  let out;
  if (ext === "pdf") out = await parsePdf(buffer);
  else if (ext === "docx") out = await parseDocx(buffer);
  else if (ext === "txt") out = parseTxt(buffer);
  else throw new Error("Tipe file tidak didukung: " + ext);

  const text = out.text || "";
  const meta = {
    title: guessTitle(text, filename.replace(/\.[^.]+$/, "")),
    year: guessYear(text),
    doi: findDoi(text),
    confidence: out.text ? (findDoi(text) ? "medium" : "low") : "low",
  };
  return { text, pages: out.pages, meta, chars: text.length };
}

module.exports = { parseByExt, parsePdf, parseDocx, parseTxt, EXT,
                   guessTitle, guessYear, findDoi };
