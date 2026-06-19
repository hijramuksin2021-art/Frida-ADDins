// rag/ingest.js — orkestrasi unggah → parse → metadata dasar → simpan (R0).
// Input dari endpoint: { filename, mime, dataBase64, workspace }.
// Chunk + embed dilakukan di R1.

const parse = require("./parse");
const store = require("./store");

const MAX_BYTES = 25 * 1024 * 1024; // 25 MB

function extOf(filename, mime) {
  const m = /\.([a-z0-9]+)$/i.exec(filename || "");
  let ext = m ? m[1].toLowerCase() : "";
  if (!ext && mime) {
    if (/pdf/.test(mime)) ext = "pdf";
    else if (/wordprocessingml|msword/.test(mime)) ext = "docx";
    else if (/text\/plain/.test(mime)) ext = "txt";
  }
  return ext;
}

async function ingestUpload({ filename, mime, dataBase64, workspace }) {
  if (!filename || !dataBase64) throw new Error("filename & dataBase64 wajib diisi");
  const ext = extOf(filename, mime);
  if (!["pdf", "docx", "txt"].includes(ext)) {
    throw new Error("Tipe tidak didukung (hanya pdf, docx, txt): " + (ext || mime || "?"));
  }
  const buffer = Buffer.from(dataBase64, "base64");
  if (buffer.length === 0) throw new Error("File kosong");
  if (buffer.length > MAX_BYTES) throw new Error("File melebihi 25 MB");

  const hash = store.sha256(buffer);
  const existing = store.findByHash(hash);
  if (existing) return { document: existing, duplicate: true };

  const parsed = await parse.parseByExt(ext, buffer, filename);
  if (!parsed.text || parsed.text.length < 1) {
    throw new Error("Tidak ada teks yang bisa diekstrak (mungkin PDF hasil scan/gambar).");
  }

  const doc = store.save({
    filename, ext, mime: mime || parse.EXT[ext], hash,
    title: parsed.meta.title, year: parsed.meta.year, doi: parsed.meta.doi,
    confidence: parsed.meta.confidence, pages: parsed.pages, chars: parsed.chars,
    text: parsed.text, workspace: workspace || "default",
  });
  return { document: doc, duplicate: false };
}

module.exports = { ingestUpload, extOf };
