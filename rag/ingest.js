// rag/ingest.js — orkestrasi unggah → parse → metadata dasar → simpan (R0).
// Input dari endpoint: { filename, mime, dataBase64, workspace }.
// Chunk + embed dilakukan di R1.

const parse = require("./parse");
const store = require("./store");
const chunker = require("./chunk");
const vectors = require("./vectors");
const embeddings = require("./embeddings");
const metadata = require("./metadata");

const MAX_BYTES = 25 * 1024 * 1024; // 25 MB
const EMBED_BATCH = 32;

// Chunk + embed teks dokumen -> simpan vektor. Dipakai saat upload & reindex.
async function indexDocument(docId, text) {
  const pieces = chunker.chunkText(text);
  if (!pieces.length) return { numChunks: 0 };
  const out = [];
  for (let i = 0; i < pieces.length; i += EMBED_BATCH) {
    const batch = pieces.slice(i, i + EMBED_BATCH);
    const vecs = await embeddings.embed(batch.map((p) => p.text));
    batch.forEach((p, j) => {
      out.push({ chunk_id: docId + ":" + p.ordinal, ordinal: p.ordinal,
                 text: p.text, embedding: vecs[j] });
    });
  }
  vectors.saveChunks(docId, out);
  return { numChunks: out.length };
}

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

  // Metadata sitasi (CSL): DOI->Crossref bila ada, else tebakan lokal.
  let meta;
  try { meta = await metadata.build({ filename, text: parsed.text, parsedMeta: parsed.meta }); }
  catch (_) { meta = { csl: null, confidence: parsed.meta.confidence }; }

  const doc = store.save({
    filename, ext, mime: mime || parse.EXT[ext], hash,
    title: (meta.csl && meta.csl.title) || parsed.meta.title,
    year: (meta.csl && meta.csl.issued && meta.csl.issued.year) || parsed.meta.year,
    doi: parsed.meta.doi,
    confidence: meta.confidence, csl: meta.csl,
    pages: parsed.pages, chars: parsed.chars,
    text: parsed.text, workspace: workspace || "default",
  });

  // Chunk + embed (boleh gagal: dok tetap tersimpan, bisa di-reindex nanti).
  let indexed = { numChunks: 0 };
  let indexError = null;
  try { indexed = await indexDocument(doc.id, parsed.text); }
  catch (e) { indexError = String(e.message || e); }

  return { document: doc, duplicate: false,
           numChunks: indexed.numChunks, indexError };
}

// Reindex dokumen yang belum punya vektor (mis. diunggah sebelum R1).
async function reindexAll() {
  const result = [];
  for (const s of store.list()) {
    if (vectors.hasChunks(s.id)) { result.push({ id: s.id, skipped: true }); continue; }
    const full = store.get(s.id);
    if (!full || !full.text) { result.push({ id: s.id, error: "tak ada teks" }); continue; }
    try {
      const r = await indexDocument(s.id, full.text);
      result.push({ id: s.id, numChunks: r.numChunks });
    } catch (e) {
      result.push({ id: s.id, error: String(e.message || e) });
    }
  }
  return result;
}

module.exports = { ingestUpload, indexDocument, reindexAll, extOf };
