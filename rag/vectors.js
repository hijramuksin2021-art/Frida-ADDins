// rag/vectors.js — vektor chunk berbasis FILE + pencarian cosine brute-force (R1).
// Layout: data/sources/<id>.chunks.json = [{ chunk_id, ordinal, text, embedding:[...] }]
// Skala: brute-force cukup utk puluhan ribu chunk (cosine = dot product, vektor
// sudah dinormalisasi). Untuk skala lebih besar/multi-user -> sqlite-vec / pgvector.

const fs = require("fs");
const path = require("path");
const { ROOT } = require("./store");

function chunkFile(id) { return path.join(ROOT, id + ".chunks.json"); }
function hasChunks(id) { return fs.existsSync(chunkFile(id)); }
function saveChunks(id, chunks) {
  fs.mkdirSync(ROOT, { recursive: true });
  fs.writeFileSync(chunkFile(id), JSON.stringify(chunks));
}
function loadChunks(id) {
  try { return JSON.parse(fs.readFileSync(chunkFile(id), "utf8")); }
  catch (_) { return null; }
}
function removeChunks(id) { try { fs.unlinkSync(chunkFile(id)); } catch (_) {} }

function dot(a, b) {
  let s = 0;
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) s += a[i] * b[i];
  return s;
}

// search(queryVec, docIds[], {k, minScore}) -> [{document_id, chunk_id, ordinal, text, score}]
function search(queryVec, docIds, opts) {
  opts = opts || {};
  const k = opts.k || 8;
  const minScore = opts.minScore != null ? opts.minScore : 0.25;
  const hits = [];
  for (const id of docIds) {
    const chunks = loadChunks(id);
    if (!chunks) continue;
    for (const c of chunks) {
      if (!c.embedding) continue;
      const score = dot(queryVec, c.embedding);
      hits.push({ document_id: id, chunk_id: c.chunk_id, ordinal: c.ordinal, text: c.text, score });
    }
  }
  hits.sort((a, b) => b.score - a.score);
  return hits.filter((h) => h.score >= minScore).slice(0, k);
}

module.exports = { chunkFile, hasChunks, saveChunks, loadChunks, removeChunks, search };
