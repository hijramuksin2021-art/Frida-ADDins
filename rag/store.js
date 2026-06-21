// rag/store.js — penyimpanan KB berbasis FILE (R0; tanpa native dep).
// Layout:
//   data/sources/index.json            -> array ringkasan dokumen (tanpa teks)
//   data/sources/<id>.json             -> { ...summary, text }  (teks penuh)
// R1 mengganti ini dgn SQLite+sqlite-vec (interface dijaga mirip).

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const ROOT = path.join(__dirname, "..", "data", "sources");
const INDEX = path.join(ROOT, "index.json");

function ensureDir() { fs.mkdirSync(ROOT, { recursive: true }); }

function readIndex() {
  try { return JSON.parse(fs.readFileSync(INDEX, "utf8")); }
  catch (_) { return []; }
}
function writeIndex(list) {
  ensureDir();
  fs.writeFileSync(INDEX, JSON.stringify(list, null, 2));
}

function sha256(buf) { return crypto.createHash("sha256").update(buf).digest("hex"); }
function newId() { return "src_" + crypto.randomBytes(6).toString("hex"); }

// Cari dokumen dgn hash sama (dedup).
function findByHash(hash) {
  return readIndex().find((d) => d.hash === hash) || null;
}

// Simpan dokumen baru. doc = {filename,mime,ext,hash,title,year,doi,confidence,pages,chars,text,workspace}
function save(doc) {
  ensureDir();
  const id = newId();
  const summary = {
    id,
    filename: doc.filename,
    ext: doc.ext,
    mime: doc.mime,
    hash: doc.hash,
    title: doc.title,
    year: doc.year || null,
    doi: doc.doi || null,
    confidence: doc.confidence || "low",
    csl: doc.csl || null,        // metadata sitasi (CSL-JSON-ish)
    pages: doc.pages || null,
    chars: doc.chars || (doc.text ? doc.text.length : 0),
    workspace: doc.workspace || "default",
    uploaded_at: new Date().toISOString(),
    status: "ready",
  };
  fs.writeFileSync(path.join(ROOT, id + ".json"),
    JSON.stringify(Object.assign({}, summary, { text: doc.text || "" }), null, 2));
  const list = readIndex();
  list.push(summary);
  writeIndex(list);
  return summary;
}

function list(workspace) {
  const all = readIndex();
  return workspace ? all.filter((d) => d.workspace === workspace) : all;
}

function get(id) {
  try { return JSON.parse(fs.readFileSync(path.join(ROOT, id + ".json"), "utf8")); }
  catch (_) { return null; }
}

// Perbarui metadata CSL (koreksi pengguna). patch = sebagian field CSL + opsional title/year.
function updateMetadata(id, csl) {
  const full = get(id);
  if (!full) return null;
  full.csl = Object.assign({}, full.csl, csl);
  if (full.csl.title) full.title = full.csl.title;
  if (full.csl.issued && full.csl.issued.year) full.year = full.csl.issued.year;
  full.confidence = "user";
  fs.writeFileSync(path.join(ROOT, id + ".json"), JSON.stringify(full, null, 2));
  // sinkron ke index
  const list0 = readIndex();
  const i = list0.findIndex((d) => d.id === id);
  if (i >= 0) {
    list0[i].csl = full.csl; list0[i].title = full.title;
    list0[i].year = full.year; list0[i].confidence = "user";
    writeIndex(list0);
  }
  return full.csl;
}

function remove(id) {
  const list0 = readIndex();
  const next = list0.filter((d) => d.id !== id);
  if (next.length === list0.length) return false;
  writeIndex(next);
  try { fs.unlinkSync(path.join(ROOT, id + ".json")); } catch (_) {}
  return true;
}

module.exports = { save, list, get, remove, findByHash, sha256, updateMetadata, ROOT };
