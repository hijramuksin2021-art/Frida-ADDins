// rag/cite.js — jembatan antara metadata tersimpan dan formatter (csl.js).
// Dipakai endpoint /api/sources/cite & /bibliography. Sitasi SELALU dari metadata
// terverifikasi (store), bukan dari LLM.

const store = require("./store");
const csl = require("./csl");

function inTextFor(source_id, style, opts) {
  const doc = store.get(source_id);
  if (!doc || !doc.csl) return null;
  return csl.inText(doc.csl, style, opts || {});
}

function entryFor(source_id, style) {
  const doc = store.get(source_id);
  if (!doc || !doc.csl) return null;
  return csl.bibEntry(doc.csl, style);
}

// Daftar pustaka utk beberapa sumber. Author-date -> urut alfabet; IEEE -> urut input.
function bibliography(source_ids, style) {
  const ids = (source_ids && source_ids.length)
    ? source_ids : store.list().map((d) => d.id);
  let entries = ids.map((id) => ({ source_id: id, text: entryFor(id, style) }))
    .filter((e) => e.text);
  if (csl.normStyle(style) !== "IEEE") {
    entries.sort((a, b) => a.text.localeCompare(b.text));
  }
  return entries;
}

module.exports = { inTextFor, entryFor, bibliography };
