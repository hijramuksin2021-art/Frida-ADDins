// rag/chunk.js — pemecah teks jadi chunk untuk retrieval.
// Berbasis kalimat dgn target ~karakter (1 token ≈ 4 char). Overlap menjaga konteks
// di batas chunk. Kalimat super-panjang (tanpa tanda baca) dipecah keras.

function splitSentences(text) {
  const clean = String(text || "").replace(/\r/g, "");
  // pisah pada akhir kalimat ATAU baris kosong (paragraf)
  return clean
    .split(/(?<=[.!?])\s+|\n{2,}/)
    .map((s) => s.trim())
    .filter(Boolean);
}

function hardSplit(s, max) {
  const out = [];
  for (let i = 0; i < s.length; i += max) out.push(s.slice(i, i + max));
  return out;
}

// chunkText(text, {target, overlap}) -> [{ ordinal, text }]
function chunkText(text, opts) {
  opts = opts || {};
  const target = opts.target || 1200;   // ~300 token
  const overlap = opts.overlap || 200;  // ~50 token

  const sentences = [];
  splitSentences(text).forEach((s) => {
    if (s.length > target) hardSplit(s, target).forEach((p) => sentences.push(p));
    else sentences.push(s);
  });

  const chunks = [];
  let cur = "";
  for (const s of sentences) {
    if (cur.length + s.length + 1 > target && cur.length > 0) {
      chunks.push(cur.trim());
      const tail = cur.slice(Math.max(0, cur.length - overlap));
      cur = (tail + " " + s).trim();
    } else {
      cur = cur ? cur + " " + s : s;
    }
  }
  if (cur.trim()) chunks.push(cur.trim());

  return chunks.map((t, i) => ({ ordinal: i, text: t }));
}

module.exports = { chunkText, splitSentences };
