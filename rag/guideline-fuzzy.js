// rag/guideline-fuzzy.js — deteksi dan aktivasi guideline dari chat user message.
// Helper: cari guideline yang disebutkan user secara eksplisit dalam pesan
// (mis. "perbaiki sesuai pedoman Fakultas Pertanian Universitas Khairun").

const fs = require("fs");
const path = require("path");

const GUIDELINES_DIR = path.join(__dirname, "guidelines");

// Tokenisasi string untuk fuzzy matching
function tokenize(str) {
  return String(str || "")
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean);
}

// Hitung kesamaan antar dua array token (Jaccard similarity)
function jaccardSimilarity(tokens1, tokens2) {
  const set1 = new Set(tokens1);
  const set2 = new Set(tokens2);
  const intersection = [...set1].filter(t => set2.has(t)).length;
  const union = new Set([...set1, ...set2]).size;
  return union === 0 ? 0 : intersection / union;
}

// Load semua guideline available
function loadAllGuidelines() {
  try {
    const files = fs.readdirSync(GUIDELINES_DIR).filter(f => f.endsWith(".json"));
    return files.map(f => {
      try {
        const content = JSON.parse(fs.readFileSync(path.join(GUIDELINES_DIR, f), "utf8"));
        return {
          id: content.id,
          nama: content.nama || "",
          fakultas: content.fakultas || "",
          universitas: content.universitas || "",
          keywords: tokenize(
            (content.nama || "") + " " + (content.fakultas || "") + " " + (content.universitas || "")
          ),
        };
      } catch (_) {
        return null;
      }
    }).filter(Boolean);
  } catch (_) {
    return [];
  }
}

// Cari guideline dari chat message user — return best match atau null
function detectGuidelineFromMessage(message) {
  const msg = String(message || "").toLowerCase();

  // Quick check: apakah message mengandung keyword seprti "pedoman", "guidelines", "panduan"
  if (!/pedoman|panduan|guideline|guideline penulisan|format penulisan/.test(msg)) {
    return null;
  }

  const msgTokens = tokenize(msg);
  const allGuidelines = loadAllGuidelines();

  let bestMatch = null;
  let bestScore = 0;

  allGuidelines.forEach((gl) => {
    const score = jaccardSimilarity(msgTokens, gl.keywords);
    if (score > bestScore) {
      bestScore = score;
      bestMatch = gl;
    }
  });

  // Hanya return jika skor cukup tinggi (> 0.2)
  return bestScore > 0.2 ? bestMatch : null;
}

module.exports = { detectGuidelineFromMessage, loadAllGuidelines };
