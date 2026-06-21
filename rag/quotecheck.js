// rag/quotecheck.js — verifikasi kutipan dalam paragraf yang dihasilkan (R6).
// Ketika paragraf berisi teks yang dikutip (dalam tanda petik), verifikasi bahwa
// kutipan tersebut benar-benar muncul di source chunks.

const levenshtein = (a, b) => {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i - 1] === b[j - 1]
        ? dp[i - 1][j - 1]
        : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
    }
  }
  return dp[m][n];
};

// Hitung similarity dua string (0-1, dimana 1 = identik)
function stringSimilarity(a, b) {
  if (!a || !b) return 0;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  const dist = levenshtein(a.toLowerCase(), b.toLowerCase());
  return 1 - (dist / maxLen);
}

// Ekstrak kutipan dari paragraf (teks dalam tanda petik)
// Format: "..." atau '...'
function extractQuotes(text) {
  const out = [];
  const seen = new Set();

  // Double quotes: "..."
  const doubleQuoteRegex = /"([^"]{5,200})"/g;
  let m;
  while ((m = doubleQuoteRegex.exec(text))) {
    const quote = m[1].trim();
    if (quote && !seen.has(quote)) {
      seen.add(quote);
      out.push({ raw: m[0], text: quote, type: "double" });
    }
  }

  // Single quotes: '...'
  const singleQuoteRegex = /'([^']{5,200})'/g;
  while ((m = singleQuoteRegex.exec(text))) {
    const quote = m[1].trim();
    if (quote && !seen.has(quote)) {
      seen.add(quote);
      out.push({ raw: m[0], text: quote, type: "single" });
    }
  }

  return out;
}

// Cek apakah kutipan ada di salah satu chunk (exact match atau fuzzy match > 0.85)
function quoteInChunks(quote, chunks, threshold = 0.85) {
  const quoteText = quote.text.toLowerCase();

  for (const chunk of chunks) {
    const chunkText = (chunk && chunk.text) ? chunk.text.toLowerCase() : "";

    // Exact match (case-insensitive)
    if (chunkText.includes(quoteText)) {
      return { found: true, method: "exact", chunk_id: chunk.id, source_id: chunk.source_id };
    }

    // Fuzzy match: cari substring terpanjang yang cocok
    for (let start = 0; start < chunkText.length - 10; start++) {
      for (let end = start + 10; end <= chunkText.length; end++) {
        const substr = chunkText.substring(start, end);
        const sim = stringSimilarity(quoteText, substr);
        if (sim >= threshold) {
          return { found: true, method: "fuzzy", similarity: sim.toFixed(2), chunk_id: chunk.id, source_id: chunk.source_id };
        }
      }
    }
  }

  return { found: false, method: null };
}

// Verifikasi semua kutipan dalam paragraf terhadap chunks sumber
// -> { quotes: [...], verified: [{...}], flagged: [{...}] }
function verifyQuotes(paragraph, chunks) {
  const quotes = extractQuotes(paragraph);
  const verified = [], flagged = [];

  quotes.forEach((q) => {
    const result = quoteInChunks(q, chunks);
    if (result.found) {
      verified.push({ ...q, ...result });
    } else {
      flagged.push(q);
    }
  });

  return { quotes, verified, flagged };
}

module.exports = { extractQuotes, quoteInChunks, verifyQuotes, stringSimilarity };
