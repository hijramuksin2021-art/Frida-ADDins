// rag/crossref.js — ambil metadata RESMI dari DOI lewat Crossref (sumber kebenaran sitasi).
// Hanya DOI yang dikirim keluar (bukan isi dokumen). Gagal/timeout -> null (pakai tebakan lokal).

function mapCrossref(msg) {
  if (!msg) return null;
  const issuedParts = msg.issued && msg.issued["date-parts"] && msg.issued["date-parts"][0];
  const typeMap = {
    "journal-article": "article-journal",
    "book": "book",
    "book-chapter": "chapter",
    "proceedings-article": "paper-conference",
    "dissertation": "thesis",
  };
  return {
    type: typeMap[msg.type] || "article-journal",
    title: Array.isArray(msg.title) ? msg.title[0] : msg.title || "",
    author: (msg.author || []).map((a) => ({ family: a.family || "", given: a.given || "" })),
    issued: { year: issuedParts ? issuedParts[0] : null },
    container: Array.isArray(msg["container-title"]) ? msg["container-title"][0] : msg["container-title"] || "",
    volume: msg.volume || null,
    issue: msg.issue || null,
    page: msg.page || null,
    publisher: msg.publisher || null,
    DOI: msg.DOI || null,
    URL: msg.URL || null,
    _source: "crossref",
  };
}

async function fetchByDoi(doi) {
  if (!doi) return null;
  const clean = String(doi).replace(/^https?:\/\/(dx\.)?doi\.org\//i, "").trim();
  const url = "https://api.crossref.org/works/" + encodeURIComponent(clean);
  try {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 8000);
    const resp = await fetch(url, {
      headers: { "User-Agent": "FRIDA-ResearchCopilot/1.0 (mailto:local@frida)" },
      signal: ctrl.signal,
    });
    clearTimeout(t);
    if (!resp.ok) return null;
    const data = await resp.json();
    return mapCrossref(data && data.message);
  } catch (_) { return null; }
}

module.exports = { fetchByDoi, mapCrossref };
