// rag/csl.js — formatter sitasi DETERMINISTIK dari metadata (CSL-JSON-ish).
// PRINSIP ANTI-HALUSINASI: sitasi TIDAK PERNAH ditulis LLM. Semua string sitasi (in-text &
// daftar pustaka) dibuat KODE ini dari record metadata terverifikasi. LLM hanya menyebut source_id.
//
// Bentuk metadata (subset CSL-JSON):
//   { type:"article-journal"|"thesis"|"book"|"chapter"|"webpage",
//     title, author:[{family,given}], issued:{year}, container, volume, issue, page,
//     publisher, place, institution, DOI, URL }
//
// Catatan: ini formatter ringan untuk kasus umum. Untuk fidelitas CSL penuh (semua edge case)
// upgrade ke citeproc-js + file CSL (interface metadata sengaja CSL-JSON agar mudah dipindah).

function authorsArray(meta) {
  const a = meta && meta.author;
  if (!Array.isArray(a)) return [];
  return a.filter((x) => x && (x.family || x.given || x.literal));
}
function family(a) { return a.family || a.literal || ""; }
function given(a) { return a.given || ""; }
function initials(given) {
  return String(given || "").trim().split(/[\s.]+/).filter(Boolean)
    .map((p) => p[0].toUpperCase() + ".").join(" ");
}
function year(meta) {
  const y = meta && meta.issued && meta.issued.year;
  return y ? String(y) : "n.d.";
}
function titleOf(meta) { return (meta && meta.title) || "Tanpa judul"; }

// ---------- IN-TEXT ----------
// style: APA7 | MLA | Chicago | Harvard | IEEE
// opts: { page, narrative, number } (number untuk IEEE)
function inText(meta, style, opts) {
  opts = opts || {};
  const auts = authorsArray(meta);
  const y = year(meta);
  const sur = surnameLabel(auts, style);

  switch (normStyle(style)) {
    case "IEEE":
      return "[" + (opts.number || 1) + "]";
    case "MLA": {
      const p = opts.page ? " " + String(opts.page).replace(/^p+\.?\s*/i, "") : "";
      return opts.narrative ? sur : "(" + sur + p + ")";
    }
    case "Chicago": {
      const p = opts.page ? ", " + String(opts.page).replace(/^p+\.?\s*/i, "") : "";
      return opts.narrative ? sur + " (" + y + ")" : "(" + sur + " " + y + p + ")";
    }
    case "Harvard": {
      const p = opts.page ? ", p. " + String(opts.page).replace(/^p+\.?\s*/i, "") : "";
      return opts.narrative ? sur + " (" + y + ")" : "(" + sur + ", " + y + p + ")";
    }
    case "APA7":
    default: {
      const p = opts.page ? ", p. " + String(opts.page).replace(/^p+\.?\s*/i, "") : "";
      return opts.narrative ? sur + " (" + y + ")" : "(" + sur + ", " + y + p + ")";
    }
  }
}

// Label nama untuk in-text (et al. rule sederhana).
function surnameLabel(auts, style) {
  if (!auts.length) return "Anonim";
  const s = normStyle(style);
  const fams = auts.map(family);
  if (fams.length === 1) return fams[0];
  if (s === "APA7" || s === "Harvard") {
    return fams.length >= 3 ? fams[0] + " et al." : fams[0] + " & " + fams[1];
  }
  if (s === "MLA") {
    return fams.length >= 3 ? fams[0] + " et al." : fams[0] + " and " + fams[1];
  }
  // Chicago
  return fams.length >= 4 ? fams[0] + " et al." : fams.join(", ").replace(/, ([^,]*)$/, " and $1");
}

// ---------- DAFTAR PUSTAKA (entry) ----------
function bibEntry(meta, style) {
  switch (normStyle(style)) {
    case "MLA": return mla(meta);
    case "Chicago": return chicago(meta);
    case "Harvard": return harvard(meta);
    case "IEEE": return ieee(meta);
    case "APA7":
    default: return apa7(meta);
  }
}

function apa7(m) {
  const auts = authorsArray(m);
  const names = auts.map((a) => family(a) + (given(a) ? ", " + initials(given(a)) : ""));
  const authStr = joinNames(names, "&");
  const y = year(m);
  let s = (authStr ? authStr + " " : "") + "(" + y + "). ";
  if (m.type === "thesis") {
    s += titleOf(m) + " [Skripsi/Tesis" + (m.institution ? ", " + m.institution : "") + "]. ";
    if (m.URL) s += m.URL;
  } else if (m.type === "book") {
    s += italic(titleOf(m)) + ". " + (m.publisher || "") + ".";
  } else { // article-journal default
    s += titleOf(m) + ". " + italic(m.container || "") +
      (m.volume ? ", " + italic(String(m.volume)) : "") +
      (m.issue ? "(" + m.issue + ")" : "") +
      (m.page ? ", " + m.page : "") + ".";
    if (m.DOI) s += " https://doi.org/" + m.DOI.replace(/^https?:\/\/doi\.org\//, "");
  }
  return s.trim();
}

function mla(m) {
  const auts = authorsArray(m);
  const lead = auts.length ? family(auts[0]) + (given(auts[0]) ? ", " + given(auts[0]) : "") : "";
  const rest = auts.slice(1).map((a) => given(a) + " " + family(a)).join(", ");
  const authStr = auts.length >= 3 ? family(auts[0]) + ", et al" : (rest ? lead + ", and " + rest : lead);
  const y = year(m);
  if (m.type === "book") return (authStr ? authStr + ". " : "") + italic(titleOf(m)) + ". " + (m.publisher || "") + ", " + y + ".";
  return (authStr ? authStr + ". " : "") + '"' + titleOf(m) + '." ' + italic(m.container || "") +
    (m.volume ? ", vol. " + m.volume : "") + (m.issue ? ", no. " + m.issue : "") +
    ", " + y + (m.page ? ", pp. " + m.page : "") + ".";
}

function chicago(m) {
  const auts = authorsArray(m);
  const lead = auts.length ? family(auts[0]) + (given(auts[0]) ? ", " + given(auts[0]) : "") : "";
  const rest = auts.slice(1).map((a) => given(a) + " " + family(a)).join(", and ");
  const authStr = rest ? lead + ", and " + rest : lead;
  const y = year(m);
  if (m.type === "book") return (authStr ? authStr + ". " : "") + y + ". " + italic(titleOf(m)) + ". " + (m.place ? m.place + ": " : "") + (m.publisher || "") + ".";
  return (authStr ? authStr + ". " : "") + y + '. "' + titleOf(m) + '." ' + italic(m.container || "") +
    (m.volume ? " " + m.volume : "") + (m.issue ? " (" + m.issue + ")" : "") +
    (m.page ? ": " + m.page : "") + ".";
}

function harvard(m) {
  const auts = authorsArray(m);
  const names = auts.map((a) => family(a) + (given(a) ? ", " + initials(given(a)) : ""));
  const authStr = joinNames(names, "and");
  const y = year(m);
  if (m.type === "book") return (authStr ? authStr + " " : "") + "(" + y + ") " + italic(titleOf(m)) + ". " + (m.place ? m.place + ": " : "") + (m.publisher || "") + ".";
  return (authStr ? authStr + " " : "") + "(" + y + ") '" + titleOf(m) + "', " + italic(m.container || "") +
    (m.volume ? ", " + m.volume : "") + (m.issue ? "(" + m.issue + ")" : "") +
    (m.page ? ", pp. " + m.page : "") + ".";
}

function ieee(m) {
  const auts = authorsArray(m);
  const names = auts.map((a) => (initials(given(a)) ? initials(given(a)) + " " : "") + family(a));
  const authStr = names.length ? (names.length > 6 ? names[0] + " et al." : names.join(", ")) : "";
  const y = year(m);
  if (m.type === "book") return (authStr ? authStr + ", " : "") + italic(titleOf(m)) + ". " + (m.place ? m.place + ": " : "") + (m.publisher || "") + ", " + y + ".";
  return (authStr ? authStr + ", " : "") + '"' + titleOf(m) + '," ' + italic(m.container || "") +
    (m.volume ? ", vol. " + m.volume : "") + (m.issue ? ", no. " + m.issue : "") +
    (m.page ? ", pp. " + m.page : "") + ", " + y + ".";
}

// ---------- util ----------
function joinNames(names, amp) {
  if (!names.length) return "";
  if (names.length === 1) return names[0];
  return names.slice(0, -1).join(", ") + ", " + amp + " " + names[names.length - 1];
}
function italic(s) { return s ? "*" + s + "*" : ""; } // penanda; klien render miring jika perlu
function normStyle(s) {
  s = String(s || "APA7").toUpperCase().replace(/[\s-]/g, "");
  if (s.indexOf("APA") === 0) return "APA7";
  if (s.indexOf("MLA") === 0) return "MLA";
  if (s.indexOf("CHICAGO") === 0) return "Chicago";
  if (s.indexOf("HARVARD") === 0) return "Harvard";
  if (s.indexOf("IEEE") === 0) return "IEEE";
  return "APA7";
}

const STYLES = ["APA7", "MLA", "Chicago", "Harvard", "IEEE"];

module.exports = { inText, bibEntry, normStyle, STYLES, authorsArray };
