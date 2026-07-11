// scripts/publicUrl.js — resolusi URL publik add-in, satu sumber kebenaran.
//
// Dipakai server.js (route /manifest.xml + log) dan scripts/gen-manifest.js.
// Prioritas:
//   1. process.env.PUBLIC_URL            (mis. hasil deploy Railway)
//   2. process.env.RAILWAY_PUBLIC_DOMAIN (di-set otomatis oleh Railway; ditambah https://)
//   3. https://localhost:<PORT>          (development lokal — fallback default)
//
// Selalu tanpa trailing slash supaya "{{PUBLIC_URL}}/taskpane.html" rapi.

function normalize(u) {
  return String(u || "").trim().replace(/\/+$/, "");
}

function resolvePublicUrl(port) {
  const p = Number(port || process.env.FRIDA_PORT || 3001);
  if (process.env.PUBLIC_URL) return normalize(process.env.PUBLIC_URL);
  if (process.env.RAILWAY_PUBLIC_DOMAIN) {
    const d = process.env.RAILWAY_PUBLIC_DOMAIN.replace(/^https?:\/\//, "");
    return "https://" + normalize(d);
  }
  return "https://localhost:" + p;
}

function isLocalhost(url) {
  return /^https?:\/\/localhost\b/i.test(url || "");
}

module.exports = { resolvePublicUrl, isLocalhost, normalize };
