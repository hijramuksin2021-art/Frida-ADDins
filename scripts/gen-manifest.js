// scripts/gen-manifest.js — hasilkan manifest.xml dari manifest.template.xml,
// mengganti token {{PUBLIC_URL}} dengan URL publik ter-resolusi (lihat publicUrl.js).
//
// Jalankan:  npm run manifest      (juga dipanggil otomatis di awal npm start)
// Untuk sideload ke Word memakai URL Railway:
//   PUBLIC_URL=https://<app>.up.railway.app  node scripts/gen-manifest.js

const fs = require("fs");
const path = require("path");
const { resolvePublicUrl, isLocalhost } = require("./publicUrl");

// Muat .env agar PUBLIC_URL bisa dibaca saat dijalankan mandiri.
(function loadDotEnv() {
  try {
    const p = path.join(__dirname, "..", ".env");
    if (!fs.existsSync(p)) return;
    for (const line of fs.readFileSync(p, "utf8").split(/\r?\n/)) {
      const m = line.match(/^\s*([\w.-]+)\s*=\s*(.*)\s*$/);
      if (!m || line.trim().startsWith("#")) continue;
      let v = m[2].trim();
      if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) v = v.slice(1, -1);
      if (!(m[1] in process.env)) process.env[m[1]] = v;
    }
  } catch (_) {}
})();

function renderManifest(publicUrl) {
  const tpl = fs.readFileSync(path.join(__dirname, "..", "manifest.template.xml"), "utf8");
  const notice = "FILE HASIL GENERATE — JANGAN edit langsung. Edit manifest.template.xml lalu jalankan `npm run manifest`. "
    + "URL di bawah berasal dari PUBLIC_URL = " + publicUrl;
  return tpl
    .replace(/@@GENERATED_NOTICE@@/g, notice)
    .replace(/\{\{PUBLIC_URL\}\}/g, publicUrl);
}

function writeManifest() {
  const publicUrl = resolvePublicUrl();
  const out = renderManifest(publicUrl);
  const dest = path.join(__dirname, "..", "manifest.xml");
  fs.writeFileSync(dest, out);
  return publicUrl;
}

// Ekspor untuk dipakai server.js (render dinamis tanpa menulis file).
module.exports = { renderManifest, writeManifest };

// Bila dijalankan langsung: tulis manifest.xml + info.
if (require.main === module) {
  const url = writeManifest();
  console.log("manifest.xml dibuat dengan PUBLIC_URL = " + url);
  if (isLocalhost(url)) {
    console.log("Catatan: masih localhost. Set PUBLIC_URL (mis. URL Railway) lalu jalankan ulang");
    console.log("         `npm run manifest` sebelum sideload manifest ke Word untuk akses publik.");
  }
}
