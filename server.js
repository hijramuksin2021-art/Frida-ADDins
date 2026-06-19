// Server lokal untuk add-in Claude di Word.
// Tugas:
//   1) Menyajikan file add-in (taskpane.html dll) lewat HTTPS  -> Word butuh HTTPS.
//   2) Proxy ke provider (aerolink) sambil menyisipkan API key  -> key tidak pernah
//      masuk ke dokumen / ke kode yang berjalan di Word.
//
// Jalankan:  npm start   (setelah `npm install` dan `npm run cert`)

const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

// Registry tool (Fase 1) — sumber kebenaran schema yang dikirim ke LLM.
// Endpoint agentic yang memakainya menyusul di Fase 2; di sini cukup dimuat & dilaporkan.
const { SCHEMAS: TOOL_SCHEMAS } = require("./tools/schemas");

// ---- konfigurasi: env DULU, lalu config.json sbg fallback (nilai non-rahasia) ----
// API key TIDAK boleh disimpan di config.json yang ter-commit. Taruh di .env / env OS.
// Loader .env mini (tanpa dependency tambahan): KEY=VALUE per baris, # = komentar.
function loadDotEnv() {
  const p = path.join(__dirname, ".env");
  if (!fs.existsSync(p)) return;
  for (const line of fs.readFileSync(p, "utf8").split(/\r?\n/)) {
    const m = line.match(/^\s*([\w.-]+)\s*=\s*(.*)\s*$/);
    if (!m || line.trim().startsWith("#")) continue;
    let val = m[2].trim();
    if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
      val = val.slice(1, -1);
    }
    if (!(m[1] in process.env)) process.env[m[1]] = val; // env OS menang atas .env
  }
}
loadDotEnv();

// config.json sekarang opsional & hanya untuk nilai non-rahasia (port, model, baseUrl).
let fileCfg = {};
try { fileCfg = JSON.parse(fs.readFileSync(path.join(__dirname, "config.json"), "utf8")); }
catch (_) { /* boleh tidak ada */ }

const cfg = {
  apiKey:    process.env.AERO_API_KEY    || fileCfg.apiKey,
  baseUrl:   process.env.AERO_BASE_URL   || fileCfg.baseUrl || "https://capi.aerolink.lat/",
  model:     process.env.FRIDA_MODEL     || fileCfg.model   || "claude-opus-4-8",
  maxTokens: Number(process.env.FRIDA_MAX_TOKENS || fileCfg.maxTokens || 8000),
  port:      Number(process.env.FRIDA_PORT || fileCfg.port || 3001),
};

if (!cfg.apiKey) {
  console.error("API key belum di-set. Buat file .env berisi:  AERO_API_KEY=...");
  console.error("(lihat .env.example). Jangan menaruh key di config.json yang ter-commit.");
  process.exit(1);
}
if (fileCfg.apiKey) {
  console.warn("PERINGATAN: config.json masih memuat apiKey. Pindahkan ke .env lalu hapus dari config.json.");
}

const PORT = cfg.port;

// ---- file statis yang boleh disajikan ----
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".xml": "text/xml",
  ".json": "application/json; charset=utf-8",
};

function serveStatic(req, res) {
  let urlPath = decodeURIComponent(req.url.split("?")[0]);
  if (urlPath === "/") urlPath = "/taskpane.html";
  const filePath = path.resolve(__dirname, "." + path.posix.normalize(urlPath));
  // jangan biarkan keluar dari folder proyek (cek batas dgn pemisah path, bukan prefix string)
  const root = __dirname + path.sep;
  if (filePath !== __dirname && !filePath.startsWith(root)) {
    res.writeHead(403); res.end("Forbidden"); return;
  }
  // jangan sajikan file rahasia walau diminta langsung
  if (/^(\.env|config\.json)$/i.test(path.basename(filePath))) {
    res.writeHead(403); res.end("Forbidden"); return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) { res.writeHead(404); res.end("Not found"); return; }
    res.writeHead(200, { "Content-Type": MIME[path.extname(filePath)] || "application/octet-stream" });
    res.end(data);
  });
}

// Definisi "tool" untuk memaksa model menjawab dalam JSON terstruktur.
// Cara ini jauh lebih andal daripada meminta model menulis JSON sebagai teks.
const EDIT_TOOL = {
  name: "apply_changes",
  description:
    "Terapkan perubahan ke dokumen Word. Pilih aksi yang sesuai instruksi: " +
    "memperbaiki paragraf (replace), menambah/menyisipkan paragraf baru (insertAfter / append), " +
    "atau merapikan sel tabel (tableOps). Sertakan hanya perubahan yang benar-benar diperlukan.",
  input_schema: {
    type: "object",
    properties: {
      summary: { type: "string", description: "Ringkasan singkat perubahan dalam Bahasa Indonesia." },
      paragraphOps: {
        type: "array",
        description: "Perubahan pada paragraf.",
        items: {
          type: "object",
          properties: {
            action: {
              type: "string",
              enum: ["replace", "insertAfter", "append"],
              description:
                "replace = ganti isi paragraf indeks i; " +
                "insertAfter = sisipkan paragraf BARU setelah paragraf indeks i; " +
                "append = tambahkan paragraf BARU di akhir dokumen.",
            },
            i: { type: "integer", description: "Indeks paragraf target (untuk replace & insertAfter). Abaikan untuk append." },
            newText: {
              type: "string",
              description: "Teks baru. Untuk insertAfter/append boleh beberapa paragraf, pisahkan tiap paragraf dengan baris baru (\\n).",
            },
            reason: { type: "string", description: "Alasan singkat." },
          },
          required: ["action", "newText"],
        },
      },
      tableOps: {
        type: "array",
        description: "Perubahan isi sel tabel yang sedang diseleksi (jika ada). Baris & kolom mulai dari 0.",
        items: {
          type: "object",
          properties: {
            r: { type: "integer", description: "Indeks baris (mulai 0)." },
            c: { type: "integer", description: "Indeks kolom (mulai 0)." },
            newText: { type: "string", description: "Isi sel setelah dirapikan." },
          },
          required: ["r", "c", "newText"],
        },
      },
    },
    required: ["summary"],
  },
};

const SYSTEM_PROMPT = [
  "Nama Anda FRIDA, asisten penyunting cerdas yang tertanam langsung di dalam Microsoft Word.",
  "Anda menerima: (a) seluruh dokumen sebagai array paragraf {\"i\": indeks, \"text\": isi};",
  "(b) opsional, teks yang sedang diseleksi pengguna; (c) opsional, isi tabel yang diseleksi sebagai grid baris x kolom.",
  "Tugas Anda MENERAPKAN instruksi pengguna langsung ke dokumen, bukan sekadar menjawab. Aturan memilih aksi:",
  "- Jika diminta memperbaiki/mengubah/menerjemahkan teks yang sudah ada -> paragraphOps action 'replace' pada paragraf terkait.",
  "- Jika diminta MENAMBAHKAN/MENYISIPKAN kalimat atau paragraf baru di lokasi tertentu (mis. 'tambahkan di paragraf ini', 'lanjutkan paragraf ini') -> 'insertAfter' pada indeks paragraf itu, ATAU 'replace' bila teks baru menyatu dengan paragraf yang sama.",
  "- Jika diminta menambah paragraf di AKHIR dokumen (mis. 'tambahkan 2 paragraf lagi di akhir') -> 'append', satu op per paragraf baru, isi yang relevan dan nyambung dengan konteks dokumen.",
  "- Jika ada tabel diseleksi dan diminta merapikan/membetulkan tabel -> gunakan tableOps untuk sel yang perlu diperbaiki (rapikan ejaan, kapitalisasi, spasi, konsistensi; jangan mengubah makna data).",
  "Jika pengguna menyebut 'paragraf ini' atau 'di sini' dan ada teks terseleksi, anggap itu paragraf yang diseleksi.",
  "Pertahankan bahasa dokumen. Jangan mengubah bagian yang tidak diminta. Laporkan hasil lewat tool apply_changes.",
].join("\n");

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// satu kali panggilan ke provider
async function callOnce(userContent) {
  const url = cfg.baseUrl.replace(/\/?$/, "/") + "v1/messages";
  const resp = await fetch(url, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": cfg.apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: cfg.model,
      max_tokens: cfg.maxTokens || 8000,
      system: SYSTEM_PROMPT,
      tools: [EDIT_TOOL],
      tool_choice: { type: "tool", name: "apply_changes" },
      messages: [{ role: "user", content: userContent }],
    }),
  });
  const raw = await resp.text();
  if (!resp.ok) throw new Error("Provider " + resp.status + ": " + raw.slice(0, 500));
  const data = JSON.parse(raw);
  // Catatan: beberapa provider/proxy mengganti nama tool di respons,
  // jadi jangan cocokkan nama persis — ambil blok tool_use pertama saja.
  const toolUse = (data.content || []).find(b => b.type === "tool_use" && b.input);
  if (!toolUse) throw new Error("Model tidak mengembalikan hasil terstruktur.");
  return toolUse.input;
}

// ---- panggil model dengan retry (provider kadang tersendat sesaat) ----
async function callClaude(userContent) {
  const maxTries = 3;
  let lastErr;
  for (let attempt = 1; attempt <= maxTries; attempt++) {
    try {
      return await callOnce(userContent);
    } catch (err) {
      lastErr = err;
      console.warn("Percobaan " + attempt + "/" + maxTries + " gagal: " + (err.message || err));
      if (attempt < maxTries) await sleep(800 * attempt);
    }
  }
  throw lastErr;
}

async function handleEdit(req, res) {
  let body = "";
  req.on("data", c => (body += c));
  req.on("end", async () => {
    try {
      const { paragraphs, instruction, selection, table } = JSON.parse(body || "{}");
      if (!Array.isArray(paragraphs) || !instruction) {
        res.writeHead(400, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: "paragraphs[] dan instruction wajib diisi" }));
        return;
      }
      let userContent =
        "INSTRUKSI PENGGUNA:\n" + instruction +
        "\n\nDOKUMEN (array paragraf):\n" + JSON.stringify(paragraphs);
      if (selection && selection.text) {
        userContent += "\n\nTEKS YANG SEDANG DISELEKSI PENGGUNA:\n" + JSON.stringify(selection);
      }
      if (table && Array.isArray(table.rows) && table.rows.length) {
        userContent += "\n\nTABEL YANG SEDANG DISELEKSI (grid baris x kolom, indeks mulai 0):\n" + JSON.stringify(table.rows);
      }
      const result = await callClaude(userContent);
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify(result));
    } catch (err) {
      res.writeHead(500, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: String(err.message || err) }));
    }
  });
}

(async () => {
  let httpsOptions;
  try {
    httpsOptions = await devCerts.getHttpsServerOptions();
  } catch (e) {
    console.error("Gagal memuat sertifikat HTTPS. Jalankan dulu:  npm run cert");
    console.error(String(e.message || e));
    process.exit(1);
  }

  https
    .createServer(httpsOptions, (req, res) => {
      if (req.method === "POST" && req.url.startsWith("/api/edit")) return handleEdit(req, res);
      if (req.method === "GET") return serveStatic(req, res);
      res.writeHead(405); res.end("Method not allowed");
    })
    .listen(PORT, () => {
      console.log("FRIDA berjalan di  https://localhost:" + PORT + "/taskpane.html");
      console.log("Provider :", cfg.baseUrl);
      console.log("Model    :", cfg.model);
      console.log("Tools    :", TOOL_SCHEMAS.length, "terdaftar (" + TOOL_SCHEMAS.map(t => t.name).join(", ") + ")");
      console.log("Biarkan jendela ini terbuka selama memakai add-in di Word.");
    });
})();
