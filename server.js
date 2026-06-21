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
const { SCHEMAS: TOOL_SCHEMAS, resolveName: resolveToolName, runtimeOf } = require("./tools/schemas");
const ragAgentTools = require("./rag/agent_tools");

// Tools yang DIKIRIM ke provider hanya boleh punya {name, description, input_schema}.
// Field internal (mis. `runtime`) ditolak Anthropic ("Extra inputs are not permitted").
const API_TOOLS = TOOL_SCHEMAS.map((t) => ({
  name: t.name, description: t.description, input_schema: t.input_schema,
}));

// Research Copilot / RAG (R0) — ingestion sumber + status embeddings provider.
const ingest = require("./rag/ingest");
const sourceStore = require("./rag/store");
const embeddings = require("./rag/embeddings");
const vectors = require("./rag/vectors");
const cite = require("./rag/cite");

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

// ===================== FASE 2: endpoint agentic =====================
// Berbeda dari /api/edit (one-shot, single tool), /api/agent adalah RELAY tipis:
// klien mengirim SELURUH riwayat `messages` (termasuk tool_result dari eksekusi
// sebelumnya), server memanggil provider SEKALI dengan daftar tools dari registry
// (tool_choice: auto), lalu mengembalikan blok `content` apa adanya. LOOP ada di
// KLIEN karena eksekusi tool wajib di dalam Word.run (Office.js). Server tetap
// stateless & tidak pernah menyentuh dokumen.

const AGENT_SYSTEM_PROMPT = [
  "Nama Anda FRIDA, agen penyunting yang MENGENDALIKAN Microsoft Word lewat tool.",
  "Anda tidak menyunting teks secara langsung; Anda memanggil tool yang disediakan.",
  "ALUR WAJIB:",
  "1) Panggil get_document_outline DULU untuk memahami struktur & indeks paragraf (kecuali instruksi jelas hanya soal seleksi aktif).",
  "2) Susun rencana seminimal mungkin, lalu panggil tool write satu per satu.",
  "3) Pakai selektor 'target' yang tepat: mode 'heading' untuk semua judul, 'whole_document' untuk seluruh dokumen, 'selection' untuk blok aktif, 'paragraph_index' untuk paragraf tertentu, 'search' untuk kemunculan teks.",
  "PEMILIHAN TOOL (penting, jangan keliru):",
  "- 'ubah/ganti POSISI / ORIENTASI / TATA LETAK halaman', 'jadikan landscape/portrait', 'ganti ukuran kertas/A4', 'atur margin' -> set_page_layout. JANGAN pakai set_page_numbers untuk ini.",
  "- 'beri/tambahkan NOMOR halaman', 'page number' -> set_page_numbers. Tool ini BISA menaruh nomor di ATAS (position=top) atau BAWAH (position=bottom), dan rata kiri/tengah/kanan. 'nomor di tengah atas' -> position=top, alignment=Centered. Ini nomor berjalan otomatis (1,2,3), jadi JANGAN menolak permintaan nomor di atas.",
  "- 'tulis teks di header/footer' -> manage_header_footer.",
  "- 'buat tabel bergaris penuh/grid', 'beri garis di semua sel', 'ubah border tabel' (tabel SUDAH ADA) -> format_table dengan borders='all'. JANGAN pakai create_table untuk mengubah tabel yang sudah ada.",
  "- 'buat tabel baru' atau 'ubah teks jadi tabel' -> create_table.",
  "- MENAMBAH/MENYISIPKAN paragraf atau teks baru ('tambahkan paragraf', 'tulis di halaman/paragraf X', 'isi halaman kosong') -> insert_paragraph (location end/after_index/before_index/after_selection). DILARANG memakai replace_text dengan find kosong untuk menambah teks — itu error.",
  "- 'buatkan cover/halaman judul/halaman sampul' -> insert_cover_page (1 panggilan, isi judul/penulis/tanggal dari konteks).",
  "- 'format jadi proposal bisnis', 'rapikan jadi dokumen profesional' -> format_business_proposal (1 panggilan, jangan urai jadi banyak tool kecil).",
  "- Pertanyaan/RINGKASAN yang merujuk dokumen/jurnal yang DIUNGGAH ('cari di sumber', 'ringkas jurnal ini', 'menurut paper terunggah') -> panggil search_uploaded_sources DULU, lalu jawab HANYA berdasarkan kutipan (sertakan source_id). JANGAN mengarang.",
  "- MENULIS/MENAMBAH PARAGRAF berbasis sumber ('tambahkan paragraf tentang X berdasarkan jurnal', 'tulis paragraf dari sumber') -> WAJIB pakai generate_paragraph_from_source (jangan menulis paragraf sendiri). Bila hasilnya needsMoreEvidence=true, sampaikan ke pengguna bahwa bukti tak cukup dan JANGAN menyisipkan apa pun. Bila ada paragraf, sisipkan dengan insert_paragraph memakai field 'paragraph' apa adanya, lalu beri tahu pengguna sumber/sitasi (verifiedCitations) dan peringatan bila ada flaggedCitations.",
  "Jika search_uploaded_sources tak mengembalikan kutipan relevan, katakan terus terang bahwa bukti di sumber tak cukup — jangan mengarang.",
  "- RESOLVE SUMBER DARI NAMA ALAMI ('jurnal Hijra', 'paper Nair 2012', 'sumber tentang agroforestri') → pakai resolve_source DULU untuk mendapat source_id (best_id), BARU panggil summarize_source / compare_sources / insert_citation.",
  "- RINGKAS SATU SUMBER ('ringkas jurnal ini', 'jelaskan paper X', 'apa isi sumber Y') → pastikan punya source_id (resolve_source jika pengguna hanya menyebut nama) lalu panggil summarize_source. Sampaikan hasilnya ke pengguna.",
  "- BANDINGKAN SUMBER ('bandingkan jurnal A dan B', 'perbedaan ketiga paper', 'compare sources') → pastikan punya ≥2 source_id (resolve_source bila perlu) lalu panggil compare_sources. Sampaikan comparison + similarities + differences ke pengguna.",
  "- SITASI ('sisipkan sitasi APA7', 'kasih sitasi') → insert_citation dengan source_id (dapatkan source_id dari hasil search_uploaded_sources/generate_paragraph_from_source/resolve_source — JANGAN menulis nama/tahun sendiri). 'buat daftar pustaka'/'bibliography' → insert_bibliography. Jika tool sitasi mengembalikan error metadata kosong, beri tahu pengguna untuk melengkapi/mengoreksi metadata sumber di panel Sumber.",
  "Jika sebuah tool mengembalikan error, JANGAN mengulang tool yang sama berkali-kali; baca pesan error, perbaiki argumen, atau laporkan ke pengguna dengan teks.",
  "4) Setelah semua tool selesai dan tujuan tercapai, jawab dengan teks ringkas (tanpa memanggil tool lagi) yang merangkum apa yang dilakukan, dalam Bahasa Indonesia.",
  "Jangan mengubah bagian yang tidak diminta. Pertahankan bahasa dokumen.",
  "Jika instruksi ambigu atau berisiko (mis. mengganti di seluruh dokumen), tetap usulkan tool call yang paling masuk akal; konfirmasi keamanan ditangani oleh aplikasi klien.",
].join("\n");

async function callAgentOnce(messages) {
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
      system: AGENT_SYSTEM_PROMPT,
      tools: API_TOOLS,
      tool_choice: { type: "auto" },
      messages,
    }),
  });
  const raw = await resp.text();
  if (!resp.ok) throw new Error("Provider " + resp.status + ": " + raw.slice(0, 500));
  const data = JSON.parse(raw);
  // Kembalikan apa adanya yang dibutuhkan klien untuk melanjutkan loop.
  return {
    stop_reason: data.stop_reason,
    content: data.content || [],
  };
}

async function callAgent(messages) {
  const maxTries = 3;
  let lastErr;
  for (let attempt = 1; attempt <= maxTries; attempt++) {
    try {
      return await callAgentOnce(messages);
    } catch (err) {
      lastErr = err;
      console.warn("Agent percobaan " + attempt + "/" + maxTries + " gagal: " + (err.message || err));
      if (attempt < maxTries) await sleep(800 * attempt);
    }
  }
  throw lastErr;
}

// Jalankan satu tool server (RAG) -> blok tool_result.
async function runServerTool(tu) {
  const real = resolveToolName(tu.name) || tu.name;
  const out = await ragAgentTools.executeServerTool(real, tu.input || {});
  return {
    type: "tool_result", tool_use_id: tu.id,
    is_error: !!(out && out.error),
    content: JSON.stringify(out),
  };
}

// Loop agentic di SERVER: tool server (RAG) dieksekusi di sini; saat model
// memanggil tool client (Word), kembalikan ke task pane untuk Word.run.
const AGENT_MAX_STEPS = 12;
async function runAgentServerLoop(messages) {
  for (let step = 0; step < AGENT_MAX_STEPS; step++) {
    const data = await callAgent(messages);
    messages.push({ role: "assistant", content: data.content });

    const toolUses = (data.content || []).filter((b) => b.type === "tool_use");
    if (!toolUses.length) return { done: true, content: data.content, messages };

    const serverTU = [], clientTU = [];
    toolUses.forEach((tu) => (runtimeOf(tu.name) === "server" ? serverTU : clientTU).push(tu));

    if (clientTU.length === 0) {
      // semua server-tool -> eksekusi & lanjut loop tanpa ke klien
      const results = [];
      for (const tu of serverTU) results.push(await runServerTool(tu));
      messages.push({ role: "user", content: results });
      continue;
    }
    // ada client-tool -> eksekusi server-tool yang menyertai, lalu kembali ke klien
    const serverResults = [];
    for (const tu of serverTU) serverResults.push(await runServerTool(tu));
    return { done: false, content: data.content, messages, serverResults };
  }
  return { done: false, content: [{ type: "text", text: "Batas langkah server tercapai." }],
           messages, serverResults: [] };
}

function handleAgent(req, res) {
  let body = "";
  req.on("data", (c) => (body += c));
  req.on("end", async () => {
    try {
      const { messages } = JSON.parse(body || "{}");
      if (!Array.isArray(messages) || messages.length === 0) {
        res.writeHead(400, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: "messages[] wajib diisi" }));
        return;
      }
      const result = await runAgentServerLoop(messages);
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify(result));
    } catch (err) {
      res.writeHead(500, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: String(err.message || err) }));
    }
  });
}
// =====================================================================

// ===================== Research Copilot R0: /api/sources =====================
function readBody(req) {
  return new Promise((resolve) => {
    let b = ""; req.on("data", (c) => (b += c)); req.on("end", () => resolve(b));
  });
}
function sendJson(res, code, obj) {
  res.writeHead(code, { "Content-Type": "application/json" });
  res.end(JSON.stringify(obj));
}

async function handleSources(req, res) {
  const url = req.url.split("?")[0];
  try {
    // POST /api/sources/upload  { filename, mime, dataBase64, workspace }
    if (req.method === "POST" && url === "/api/sources/upload") {
      const body = JSON.parse((await readBody(req)) || "{}");
      const result = await ingest.ingestUpload(body);
      return sendJson(res, 200, result);
    }
    // GET /api/sources  -> daftar KB
    if (req.method === "GET" && url === "/api/sources") {
      const ws = (req.url.split("?")[1] || "").match(/workspace=([^&]+)/);
      return sendJson(res, 200, { sources: sourceStore.list(ws ? decodeURIComponent(ws[1]) : null) });
    }
    // GET /api/sources/embed-status -> status provider embeddings (tanpa key)
    if (req.method === "GET" && url === "/api/sources/embed-status") {
      return sendJson(res, 200, embeddings.status());
    }
    // POST /api/sources/reindex -> embed dokumen yg belum ber-vektor
    if (req.method === "POST" && url === "/api/sources/reindex") {
      const result = await ingest.reindexAll();
      return sendJson(res, 200, { result });
    }
    // POST /api/sources/search  { query, k, document_ids, workspace }
    if (req.method === "POST" && url === "/api/sources/search") {
      const body = JSON.parse((await readBody(req)) || "{}");
      if (!body.query) return sendJson(res, 400, { error: "query wajib diisi" });
      const docs = (body.document_ids && body.document_ids.length)
        ? body.document_ids
        : sourceStore.list(body.workspace).map((d) => d.id);
      const [qvec] = await embeddings.embed([body.query]);
      const hits = vectors.search(qvec, docs, { k: body.k || 8, minScore: body.minScore });
      // sertakan judul sumber utk konteks
      const titles = {};
      sourceStore.list().forEach((d) => (titles[d.id] = d.title));
      return sendJson(res, 200, {
        hits: hits.map((h) => ({ ...h, title: titles[h.document_id] || null })),
      });
    }
    // POST /api/sources/cite  { source_id, style, page, narrative }
    if (req.method === "POST" && url === "/api/sources/cite") {
      const b = JSON.parse((await readBody(req)) || "{}");
      const text = cite.inTextFor(b.source_id, b.style || "APA7",
        { page: b.page, narrative: b.narrative });
      if (text == null) return sendJson(res, 404, { error: "metadata sumber kosong; lengkapi dulu metadata." });
      return sendJson(res, 200, { text });
    }
    // POST /api/sources/bibliography  { source_ids, style }
    if (req.method === "POST" && url === "/api/sources/bibliography") {
      const b = JSON.parse((await readBody(req)) || "{}");
      const entries = cite.bibliography(b.source_ids, b.style || "APA7");
      return sendJson(res, 200, { entries });
    }
    // PATCH /api/sources/:id/metadata  { csl }
    if (req.method === "PATCH" && /^\/api\/sources\/[\w-]+\/metadata$/.test(url)) {
      const id = url.split("/")[3];
      const b = JSON.parse((await readBody(req)) || "{}");
      const csl2 = sourceStore.updateMetadata(id, b.csl || {});
      if (!csl2) return sendJson(res, 404, { error: "sumber tak ditemukan" });
      return sendJson(res, 200, { csl: csl2 });
    }
    // DELETE /api/sources/:id
    if (req.method === "DELETE" && /^\/api\/sources\/[\w-]+$/.test(url)) {
      const id = url.split("/").pop();
      vectors.removeChunks(id);
      return sendJson(res, 200, { removed: sourceStore.remove(id) });
    }
    return sendJson(res, 404, { error: "rute sumber tidak dikenal" });
  } catch (err) {
    return sendJson(res, 500, { error: String(err.message || err) });
  }
}
// =============================================================================

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
      if (req.url.startsWith("/api/sources")) return handleSources(req, res);
      if (req.method === "POST" && req.url.startsWith("/api/agent")) return handleAgent(req, res);
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
