/* FRIDA Agent (Fase 2) — loop agentic di klien.
   Alur:
     1) Pengguna kirim instruksi -> ditambah ke `messages`.
     2) POST /api/agent (server relay tipis ke LLM dgn daftar tools dari registry).
     3) LLM membalas: bisa teks, bisa satu/lebih blok tool_use.
     4) Klien MENJALANKAN tool_use lewat HANDLERS di dalam Word.run (Office.js),
        membungkus hasilnya jadi tool_result, lalu KIRIM LAGI -> ulangi.
     5) Berhenti ketika LLM membalas tanpa tool_use (jawaban akhir).

   Jembatan keamanan sebelum Fase 3 (TransactionManager/rollback):
     - Tool READ (mis. get_document_outline) jalan OTOMATIS.
     - Batch yang memuat tool WRITE DITAHAN: ditampilkan dulu, butuh klik "Jalankan".
   Registry datang dari window.FRIDA_SCHEMAS & window.FRIDA_HANDLERS (file tools/). */

const READ_TOOLS = new Set(["get_document_outline", "search_text"]);
const MAX_STEPS = 40; // batas langkah loop (safety) — dinaikkan dari 12 agar tugas
                      // multi-format (heading + font + spasi + tabel) tidak mentok
                      // sebelum selesai. AI tetap diminta membatch aksi (lihat prompt).

// Nama tool dari provider kadang di-rename (lihat resolveHandler di handlers.js).
// Resolusikan ke nama registry kanonik sebelum klasifikasi read/write & pelabelan.
function canonName(name) {
  const r = window.FRIDA_HANDLERS && window.FRIDA_HANDLERS.resolveHandler(name);
  return r ? r.name : name;
}
function isReadTool(name) { return READ_TOOLS.has(canonName(name)); }

let messages = []; // riwayat percakapan untuk LLM
let running = false;
let audit = null;          // FRIDA_SAFETY.makeAuditLog()
let lastSnapshot = null;   // OOXML sebelum perintah terakhir (untuk Undo FRIDA)

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("FRIDA hanya untuk Microsoft Word.", "err");
    return;
  }
  if (!window.FRIDA_HANDLERS || !window.FRIDA_SCHEMAS || !window.FRIDA_SAFETY) {
    setStatus("Registry tool gagal dimuat (tools/schemas.js, handlers.js, safety.js).", "err");
    return;
  }
  audit = window.FRIDA_SAFETY.makeAuditLog();
  document.getElementById("send").onclick = onSend;
  const undoBtn = document.getElementById("undo");
  if (undoBtn) undoBtn.onclick = onUndo;
  document.getElementById("instruction").addEventListener("keydown", (e) => {
    // Enter = kirim; Shift+Enter = baris baru. (Ctrl/Cmd+Enter juga kirim.)
    if (e.key === "Enter" && !e.shiftKey && !e.isComposing) { e.preventDefault(); onSend(); }
  });

  // Tab navigation
  document.querySelectorAll(".nav-tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".nav-tab").forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".tab-content").forEach((c) => c.classList.remove("active"));
      tab.classList.add("active");
      const target = tab.getAttribute("data-tab");
      document.getElementById("tab-" + target).classList.add("active");
    });
  });

  // [FITUR 1] Restore riwayat chat dari localStorage
  restoreChat();

  // [FITUR 2] Dark mode: load preferensi
  initDarkMode();

  // [FITUR 3] Template perintah cepat
  initQuickPrompts();

  // [FITUR 4] Statistik dokumen — load awal
  updateDocStats();

  // [FITUR 1] Tombol bersihkan chat
  var clearBtn = document.getElementById("clearChat");
  if (clearBtn) clearBtn.onclick = clearChat;

  // [FITUR 4] Refresh stats
  var statsRefBtn = document.getElementById("statsRefresh");
  if (statsRefBtn) statsRefBtn.onclick = updateDocStats;

  // [FITUR 5] Context memory — load dari localStorage
  initContextMemory();
});

// ---------- UI helpers ----------
function setStatus(msg, cls) {
  const el = document.getElementById("status");
  el.textContent = msg || "";
  el.className = "status" + (cls ? " " + cls : "");
}
function busy(on) {
  running = on;
  document.getElementById("send").disabled = on;
  document.getElementById("instruction").disabled = on;
}
function logEl() { return document.getElementById("log"); }
function scrollDown() { const l = logEl(); l.scrollTop = l.scrollHeight; }

function escapeHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function renderMarkdown(text) {
  let out = escapeHtml(String(text || ""));
  out = out.replace(/```([\s\S]*?)```/g, '<pre class="md-codeblock"><code>$1</code></pre>');
  out = out.replace(/`([^`]+)`/g, '<code>$1</code>');
  out = out.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  out = out.replace(/\*([^*]+)\*/g, '<em>$1</em>');
  out = out.replace(/\n/g, '<br />');
  return out;
}

let typingBubble = null;
function showTypingIndicator() {
  if (typingBubble) return typingBubble;
  typingBubble = document.createElement("div");
  typingBubble.className = "msg assistant typing";
  typingBubble.innerHTML = '<div class="typing-dots"><span></span><span></span><span></span></div><div class="typing-text">FRIDA sedang berpikir...</div>';
  logEl().appendChild(typingBubble);
  scrollDown();
  return typingBubble;
}
function hideTypingIndicator() {
  if (typingBubble) {
    typingBubble.remove();
    typingBubble = null;
  }
}

function addBubble(role, text) {
  const div = document.createElement("div");
  div.className = "msg " + role;
  const raw = String(text || "");
  const content = role === "assistant" ? renderMarkdown(raw) : escapeHtml(raw).replace(/\n/g, "<br />");
  div.innerHTML = '<div class="msg-body">' + content + '</div>';
  logEl().appendChild(div);
  scrollDown();
  return div;
}

// ---------- alur utama ----------
async function onSend() {
  if (running) return;
  const input = document.getElementById("instruction");
  const text = input.value.trim();
  if (!text) { setStatus("Tulis dulu perintahnya untuk FRIDA.", "err"); return; }

  // Snapshot riwayat SEBELUM giliran ini. Bila request gagal (mis. gateway memblokir
  // konten / error jaringan), kita rollback ke sini supaya pesan yang gagal TIDAK ikut
  // terkirim lagi di request berikutnya dan memblokir seluruh percakapan.
  const historyBefore = messages.slice();

  addBubble("user", text);
  messages.push({ role: "user", content: text });
  input.value = "";

  busy(true);
  setStatus("FRIDA berpikir…");
  showTypingIndicator();
  try {
    await runAgentLoop();
    setStatus("Selesai.", "ok");
  } catch (err) {
    messages = historyBefore; // buang giliran yang gagal dari riwayat
    setStatus("Gagal: " + (err.message || err), "err");
    addBubble("error", "⚠️ " + (err.message || String(err)));
  } finally {
    hideTypingIndicator();
    busy(false);
    // Bersihkan riwayat aksi setelah selesai agar chat tetap rapi
    var auditBox = document.getElementById("audit");
    if (auditBox) auditBox.innerHTML = "";
    if (audit && audit.entries) audit.entries = [];
    // [FITUR 1] Simpan riwayat chat
    saveChat();
    // [FITUR 4] Refresh statistik dokumen setelah edit
    updateDocStats();
  }
}

// Tool dengan runtime "server" (mis. RAG) dieksekusi di server; klien hanya
// menjalankan tool Word (runtime client). resolusi via registry (tahan rename).
function isClientTool(name) {
  const rt = window.FRIDA_SCHEMAS && window.FRIDA_SCHEMAS.runtimeOf
    ? window.FRIDA_SCHEMAS.runtimeOf(name) : "client";
  return rt !== "server";
}

async function runAgentLoop() {
  for (let step = 0; step < MAX_STEPS; step++) {
    setStatus("FRIDA berpikir… (langkah " + (step + 1) + ")");
    // [FITUR 5] Sertakan context memory jika ada
    var cm = getContextMemory();
    var agentBody = { messages: messages };
    if (cm) agentBody.contextMemory = cm;
    const resp = await fetch("/api/agent", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(agentBody),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || ("HTTP " + resp.status));

    // Server adalah sumber kebenaran riwayat (sudah menjalankan tool server/RAG).
    if (Array.isArray(data.messages)) messages = data.messages;
    const content = Array.isArray(data.content) ? data.content : [];

    // tampilkan teks asisten
    content.filter((b) => b.type === "text" && b.text && b.text.trim())
           .forEach((b) => addBubble("assistant", b.text));

    // tampilkan aktivitas tool server (mis. pencarian sumber) sbg info
    const serverResults = Array.isArray(data.serverResults) ? data.serverResults : [];
    renderServerActivity(content, serverResults);

    if (data.done) return; // jawaban akhir

    // hanya tool client (Word) yang dieksekusi di sini
    const clientToolUses = content.filter((b) => b.type === "tool_use" && isClientTool(b.name));
    if (clientToolUses.length === 0) {
      // tak ada yang bisa dikerjakan klien; lanjut bawa hasil server bila ada
      if (serverResults.length) { messages.push({ role: "user", content: serverResults }); continue; }
      return;
    }

    const { needsConfirm, Permissions } = window.FRIDA_SAFETY;

    const blocked = clientToolUses.filter((t) => Permissions.isBlocked(canonName(t.name)));
    if (blocked.length) {
      messages.push({ role: "user", content: serverResults.concat(clientToolUses.map((t) => ({
        type: "tool_result", tool_use_id: t.id, is_error: true,
        content: "Ditolak kebijakan keamanan: " + canonName(t.name),
      }))) });
      addBubble("error", "Aksi diblokir kebijakan: " + blocked.map((t) => canonName(t.name)).join(", "));
      return;
    }

    const risky = clientToolUses.filter((t) => needsConfirm({ name: canonName(t.name), input: t.input }));
    const hasWrite = clientToolUses.some((t) => !isReadTool(t.name));
    const mustConfirm = risky.length > 0 || (previewMode() && hasWrite);
    if (mustConfirm) {
      setStatus("FRIDA menyiapkan pratinjau…");
      const previews = await computePreviews(clientToolUses);
      const okGo = await confirmBatch(clientToolUses, risky, previews);
      if (!okGo) {
        messages.push({ role: "user", content: serverResults.concat(clientToolUses.map((t) => ({
          type: "tool_result", tool_use_id: t.id, is_error: true,
          content: "Dibatalkan oleh pengguna.",
        }))) });
        addBubble("error", "Dibatalkan. FRIDA tidak menerapkan perubahan.");
        return;
      }
    } else {
      renderToolCalls(clientToolUses, hasWrite ? "menerapkan…" : "membaca dokumen…");
    }

    setStatus("FRIDA menjalankan " + clientToolUses.length + " aksi…");
    const results = await executeToolUses(clientToolUses, hasWrite);
    // gabung hasil server-tool (jika ada di turn yg sama) + hasil client-tool
    messages.push({ role: "user", content: serverResults.concat(results) });
    if (hasWrite) renderAudit();
  }
  addBubble("error", "Batas langkah tercapai (" + MAX_STEPS + ").");
}

// Tampilkan ringkas hasil tool server (mis. pencarian sumber) di chat.
function renderServerActivity(content, serverResults) {
  if (!serverResults || !serverResults.length) return;
  const byId = {};
  serverResults.forEach((r) => (byId[r.tool_use_id] = r));
  content.filter((b) => b.type === "tool_use" && byId[b.id]).forEach((b) => {
    let info = "";
    try {
      const out = JSON.parse(byId[b.id].content);
      if (out.hits) info = out.hits.length + " kutipan ditemukan" +
        (out.hits[0] ? " (skor " + out.hits[0].score + ")" : "");
      else if (out.note) info = out.note;
      else if (out.error) info = "error: " + out.error;
    } catch (_) {}
    const card = document.createElement("div");
    card.className = "tool";
    card.innerHTML = '<div class="tag">🔎 ' + escapeHtml(canonName(b.name)) + "</div>" +
      '<div class="tnote ok">' + escapeHtml(info) + "</div>";
    logEl().appendChild(card);
  });
  scrollDown();
}

// Eksekusi satu batch tool_use dalam SATU Word.run -> array tool_result.
// isWrite=true: ambil snapshot OOXML dulu (TransactionManager); bila ada kegagalan
// di tengah batch -> rollback seluruh batch agar dokumen tidak setengah jadi.
async function executeToolUses(toolUses, isWrite) {
  const { resolveHandler } = window.FRIDA_HANDLERS;
  const { TransactionManager } = window.FRIDA_SAFETY;
  const results = [];
  let anyError = false;

  await Word.run(async (context) => {
    const tx = isWrite ? await TransactionManager.begin(context) : null;

    for (const tu of toolUses) {
      const resolved = resolveHandler(tu.name);
      if (!resolved) {
        results.push(toolResult(tu.id, { error: "tool tidak dikenal: " + tu.name }, true));
        markToolDone(tu.id, { error: "tidak dikenal" }, true);
        audit && audit.record({ name: tu.name, input: tu.input }, { error: "tidak dikenal" }, true);
        anyError = true;
        continue;
      }
      try {
        const out = await resolved.fn(context, tu.input || {});
        results.push(toolResult(tu.id, out, false));
        markToolDone(tu.id, out);
        audit && audit.record({ name: resolved.name, input: tu.input }, out, false);
      } catch (e) {
        const msg = officeErr(e);
        results.push(toolResult(tu.id, { error: msg }, true));
        markToolDone(tu.id, { error: msg }, true);
        audit && audit.record({ name: resolved.name, input: tu.input }, { error: msg }, true);
        anyError = true;
      }
    }

    if (tx) {
      if (anyError) {
        // batch write gagal sebagian -> kembalikan ke kondisi sebelum batch
        await tx.rollback();
        addBubble("error", "Sebagian aksi gagal — perubahan batch ini dibatalkan (rollback).");
      } else {
        // batch sukses: simpan snapshot SEBELUM batch sbg titik Undo FRIDA
        lastSnapshot = tx.snapshot;
        showUndo(true);
      }
    }
  });
  return results;
}

function toolResult(id, obj, isError) {
  return { type: "tool_result", tool_use_id: id, is_error: !!isError,
           content: JSON.stringify(obj) };
}

// Ekstrak detail dari error Office.js (RichApi.Error punya .code & .debugInfo)
// sehingga model & pengguna melihat lebih dari sekadar "GeneralException".
function officeErr(e) {
  if (!e) return "error tak diketahui";
  let s = (e.code ? e.code + ": " : "") + (e.message || String(e));
  if (e.debugInfo) {
    const d = e.debugInfo;
    const extra = [d.errorLocation, d.message].filter(Boolean).join(" @ ");
    if (extra && extra !== e.message) s += " (" + extra + ")";
  }
  return s;
}

// ---------- pratinjau (Fase 5) ----------
function previewMode() {
  const c = document.getElementById("previewToggle");
  return c ? c.checked : true; // default: tinjau dulu
}

// Hitung dampak READ-ONLY tiap tool dalam SATU Word.run (tanpa mengubah dokumen).
async function computePreviews(toolUses) {
  const { previewTool } = window.FRIDA_HANDLERS;
  const out = {};
  try {
    await Word.run(async (context) => {
      for (const tu of toolUses) {
        out[tu.id] = await previewTool(context, tu.name, tu.input);
      }
    });
  } catch (e) { /* pratinjau gagal: biarkan kosong, tetap bisa konfirmasi */ }
  return out;
}

// ---------- konfirmasi & undo & audit ----------
function confirmBatch(toolUses, risky, previews) {
  renderToolCalls(toolUses, "menunggu konfirmasi…", previews);
  const riskNames = (risky || []).map((t) => canonName(t.name));
  return new Promise((resolve) => {
    const wrap = document.createElement("div");
    wrap.className = "confirm";
    if (riskNames.length) {
      const warn = document.createElement("div");
      warn.className = "risk";
      warn.textContent = "⚠️ Aksi berisiko: " + riskNames.join(", ") +
        ". Periksa dulu sebelum menjalankan.";
      wrap.appendChild(warn);
    }
    const btns = document.createElement("div");
    btns.className = "confirm-btns";
    const yes = document.createElement("button");
    yes.className = "primary"; yes.textContent = "Jalankan " + toolUses.length + " aksi";
    const no = document.createElement("button");
    no.textContent = "Batal";
    yes.onclick = () => { wrap.remove(); resolve(true); };
    no.onclick = () => { wrap.remove(); resolve(false); };
    btns.appendChild(yes); btns.appendChild(no);
    wrap.appendChild(btns);
    logEl().appendChild(wrap);
    scrollDown();
  });
}

function showUndo(on) {
  const btn = document.getElementById("undo");
  if (btn) btn.style.display = on ? "block" : "none";
}

async function onUndo() {
  if (running || lastSnapshot == null) return;
  busy(true);
  setStatus("Mengembalikan dokumen…");
  try {
    await window.FRIDA_SAFETY.TransactionManager.restore(lastSnapshot);
    addBubble("assistant", "↩️ Dokumen dikembalikan ke kondisi sebelum perubahan terakhir.");
    lastSnapshot = null;
    showUndo(false);
    setStatus("Dikembalikan.", "ok");
  } catch (err) {
    setStatus("Gagal undo: " + (err.message || err), "err");
  } finally {
    busy(false);
  }
}

function renderAudit() {
  const box = document.getElementById("audit");
  if (!box || !audit) return;
  const rows = audit.entries.slice(-8).reverse();
  if (!rows.length) { box.innerHTML = ""; return; }
  box.innerHTML = "<div class='audit-h'>Riwayat aksi</div>" +
    rows.map((e) =>
      "<div class='audit-row " + (e.ok ? "ok" : "err") + "'>" +
      escapeHtml(e.tool) + " — " + escapeHtml(e.result || (e.ok ? "ok" : "gagal")) +
      "</div>").join("");
}

function renderToolCalls(toolUses, note, previews) {
  toolUses.forEach((tu) => {
    const card = document.createElement("div");
    card.className = "tool";
    card.id = "tool-" + tu.id;
    const isRead = isReadTool(tu.name);
    const pv = previews && previews[tu.id];
    card.innerHTML =
      '<div class="tag">' + (isRead ? "🔍 " : "✏️ ") + escapeHtml(canonName(tu.name)) + "</div>" +
      (pv ? '<div class="preview">↳ ' + escapeHtml(pv) + "</div>" : "") +
      '<div class="args">' + escapeHtml(summarizeArgs(tu.input)) + "</div>" +
      '<div class="tnote">' + escapeHtml(note || "") + "</div>";
    logEl().appendChild(card);
  });
  scrollDown();
}

function markToolDone(id, out, isError) {
  const card = document.getElementById("tool-" + id);
  if (!card) return;
  const note = card.querySelector(".tnote");
  if (!note) return;
  if (isError) { note.textContent = "gagal: " + (out.error || ""); note.className = "tnote err"; }
  else { note.textContent = "selesai " + summarizeOut(out); note.className = "tnote ok"; }
}

function summarizeArgs(input) {
  if (!input) return "";
  const t = input.target;
  const scope = t ? (t.mode + (t.value ? " '" + t.value + "'" : "") +
                     (t.index != null ? " #" + t.index : "")) : "";
  const rest = Object.keys(input)
    .filter((k) => k !== "target")
    .map((k) => k + "=" + JSON.stringify(input[k]))
    .join(", ");
  return [scope, rest].filter(Boolean).join(" | ");
}
function summarizeOut(out) {
  if (out == null) return "";
  if (out.replaced != null) return "(" + out.replaced + " diganti)";
  if (out.applied != null) return "(" + out.applied + " range)";
  if (out.paragraphs != null) return "(" + out.paragraphs.length + " paragraf)";
  return "";
}

// =====================================================================
// FITUR 1: RIWAYAT CHAT PERSISTEN
// =====================================================================
const CHAT_STORAGE_KEY = "frida_chat_history";
const CHAT_MAX_MESSAGES = 50;

function saveChat() {
  try {
    // Simpan messages (untuk LLM) + HTML visual (untuk tampilan)
    var trimmed = messages.slice(-CHAT_MAX_MESSAGES);
    var chatHtml = logEl().innerHTML;
    localStorage.setItem(CHAT_STORAGE_KEY, JSON.stringify({
      messages: trimmed,
      html: chatHtml,
      ts: Date.now(),
    }));
  } catch (_) { /* localStorage penuh / tak tersedia — abaikan */ }
}

function restoreChat() {
  try {
    var raw = localStorage.getItem(CHAT_STORAGE_KEY);
    if (!raw) return;
    var data = JSON.parse(raw);
    if (data.messages && data.messages.length) {
      messages = data.messages;
    }
    if (data.html) {
      logEl().innerHTML = data.html;
      scrollDown();
    }
  } catch (_) { /* data rusak — mulai bersih */ }
}

function clearChat() {
  if (running) return;
  messages = [];
  logEl().innerHTML = "";
  localStorage.removeItem(CHAT_STORAGE_KEY);
  lastSnapshot = null;
  showUndo(false);
  setStatus("");
}

// =====================================================================
// FITUR 2: DARK MODE
// =====================================================================
const DARK_MODE_KEY = "frida_dark_mode";

function initDarkMode() {
  var toggle = document.getElementById("darkModeToggle");
  if (!toggle) return;
  // Load preferensi
  var saved = localStorage.getItem(DARK_MODE_KEY);
  if (saved === "true") {
    document.body.setAttribute("data-theme", "dark");
    toggle.checked = true;
  }
  // Toggle event
  toggle.addEventListener("change", function () {
    if (toggle.checked) {
      document.body.setAttribute("data-theme", "dark");
      localStorage.setItem(DARK_MODE_KEY, "true");
    } else {
      document.body.removeAttribute("data-theme");
      localStorage.setItem(DARK_MODE_KEY, "false");
    }
  });
}

// =====================================================================
// FITUR 3: TEMPLATE PERINTAH CEPAT (QUICK PROMPTS)
// =====================================================================
var QUICK_TEMPLATES = [
  { label: "📝 Format Skripsi", text: "Format seluruh dokumen ini sesuai panduan penulisan yang aktif (heading, font, spasi, margin)." },
  { label: "✏️ Perbaiki Ejaan", text: "Perbaiki semua typo, ejaan, dan tata bahasa di seluruh dokumen. Pertahankan makna asli." },
  { label: "📊 Buat Tabel", text: "Buatkan tabel yang relevan berdasarkan konteks dokumen ini." },
  { label: "📖 Daftar Pustaka", text: "Buatkan daftar pustaka / bibliography dari semua sumber yang telah diunggah." },
  { label: "📄 Cover Page", text: "Buatkan halaman sampul / cover page untuk dokumen ini." },
  { label: "🔢 Nomor Halaman", text: "Tambahkan nomor halaman di bagian bawah tengah." },
  { label: "📑 Daftar Isi", text: "Buatkan daftar isi berdasarkan struktur heading dokumen." },
];

function initQuickPrompts() {
  var container = document.getElementById("quickPrompts");
  if (!container) return;
  QUICK_TEMPLATES.forEach(function (tpl) {
    var btn = document.createElement("button");
    btn.className = "qp-btn";
    btn.textContent = tpl.label;
    btn.title = tpl.text;
    btn.onclick = function () {
      if (running) return;
      var input = document.getElementById("instruction");
      input.value = tpl.text;
      input.focus();
    };
    container.appendChild(btn);
  });
}

// =====================================================================
// FITUR 4: STATISTIK DOKUMEN
// =====================================================================
function updateDocStats() {
  try {
    Word.run(function (context) {
      var body = context.document.body;
      body.load("text");
      var paragraphs = body.paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        var text = body.text || "";
        var words = text.trim().split(/\s+/).filter(function (w) { return w.length > 0; });
        var chars = text.replace(/\s/g, "").length;
        var paraCount = paragraphs.items.length;
        // Estimasi waktu baca (200 kata/menit rata-rata)
        var readMins = Math.max(1, Math.round(words.length / 200));
        var readText = readMins < 60
          ? readMins + " mnt baca"
          : Math.floor(readMins / 60) + " jam " + (readMins % 60) + " mnt";

        var elWords = document.getElementById("statWords");
        var elPara = document.getElementById("statParagraphs");
        var elChars = document.getElementById("statChars");
        var elRead = document.getElementById("statReadTime");
        if (elWords) elWords.textContent = words.length.toLocaleString("id-ID");
        if (elPara) elPara.textContent = paraCount.toLocaleString("id-ID");
        if (elChars) elChars.textContent = chars.toLocaleString("id-ID");
        if (elRead) elRead.textContent = readText;
      });
    }).catch(function () {
      // Gagal baca dokumen — abaikan (mungkin tidak ada dokumen terbuka)
    });
  } catch (_) { /* Word.run tidak tersedia */ }
}

// =====================================================================
// FITUR 5: CONTEXT MEMORY (INGATAN KONTEKS)
// =====================================================================
const CONTEXT_MEMORY_KEY = "frida_context_memory";

function initContextMemory() {
  var textarea = document.getElementById("contextMemory");
  var saveBtn = document.getElementById("contextSave");
  var statusEl = document.getElementById("contextSaveStatus");
  if (!textarea || !saveBtn) return;

  // Load dari localStorage
  var saved = localStorage.getItem(CONTEXT_MEMORY_KEY);
  if (saved) textarea.value = saved;

  // Simpan
  saveBtn.onclick = function () {
    var val = textarea.value.trim();
    if (val) {
      localStorage.setItem(CONTEXT_MEMORY_KEY, val);
    } else {
      localStorage.removeItem(CONTEXT_MEMORY_KEY);
    }
    if (statusEl) {
      statusEl.textContent = "✓ Ingatan disimpan!";
      setTimeout(function () { statusEl.textContent = ""; }, 2500);
    }
  };
}

function getContextMemory() {
  try {
    return localStorage.getItem(CONTEXT_MEMORY_KEY) || "";
  } catch (_) { return ""; }
}
