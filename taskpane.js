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
const MAX_STEPS = 12; // batas langkah loop (safety)

// Nama tool dari provider kadang di-rename (lihat resolveHandler di handlers.js).
// Resolusikan ke nama registry kanonik sebelum klasifikasi read/write & pelabelan.
function canonName(name) {
  const r = window.FRIDA_HANDLERS && window.FRIDA_HANDLERS.resolveHandler(name);
  return r ? r.name : name;
}
function isReadTool(name) { return READ_TOOLS.has(canonName(name)); }

let messages = []; // riwayat percakapan untuk LLM
let running = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("FRIDA hanya untuk Microsoft Word.", "err");
    return;
  }
  if (!window.FRIDA_HANDLERS || !window.FRIDA_SCHEMAS) {
    setStatus("Registry tool gagal dimuat (tools/schemas.js & handlers.js).", "err");
    return;
  }
  document.getElementById("send").onclick = onSend;
  document.getElementById("instruction").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); onSend(); }
  });
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

function addBubble(role, html) {
  const div = document.createElement("div");
  div.className = "msg " + role;
  div.innerHTML = html;
  logEl().appendChild(div);
  scrollDown();
  return div;
}
function escapeHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ---------- alur utama ----------
async function onSend() {
  if (running) return;
  const input = document.getElementById("instruction");
  const text = input.value.trim();
  if (!text) { setStatus("Tulis dulu perintahnya untuk FRIDA.", "err"); return; }

  addBubble("user", escapeHtml(text));
  messages.push({ role: "user", content: text });
  input.value = "";

  busy(true);
  setStatus("FRIDA berpikir…");
  try {
    await runAgentLoop();
    setStatus("Selesai.", "ok");
  } catch (err) {
    setStatus("Gagal: " + (err.message || err), "err");
    addBubble("error", "⚠️ " + escapeHtml(err.message || String(err)));
  } finally {
    busy(false);
  }
}

async function runAgentLoop() {
  for (let step = 0; step < MAX_STEPS; step++) {
    setStatus("FRIDA berpikir… (langkah " + (step + 1) + ")");
    const resp = await fetch("/api/agent", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ messages }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || ("HTTP " + resp.status));

    const content = Array.isArray(data.content) ? data.content : [];
    messages.push({ role: "assistant", content }); // simpan turn asisten apa adanya

    // tampilkan teks yang ada
    content.filter((b) => b.type === "text" && b.text && b.text.trim())
           .forEach((b) => addBubble("assistant", escapeHtml(b.text)));

    const toolUses = content.filter((b) => b.type === "tool_use");
    if (toolUses.length === 0) return; // jawaban akhir -> selesai

    // ada write tool? -> tahan untuk konfirmasi
    const hasWrite = toolUses.some((t) => !isReadTool(t.name));
    if (hasWrite) {
      const okGo = await confirmBatch(toolUses);
      if (!okGo) {
        // user batal: balas tool_result error agar model berhenti/menyesuaikan
        messages.push({ role: "user", content: toolUses.map((t) => ({
          type: "tool_result", tool_use_id: t.id, is_error: true,
          content: "Dibatalkan oleh pengguna.",
        })) });
        addBubble("error", "Dibatalkan. FRIDA tidak menerapkan perubahan.");
        return;
      }
    } else {
      renderToolCalls(toolUses, "membaca dokumen…");
    }

    setStatus("FRIDA menjalankan " + toolUses.length + " aksi…");
    const results = await executeToolUses(toolUses);
    messages.push({ role: "user", content: results });
  }
  addBubble("error", "Batas langkah tercapai (" + MAX_STEPS + ").");
}

// Eksekusi satu batch tool_use dalam SATU Word.run -> array tool_result.
async function executeToolUses(toolUses) {
  const { resolveHandler } = window.FRIDA_HANDLERS;
  const results = [];
  await Word.run(async (context) => {
    for (const tu of toolUses) {
      const resolved = resolveHandler(tu.name); // tahan nama yang di-rename provider
      if (!resolved) {
        results.push(toolResult(tu.id, { error: "tool tidak dikenal: " + tu.name }, true));
        continue;
      }
      try {
        const out = await resolved.fn(context, tu.input || {});
        results.push(toolResult(tu.id, out, false));
        markToolDone(tu.id, out);
      } catch (e) {
        const msg = (e && e.message) || String(e);
        results.push(toolResult(tu.id, { error: msg }, true));
        markToolDone(tu.id, { error: msg }, true);
      }
    }
  });
  return results;
}

function toolResult(id, obj, isError) {
  return { type: "tool_result", tool_use_id: id, is_error: !!isError,
           content: JSON.stringify(obj) };
}

// ---------- render kartu tool ----------
function confirmBatch(toolUses) {
  renderToolCalls(toolUses, "menunggu konfirmasi…");
  return new Promise((resolve) => {
    const wrap = document.createElement("div");
    wrap.className = "confirm";
    const yes = document.createElement("button");
    yes.className = "primary"; yes.textContent = "Jalankan " + toolUses.length + " aksi";
    const no = document.createElement("button");
    no.textContent = "Batal";
    yes.onclick = () => { wrap.remove(); resolve(true); };
    no.onclick = () => { wrap.remove(); resolve(false); };
    wrap.appendChild(yes); wrap.appendChild(no);
    logEl().appendChild(wrap);
    scrollDown();
  });
}

function renderToolCalls(toolUses, note) {
  toolUses.forEach((tu) => {
    const card = document.createElement("div");
    card.className = "tool";
    card.id = "tool-" + tu.id;
    const isRead = isReadTool(tu.name);
    card.innerHTML =
      '<div class="tag">' + (isRead ? "🔍 " : "✏️ ") + escapeHtml(canonName(tu.name)) + "</div>" +
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
