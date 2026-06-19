/* sources-ui.js — panel "Sumber" (Research Copilot R0).
   Unggah PDF/DOCX/TXT (base64 JSON) -> /api/sources/upload, tampilkan daftar KB,
   hapus sumber. Retrieval & generasi grounded menyusul di R1+. */

(function () {
  const MAX_MB = 25;

  document.addEventListener("DOMContentLoaded", () => {
    const fileInput = document.getElementById("srcFile");
    const dropZone = document.getElementById("srcDrop");
    if (!fileInput) return; // panel tak ada

    fileInput.addEventListener("change", (e) => uploadFiles(e.target.files));
    if (dropZone) {
      ["dragover", "dragenter"].forEach((ev) =>
        dropZone.addEventListener(ev, (e) => { e.preventDefault(); dropZone.classList.add("over"); }));
      ["dragleave", "drop"].forEach((ev) =>
        dropZone.addEventListener(ev, (e) => { e.preventDefault(); dropZone.classList.remove("over"); }));
      dropZone.addEventListener("drop", (e) => uploadFiles(e.dataTransfer.files));
      dropZone.addEventListener("click", () => fileInput.click());
    }
    const reBtn = document.getElementById("srcReindex");
    if (reBtn) reBtn.addEventListener("click", reindex);
    refreshList();
    showEmbedStatus();
  });

  async function showEmbedStatus() {
    const el = document.getElementById("srcEmbed");
    if (!el) return;
    try {
      const s = await (await fetch("/api/sources/embed-status")).json();
      el.textContent = "embeddings: " + s.provider +
        (s.provider === "local" ? " (lokal, multilingual)" : "") +
        (s.configured ? "" : " — belum dikonfigurasi");
    } catch (_) {}
  }

  async function reindex() {
    const btn = document.getElementById("srcReindex");
    if (btn) btn.disabled = true;
    srcStatus("Mengindeks sumber (embedding)… pertama kali bisa lama (unduh model).");
    try {
      const r = await (await fetch("/api/sources/reindex", { method: "POST" })).json();
      const rows = r.result || [];
      const done = rows.filter((x) => x.numChunks != null);
      const skipped = rows.filter((x) => x.skipped);
      const errs = rows.filter((x) => x.error);
      const totChunks = done.reduce((a, x) => a + (x.numChunks || 0), 0);
      let msg;
      if (!rows.length) msg = "Belum ada sumber untuk diindeks.";
      else if (done.length) msg = "Terindeks: " + done.length + " dok baru (" + totChunks + " chunk)" +
        (skipped.length ? ", " + skipped.length + " sudah ada" : "") + ".";
      else msg = "Semua " + skipped.length + " sumber sudah terindeks ✓ — siap dicari.";
      if (errs.length) msg += " " + errs.length + " gagal.";
      srcStatus(msg, errs.length ? "err" : "ok");
    } catch (err) {
      srcStatus("Gagal indeks: " + (err.message || err), "err");
    } finally {
      if (btn) btn.disabled = false;
    }
  }

  function srcStatus(msg, cls) {
    const el = document.getElementById("srcStatus");
    if (el) { el.textContent = msg || ""; el.className = "src-status" + (cls ? " " + cls : ""); }
  }

  function readAsBase64(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(String(r.result).split(",")[1] || "");
      r.onerror = () => reject(new Error("gagal membaca file"));
      r.readAsDataURL(file);
    });
  }

  async function uploadFiles(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    for (const f of files) {
      if (f.size > MAX_MB * 1024 * 1024) { srcStatus(f.name + " > " + MAX_MB + "MB, dilewati", "err"); continue; }
      srcStatus("Mengunggah " + f.name + "…");
      try {
        const dataBase64 = await readAsBase64(f);
        const resp = await fetch("/api/sources/upload", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ filename: f.name, mime: f.type, dataBase64 }),
        });
        const data = await resp.json();
        if (!resp.ok) throw new Error(data.error || ("HTTP " + resp.status));
        srcStatus(data.duplicate ? f.name + " sudah ada" : f.name + " ditambahkan", "ok");
      } catch (err) {
        srcStatus("Gagal " + f.name + ": " + (err.message || err), "err");
      }
    }
    refreshList();
  }

  async function refreshList() {
    const box = document.getElementById("srcList");
    if (!box) return;
    try {
      const resp = await fetch("/api/sources");
      const data = await resp.json();
      const items = data.sources || [];
      const countEl = document.getElementById("srcCount");
      if (countEl) countEl.textContent = items.length ? "(" + items.length + ")" : "";
      if (!items.length) { box.innerHTML = '<div class="src-empty">Belum ada sumber. Unggah jurnal/PDF untuk mulai.</div>'; return; }
      box.innerHTML = items.map((s) => {
        const meta = [s.ext.toUpperCase(), s.year || null, s.pages ? s.pages + " hal" : null,
          Math.round((s.chars || 0) / 1000) + "k char"].filter(Boolean).join(" · ");
        const conf = s.confidence === "low"
          ? '<span class="src-warn" title="Metadata tebakan — perlu dikonfirmasi nanti">metadata?</span>' : "";
        return '<div class="src-item" data-id="' + s.id + '">' +
          '<div class="src-main"><div class="src-title">' + esc(s.title || s.filename) + " " + conf + "</div>" +
          '<div class="src-meta">' + esc(meta) + "</div></div>" +
          '<button class="src-del" data-id="' + s.id + '" title="Hapus">✕</button></div>';
      }).join("");
      box.querySelectorAll(".src-del").forEach((b) =>
        (b.onclick = () => removeSource(b.getAttribute("data-id"))));
    } catch (err) {
      box.innerHTML = '<div class="src-empty">Gagal memuat daftar sumber.</div>';
    }
  }

  async function removeSource(id) {
    try {
      await fetch("/api/sources/" + id, { method: "DELETE" });
      srcStatus("Sumber dihapus", "ok");
    } catch (_) {}
    refreshList();
  }

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
})();
