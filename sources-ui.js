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

  function getWs() { return window.FRIDA_WORKSPACE_ID || localStorage.getItem("frida_workspace_id") || "default"; }

  async function reindex() {
    const btn = document.getElementById("srcReindex");
    if (btn) btn.disabled = true;
    srcStatus("Mengindeks sumber (embedding)… pertama kali bisa lama (unduh model).");
    try {
      const resp = await fetch("/api/sources/reindex", { 
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ workspace: getWs() })
      });
      if (!resp.ok) {
        const errBody = await resp.text().catch(() => "");
        srcStatus("Gagal indeks ulang (server error " + resp.status + "): " + errBody.slice(0, 200), "err");
        return;
      }
      const r = await resp.json();
      if (r.error) {
        srcStatus("Gagal indeks ulang: " + r.error, "err");
        return;
      }
      const rows = r.result || [];
      const queued = rows.filter((x) => x.queued);
      const skipped = rows.filter((x) => x.skipped);
      const errs = rows.filter((x) => x.error);
      let msg;
      if (!rows.length) msg = "Belum ada sumber untuk diindeks.";
      else if (queued.length) msg = queued.length + " dokumen sedang diproses di background — status akan update otomatis." +
        (skipped.length ? " (" + skipped.length + " sudah terindeks)." : "");
      else msg = "Semua " + skipped.length + " sumber sudah terindeks ✓ — siap dicari.";
      if (errs.length) msg += " " + errs.length + " gagal.";
      srcStatus(msg, errs.length ? "err" : "ok");
      // Explicitly trigger a refresh right away to show pending status
      refreshList();
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
          body: JSON.stringify({ filename: f.name, mime: f.type, dataBase64, workspace: getWs() }),
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

  let pollTimer = null;

  async function refreshList() {
    const box = document.getElementById("srcList");
    if (!box) return;
    try {
      const resp = await fetch("/api/sources?workspace=" + encodeURIComponent(getWs()));
      const data = await resp.json();
      const items = data.sources || [];
      const countEl = document.getElementById("srcCount");
      if (countEl) countEl.textContent = items.length ? "(" + items.length + ")" : "";
      if (!items.length) { box.innerHTML = '<div class="src-empty">Belum ada sumber. Unggah jurnal/PDF untuk mulai.</div>'; return; }
      let hasPending = false;
      box.innerHTML = items.map((s) => {
        let statText = "";
        if (s.indexStatus === "pending") { statText = '<span style="color:#0078d4">Memproses (mengindeks)...</span>'; hasPending = true; }
        else if (s.indexStatus === "error") statText = '<span style="color:#d13438">Error: ' + esc(s.indexError || "Gagal mengindeks") + '</span>';
        else if (s.indexStatus === "done") statText = '<span style="color:#107c10">Siap dipakai</span>';

        const metaData = [s.ext.toUpperCase(), s.year || null, s.pages ? s.pages + " hal" : null,
          Math.round((s.chars || 0) / 1000) + "k char"].filter(Boolean).join(" · ");
        const meta = esc(metaData) + (statText ? " · " + statText : "");
        const conf = (s.confidence === "low" || s.confidence === "medium")
          ? '<span class="src-warn" title="Metadata tebakan — klik ✎ untuk koreksi sebelum menyitir">metadata?</span>' : "";
        return '<div class="src-item" data-id="' + s.id + '">' +
          '<div class="src-main"><div class="src-title">' + esc(s.title || s.filename) + " " + conf + "</div>" +
          '<div class="src-meta">' + meta + "</div></div>" +
          '<button class="src-edit" data-id="' + s.id + '" title="Edit metadata sitasi">✎</button>' +
          '<button class="src-del" data-id="' + s.id + '" title="Hapus">✕</button></div>' +
          '<div class="src-edit-form" id="ef-' + s.id + '"></div>';
      }).join("");

      if (hasPending) {
        if (!pollTimer) pollTimer = setTimeout(refreshList, 3000);
      } else {
        if (pollTimer) { clearTimeout(pollTimer); pollTimer = null; }
      }

      box.querySelectorAll(".src-del").forEach((b) =>
        (b.onclick = () => removeSource(b.getAttribute("data-id"))));
      box.querySelectorAll(".src-edit").forEach((b) =>
        (b.onclick = () => toggleEdit(b.getAttribute("data-id"), items)));
    } catch (err) {
      box.innerHTML = '<div class="src-empty">Gagal memuat daftar sumber.</div>';
    }
  }

  // ---- editor metadata sitasi ----
  function toggleEdit(id, items) {
    const box = document.getElementById("ef-" + id);
    if (!box) return;
    if (box.innerHTML) { box.innerHTML = ""; return; } // tutup bila terbuka
    const s = (items || []).find((x) => x.id === id) || {};
    const csl = s.csl || {};
    const a0 = (csl.author && csl.author[0]) || {};
    const f = (label, key, val) =>
      '<label class="ef-l">' + label + '<input class="ef-i" data-k="' + key + '" value="' + esc(val || "") + '"/></label>';
    box.innerHTML =
      '<div class="ef">' +
      f("Judul", "title", csl.title || s.title) +
      f("Penulis (nama belakang)", "family", a0.family) +
      f("Penulis (nama depan)", "given", a0.given) +
      f("Tahun", "year", (csl.issued && csl.issued.year) || s.year) +
      '<label class="ef-l">Tipe<select class="ef-i" data-k="type">' +
      ["article-journal", "thesis", "book", "chapter", "paper-conference"].map((t) =>
        '<option value="' + t + '"' + (csl.type === t ? " selected" : "") + ">" + t + "</option>").join("") +
      "</select></label>" +
      f("Jurnal/Penerbit (container)", "container", csl.container) +
      f("Volume", "volume", csl.volume) + f("Issue", "issue", csl.issue) +
      f("Halaman", "page", csl.page) + f("Institusi (skripsi)", "institution", csl.institution) +
      f("DOI", "DOI", csl.DOI) +
      '<button class="ef-save" data-id="' + id + '">Simpan metadata</button>' +
      "</div>";
    box.querySelector(".ef-save").onclick = () => saveMetadata(id, box);
  }

  async function saveMetadata(id, box) {
    const vals = {};
    box.querySelectorAll(".ef-i").forEach((el) => (vals[el.getAttribute("data-k")] = el.value.trim()));
    const csl = {
      type: vals.type || "article-journal",
      title: vals.title || null,
      author: vals.family ? [{ family: vals.family, given: vals.given }] : [],
      issued: { year: vals.year ? Number(vals.year) : null },
      container: vals.container || null,
      volume: vals.volume || null, issue: vals.issue || null, page: vals.page || null,
      institution: vals.institution || null, DOI: vals.DOI || null,
    };
    srcStatus("Menyimpan metadata…");
    try {
      const r = await fetch("/api/sources/" + id + "/metadata", {
        method: "PATCH", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ csl, workspace: getWs() }),
      });
      if (!r.ok) throw new Error((await r.json()).error || "gagal");
      srcStatus("Metadata tersimpan — sitasi siap dipakai.", "ok");
      box.innerHTML = "";
      refreshList();
    } catch (err) { srcStatus("Gagal simpan: " + (err.message || err), "err"); }
  }

  async function removeSource(id) {
    try {
      await fetch("/api/sources/" + id + "?workspace=" + encodeURIComponent(getWs()), { method: "DELETE" });
      srcStatus("Sumber dihapus", "ok");
    } catch (_) {}
    refreshList();
  }

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
})();
