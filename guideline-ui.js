// guideline-ui.js — logika antarmuka untuk fitur Panduan Penulisan (Guideline Profile).
// Fitur ini mengambil daftar panduan dari server, menampilkannya di dropdown, dan
// menyimpan pilihan pengguna sehingga generasi AI dapat disesuaikan otomatis dengan panduan tersebut.

(function () {
  document.addEventListener("DOMContentLoaded", () => {
    const guidelineSelect = document.getElementById("guidelineSelect");
    if (!guidelineSelect) return;

    loadGuidelines();
    guidelineSelect.addEventListener("change", saveActiveGuideline);
  });

  async function loadGuidelines() {
    try {
      const resp = await fetch("/api/guidelines");
      const data = await resp.json();
      const guidelines = data.guidelines || [];

      const guidelineSelect = document.getElementById("guidelineSelect");
      if (!guidelineSelect) return;

      // Populate options
      guidelines.forEach(gl => {
        const opt = document.createElement("option");
        opt.value = gl.id;
        opt.textContent = gl.nama || gl.id;
        guidelineSelect.appendChild(opt);
      });

      // Panggil status aktif untuk memilih dropdown yang benar
      await updateActiveStatus();
    } catch (err) {
      console.error("Gagal memuat panduan penulisan:", err);
    }
  }

  async function updateActiveStatus() {
    try {
      const resp = await fetch("/api/guideline");
      const data = await resp.json();

      const guidelineSelect = document.getElementById("guidelineSelect");
      const activeDesc = document.getElementById("activeGuidelineDesc");

      if (data.activeId) {
        guidelineSelect.value = data.activeId;
        if (activeDesc) activeDesc.innerHTML = `<span class="guideline-badge">Aktif</span> ${escapeHtml(data.activeName)}`;
      } else {
        guidelineSelect.value = "";
        if (activeDesc) activeDesc.textContent = "Menggunakan format generik.";
      }
    } catch (err) {
      console.error("Gagal memuat status guideline aktif:", err);
    }
  }

  async function saveActiveGuideline(event) {
    const newId = event.target.value;
    const activeDesc = document.getElementById("activeGuidelineDesc");
    if (activeDesc) activeDesc.innerHTML = "<em>Menyimpan pengaturan...</em>";

    try {
      const resp = await fetch("/api/guideline", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id: newId })
      });
      if (resp.ok) {
        await updateActiveStatus();
      } else {
        const data = await resp.json();
        if (activeDesc) activeDesc.innerHTML = `<span style="color:red;">Gagal: ${escapeHtml(data.error)}</span>`;
      }
    } catch (err) {
      if (activeDesc) activeDesc.innerHTML = `<span style="color:red;">Gagal menyimpan.</span>`;
    }
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
})();
