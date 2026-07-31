// guideline-ui.js — logika antarmuka untuk fitur Panduan Penulisan (Guideline Profile).
// Fitur ini mengambil daftar panduan dari server, menampilkannya di dropdown, dan
// menyimpan pilihan pengguna sehingga generasi AI dapat disesuaikan otomatis dengan panduan tersebut.

(function () {
  document.addEventListener("DOMContentLoaded", () => {
    const guidelineSelect = document.getElementById("guidelineSelect");
    const btnUpload = document.getElementById("btnUploadGuideline");
    const btnDelete = document.getElementById("btnDeleteGuideline");
    if (!guidelineSelect) return;

    loadGuidelines();
    guidelineSelect.addEventListener("change", saveActiveGuideline);
    if (btnUpload) btnUpload.addEventListener("click", handleUpload);
    if (btnDelete) btnDelete.addEventListener("click", handleDelete);
  });

  async function loadGuidelines() {
    try {
      const resp = await fetch("/api/guidelines");
      const data = await resp.json();
      const guidelines = data.guidelines || [];

      const guidelineSelect = document.getElementById("guidelineSelect");
      if (!guidelineSelect) return;

      // Populate options
      guidelineSelect.innerHTML = '<option value="">-- Tidak ada / Generik --</option>';
      guidelines.forEach(gl => {
        const opt = document.createElement("option");
        opt.value = gl.id;
        opt.textContent = (gl.nama || gl.id) + (gl.type ? ` (${gl.type})` : "");
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

  async function handleUpload() {
    const fileInput = document.getElementById("guidelineFileInput");
    const activeDesc = document.getElementById("activeGuidelineDesc");
    if (!fileInput.files || fileInput.files.length === 0) {
      alert("Pilih file JSON terlebih dahulu!");
      return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();
    reader.onload = async function(e) {
      try {
        const jsonContent = JSON.parse(e.target.result);
        activeDesc.innerHTML = "<em>Mengunggah...</em>";
        
        const resp = await fetch("/api/guidelines/upload", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(jsonContent)
        });
        
        const data = await resp.json();
        if (resp.ok) {
          activeDesc.innerHTML = `<span style="color:green;">Berhasil mengunggah pedoman!</span>`;
          fileInput.value = "";
          await loadGuidelines();
        } else {
          activeDesc.innerHTML = `<span style="color:red;">Gagal: ${escapeHtml(data.error)}</span>`;
        }
      } catch (err) {
        activeDesc.innerHTML = `<span style="color:red;">File JSON tidak valid.</span>`;
      }
    };
    reader.readAsText(file);
  }

  async function handleDelete() {
    const guidelineSelect = document.getElementById("guidelineSelect");
    const activeDesc = document.getElementById("activeGuidelineDesc");
    const id = guidelineSelect.value;
    
    if (!id) {
      alert("Pilih pedoman yang ingin dihapus terlebih dahulu.");
      return;
    }
    if (!confirm(`Yakin ingin menghapus pedoman '${id}'?`)) return;

    activeDesc.innerHTML = "<em>Menghapus...</em>";
    try {
      const resp = await fetch(`/api/guidelines/${id}`, {
        method: "DELETE"
      });
      const data = await resp.json();
      
      if (resp.ok) {
        activeDesc.innerHTML = `<span style="color:green;">Berhasil dihapus.</span>`;
        await loadGuidelines();
      } else {
        activeDesc.innerHTML = `<span style="color:red;">Gagal: ${escapeHtml(data.error)}</span>`;
      }
    } catch (err) {
      activeDesc.innerHTML = `<span style="color:red;">Gagal menghapus.</span>`;
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
