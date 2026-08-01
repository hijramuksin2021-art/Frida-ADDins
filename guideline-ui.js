// guideline-ui.js — logika antarmuka untuk fitur Panduan Penulisan (Guideline Profile).
// Open-schema: menerima JSON struktur apapun, server yang membungkus metadata.

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

  function getAuthHeaders() {
    const token = localStorage.getItem("fridaAdminToken");
    const headers = { "Content-Type": "application/json" };
    if (token) headers["X-Frida-Token"] = token;
    return headers;
  }

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
        opt.textContent = gl.displayName || gl.id;
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
    const displayNameInput = document.getElementById("guidelineDisplayName");
    const activeDesc = document.getElementById("activeGuidelineDesc");

    if (!fileInput.files || fileInput.files.length === 0) {
      alert("Pilih file JSON terlebih dahulu!");
      return;
    }

    const file = fileInput.files[0];
    // Fallback displayName: dari input teks, lalu dari nama file tanpa .json
    const displayName = (displayNameInput && displayNameInput.value.trim())
      || file.name.replace(/\.json$/i, "");

    const reader = new FileReader();
    reader.onload = async function(e) {
      let jsonContent;
      try {
        jsonContent = JSON.parse(e.target.result);
      } catch (err) {
        if (activeDesc) activeDesc.innerHTML = `<span style="color:red;">File bukan format JSON yang valid, periksa kembali file kamu.</span>`;
        return;
      }

      activeDesc.innerHTML = "<em>Mengunggah...</em>";

      try {
        const resp = await fetch("/api/guidelines/upload", {
          method: "POST",
          headers: getAuthHeaders(),
          body: JSON.stringify({ displayName, content: jsonContent })
        });

        const data = await resp.json();
        if (resp.ok) {
          const name = escapeHtml(data.displayName || displayName);
          activeDesc.innerHTML = `<span style="color:green;">Pedoman '${name}' berhasil ditambahkan dan siap digunakan.</span>`;
          fileInput.value = "";
          if (displayNameInput) displayNameInput.value = "";
          await loadGuidelines();
        } else {
          activeDesc.innerHTML = `<span style="color:red;">Gagal: ${escapeHtml(data.error)}</span>`;
        }
      } catch (err) {
        activeDesc.innerHTML = `<span style="color:red;">Gagal mengunggah.</span>`;
      }
    };
    reader.readAsText(file);
  }

  async function handleDelete() {
    const guidelineSelect = document.getElementById("guidelineSelect");
    const activeDesc = document.getElementById("activeGuidelineDesc");
    const id = guidelineSelect.value;
    // Ambil nama dari opsi yang dipilih
    const selectedOption = guidelineSelect.options[guidelineSelect.selectedIndex];
    const name = selectedOption ? selectedOption.textContent : id;

    if (!id) {
      alert("Pilih pedoman yang ingin dihapus terlebih dahulu.");
      return;
    }
    if (!confirm(`Yakin ingin menghapus pedoman '${name}'?`)) return;

    activeDesc.innerHTML = "<em>Menghapus...</em>";
    try {
      const resp = await fetch(`/api/guidelines/${id}`, {
        method: "DELETE",
        headers: getAuthHeaders()
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
