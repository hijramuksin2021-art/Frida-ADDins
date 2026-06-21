/* provider-ui.js — UI untuk Provider Settings di task pane.
   User bisa input Base URL + API Key, test koneksi, fetch model list, dan simpan settings.
   Semua tanpa restart server — perubahan langsung berlaku untuk pemanggilan AI berikutnya. */

(function () {
  let testedModels = []; // daftar model dari test terakhir yang berhasil
  let testedConfig = null; // {baseUrl, apiKey} yang terakhir di-test

  document.addEventListener("DOMContentLoaded", () => {
    const testBtn = document.getElementById("providerTest");
    const saveBtn = document.getElementById("providerSave");

    if (!testBtn) return; // panel tidak ada di halaman

    testBtn.addEventListener("click", testConnection);
    saveBtn.addEventListener("click", saveSettings);

    loadCurrentSettings();
  });

  async function loadCurrentSettings() {
    try {
      const resp = await fetch("/api/provider");
      const data = await resp.json();

      document.getElementById("providerBaseUrl").value = data.baseUrl || "";

      const keyInput = document.getElementById("providerApiKey");
      const keyHint = document.getElementById("providerKeyHint");
      if (data.hasKey && data.keyHint) {
        keyInput.placeholder = "Current: " + data.keyHint;
        keyHint.textContent = "Current key: " + data.keyHint;
      }

      // Jika ada model, tampilkan tapi Save tetap disabled sampai test berhasil
      const modelSelect = document.getElementById("providerModel");
      if (data.model) {
        const opt = document.createElement("option");
        opt.value = data.model;
        opt.textContent = data.model;
        opt.selected = true;
        modelSelect.appendChild(opt);
      }
    } catch (err) {
      providerStatus("Gagal muat pengaturan saat ini: " + (err.message || err), "err");
    }
  }

  async function testConnection() {
    const baseUrl = document.getElementById("providerBaseUrl").value.trim();
    const apiKey = document.getElementById("providerApiKey").value.trim();

    if (!baseUrl) {
      providerStatus("Base URL wajib diisi", "err");
      return;
    }

    const testBtn = document.getElementById("providerTest");
    testBtn.disabled = true;
    testBtn.textContent = "Testing...";
    providerStatus("Menghubung ke provider…");

    try {
      const resp = await fetch("/api/provider/test", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          baseUrl,
          apiKey: apiKey || undefined
        }),
      });
      const data = await resp.json();

      if (data.ok && data.models && data.models.length) {
        testedModels = data.models;
        testedConfig = { baseUrl, apiKey: apiKey || null };

        // Populate model dropdown
        const modelSelect = document.getElementById("providerModel");
        const modelLabel = document.getElementById("providerModelLabel");
        modelSelect.innerHTML = data.models.map(m =>
          '<option value="' + escHtml(m) + '">' + escHtml(m) + '</option>'
        ).join("");
        modelLabel.style.display = "flex";

        // Auto-select model pertama jika belum ada yang dipilih
        if (!modelSelect.value && data.models.length) {
          modelSelect.value = data.models[0];
        }

        providerStatus("✓ Terhubung — " + data.models.length + " model tersedia", "ok");
        document.getElementById("providerSave").disabled = false;
      } else {
        const errMsg = data.error || "Tidak ada model yang dikembalikan";
        providerStatus("Koneksi gagal: " + errMsg, "err");
        document.getElementById("providerSave").disabled = true;
      }
    } catch (err) {
      providerStatus("Koneksi gagal: " + (err.message || err), "err");
      document.getElementById("providerSave").disabled = true;
    } finally {
      testBtn.disabled = false;
      testBtn.textContent = "Tes Koneksi";
    }
  }

  async function saveSettings() {
    if (!testedConfig) {
      providerStatus("Test koneksi dulu", "err");
      return;
    }

    const model = document.getElementById("providerModel").value;
    if (!model) {
      providerStatus("Pilih model dulu", "err");
      return;
    }

    const saveBtn = document.getElementById("providerSave");
    saveBtn.disabled = true;
    saveBtn.textContent = "Menyimpan…";
    providerStatus("Menyimpan pengaturan…");

    try {
      const resp = await fetch("/api/provider", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          baseUrl: testedConfig.baseUrl,
          apiKey: testedConfig.apiKey,
          model: model,
        }),
      });
      const data = await resp.json();

      if (data.ok) {
        providerStatus("✓ Pengaturan tersimpan — provider diperbarui (tanpa restart)", "ok");
        // Kosongkan password field
        document.getElementById("providerApiKey").value = "";
        // Update hint key
        if (data.status && data.status.keyHint) {
          document.getElementById("providerKeyHint").textContent = "Key saat ini: " + data.status.keyHint;
        }
        // Reset tested state setelah save
        testedConfig = null;
        testedModels = [];
        saveBtn.disabled = true;
      } else {
        const errMsg = data.error || "kesalahan tidak diketahui";
        providerStatus("Gagal simpan: " + errMsg, "err");
      }
    } catch (err) {
      providerStatus("Gagal simpan: " + (err.message || err), "err");
    } finally {
      saveBtn.disabled = false;
      saveBtn.textContent = "Simpan Pengaturan";
    }
  }

  function providerStatus(msg, cls) {
    const el = document.getElementById("providerStatus");
    if (el) {
      el.textContent = msg || "";
      el.className = "provider-status" + (cls ? " " + cls : "");
    }
  }

  function escHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
})();
