/* provider-ui.js — UI Pengaturan AI Provider MULTI-PROVIDER di task pane.
   User memilih provider (Anthropic / OpenAI / Gemini / Custom), mengisi API key khusus
   provider itu, memilih model, tes koneksi, lalu simpan. Tiap provider disimpan terpisah
   di server → pindah provider tidak menghapus key yang sudah diisi.
   Base URL hanya muncul untuk Custom (endpoint provider resmi dikunci di server). */

(function () {
  // Daftar model ter-kurasi (hardcoded) untuk provider resmi. Custom = fetch dinamis.
  const CURATED_MODELS = {
    anthropic: ["claude-opus-4-8", "claude-opus-4-7", "claude-sonnet-4-6", "claude-haiku-4-5"],
    openai: ["gpt-5", "gpt-5-mini", "gpt-4.1", "o4-mini"],
    gemini: ["gemini-2.5-pro", "gemini-2.5-flash"],
  };

  const KEY_PLACEHOLDER = {
    anthropic: "sk-ant-…",
    openai: "sk-…",
    gemini: "AIza…",
    custom: "sk-… (key 9Router/Aerolink)",
  };

  let statusData = null; // status multi-provider terakhir dari server

  document.addEventListener("DOMContentLoaded", () => {
    const sel = document.getElementById("providerSelect");
    if (!sel) return; // panel tidak ada

    sel.addEventListener("change", () => renderProvider(sel.value));
    document.getElementById("providerTest").addEventListener("click", testConnection);
    document.getElementById("providerSave").addEventListener("click", saveSettings);

    loadStatus();
  });

  async function loadStatus() {
    try {
      const resp = await fetch("/api/provider");
      statusData = await resp.json();
      const sel = document.getElementById("providerSelect");
      sel.value = statusData.activeProvider || "custom";
      await renderProvider(sel.value);
    } catch (err) {
      providerStatus("Gagal memuat pengaturan: " + (err.message || err), "err");
    }
  }

  // Sesuaikan form dengan provider terpilih (tanpa menyentuh provider lain).
  async function renderProvider(provider) {
    const st = (statusData && statusData.providers && statusData.providers[provider]) || {};

    // Base URL hanya untuk custom
    const baseUrlLabel = document.getElementById("providerBaseUrlLabel");
    const baseUrlInput = document.getElementById("providerBaseUrl");
    if (provider === "custom") {
      baseUrlLabel.style.display = "flex";
      baseUrlInput.value = st.baseUrl || "";
    } else {
      baseUrlLabel.style.display = "none";
    }

    // API key: field selalu kosong; tampilkan hint bila sudah tersimpan
    const keyInput = document.getElementById("providerApiKey");
    keyInput.value = "";
    keyInput.placeholder = KEY_PLACEHOLDER[provider] || "masukkan API key…";
    const keyHint = document.getElementById("providerKeyHint");
    keyHint.textContent = st.hasKey && st.keyHint ? "Key tersimpan: " + st.keyHint : "Belum ada key tersimpan";

    // Model: custom pakai input teks (bebas ketik + saran datalist), provider resmi pakai dropdown.
    const modelSelect = document.getElementById("providerModel");
    const modelText = document.getElementById("providerModelText");
    if (provider === "custom") {
      modelSelect.style.display = "none";
      modelText.style.display = "block";
      modelText.value = st.model || "";
      // Isi saran datalist dari model tersimpan; lalu coba fetch daftar (bila gateway mendukung).
      setModelSuggestions(st.model ? [st.model] : []);
      if (st.hasKey) fetchCustomModels(st.model);
    } else {
      modelText.style.display = "none";
      modelSelect.style.display = "block";
      setModelOptions(modelSelect, CURATED_MODELS[provider] || [], st.model);
    }

    providerStatus("");
  }

  // Nilai model aktif (dari input teks utk custom, atau dropdown utk provider resmi).
  function getModelValue() {
    const provider = document.getElementById("providerSelect").value;
    return provider === "custom"
      ? document.getElementById("providerModelText").value.trim()
      : document.getElementById("providerModel").value;
  }

  // Isi opsi saran (datalist) untuk input model custom.
  function setModelSuggestions(models) {
    const dl = document.getElementById("providerModelList");
    const list = models && models.length ? models : [];
    dl.innerHTML = list.map((m) => '<option value="' + escHtml(m) + '"></option>').join("");
  }

  function setModelOptions(select, models, selected) {
    const list = models && models.length ? models : [];
    if (!list.length) {
      select.innerHTML = '<option value="">-- (tes koneksi untuk memuat model) --</option>';
      return;
    }
    select.innerHTML = list.map((m) =>
      '<option value="' + escHtml(m) + '"' + (m === selected ? " selected" : "") + ">" + escHtml(m) + "</option>"
    ).join("");
    if (selected && list.indexOf(selected) >= 0) select.value = selected;
  }

  // Ambil daftar model dinamis untuk custom (endpoint OpenAI-compatible /v1/models).
  // Gateway seperti AgentRouter tak menyediakannya -> diamkan, user ketik model manual.
  async function fetchCustomModels(selected) {
    try {
      const resp = await fetch("/api/provider/models?provider=custom");
      const data = await resp.json();
      if (data.ok && data.models && data.models.length) {
        setModelSuggestions(data.models);
      }
    } catch (_) { /* diamkan; user bisa ketik model manual / Tes Koneksi */ }
  }

  async function testConnection() {
    const provider = document.getElementById("providerSelect").value;
    const apiKey = document.getElementById("providerApiKey").value.trim();
    const baseUrl = document.getElementById("providerBaseUrl").value.trim();

    if (provider === "custom" && !baseUrl) { providerStatus("Base URL wajib diisi untuk Custom", "err"); return; }
    const model = getModelValue();
    // Gateway custom yang cuma punya /v1/messages (mis. AgentRouter) divalidasi via ping,
    // jadi model wajib diisi agar bisa dites.
    if (provider === "custom" && !model) { providerStatus("Isi nama model dulu untuk Tes Koneksi", "err"); return; }

    const testBtn = document.getElementById("providerTest");
    testBtn.disabled = true;
    testBtn.textContent = "Menguji…";
    providerStatus("Menghubungi provider…");

    try {
      const resp = await fetch("/api/provider/test", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          provider,
          apiKey: apiKey || undefined,               // kosong = pakai key tersimpan
          baseUrl: provider === "custom" ? baseUrl : undefined,
          model: model || undefined,
        }),
      });
      const data = await resp.json();

      if (data.ok) {
        // Custom: isi saran datalist dari daftar model (bila gateway mengembalikannya).
        if (provider === "custom" && data.models && data.models.length > 1) {
          setModelSuggestions(data.models);
          providerStatus("✓ Terhubung — " + data.models.length + " model tersedia", "ok");
        } else {
          providerStatus("✓ Koneksi berhasil — API key & model valid", "ok");
        }
      } else {
        providerStatus("Koneksi gagal: " + (data.error || "kesalahan tak diketahui"), "err");
      }
    } catch (err) {
      providerStatus("Koneksi gagal: " + (err.message || err), "err");
    } finally {
      testBtn.disabled = false;
      testBtn.textContent = "🔌 Tes Koneksi";
    }
  }

  async function saveSettings() {
    const provider = document.getElementById("providerSelect").value;
    const apiKey = document.getElementById("providerApiKey").value.trim();
    const baseUrl = document.getElementById("providerBaseUrl").value.trim();
    const model = getModelValue();
    const st = (statusData && statusData.providers && statusData.providers[provider]) || {};

    if (!model) { providerStatus(provider === "custom" ? "Isi nama model dulu" : "Pilih model dulu", "err"); return; }
    if (provider === "custom" && !baseUrl) { providerStatus("Base URL wajib diisi untuk Custom", "err"); return; }
    if (!apiKey && !st.hasKey) { providerStatus("Isi API key dulu", "err"); return; }

    const saveBtn = document.getElementById("providerSave");
    saveBtn.disabled = true;
    saveBtn.textContent = "Menyimpan…";
    providerStatus("Menyimpan pengaturan…");

    try {
      const resp = await fetch("/api/provider", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          provider,
          activeProvider: provider,                   // provider terpilih jadi aktif
          apiKey: apiKey || undefined,                // kosong = jangan timpa key lama
          baseUrl: provider === "custom" ? baseUrl : undefined,
          model,
        }),
      });
      const data = await resp.json();

      if (data.ok) {
        statusData = data.status;                     // status multi-provider terbaru
        document.getElementById("providerApiKey").value = "";
        const cur = statusData.providers[provider] || {};
        document.getElementById("providerKeyHint").textContent =
          cur.hasKey && cur.keyHint ? "Key tersimpan: " + cur.keyHint : "Belum ada key tersimpan";
        providerStatus("✓ Tersimpan — provider aktif: " + provider + " (tanpa restart)", "ok");
      } else {
        providerStatus("Gagal simpan: " + (data.error || "kesalahan tak diketahui"), "err");
      }
    } catch (err) {
      providerStatus("Gagal simpan: " + (err.message || err), "err");
    } finally {
      saveBtn.disabled = false;
      saveBtn.textContent = "💾 Simpan Pengaturan";
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
