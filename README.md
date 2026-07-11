# FRIDA — Add-in Word pribadi (multi-provider AI)

**F**emale **R**eplacement **I**ntelligent **D**igital **A**ssistant — add-in untuk Microsoft Word yang:
- membaca **seluruh isi dokumen** (atau hanya bagian yang Anda seleksi),
- memperbaikinya sesuai instruksi Anda lewat **AI provider pilihan Anda**,
- **langsung menerapkan** hasilnya ke dokumen,
- plus Research Copilot: unggah referensi (PDF/DOCX/TXT) lalu menulis dari sumber.

Tidak perlu login akun mana pun — Anda memakai **API key provider Anda sendiri**, dan
key tidak pernah masuk ke dokumen (disimpan di server, gitignored).

---

## Provider yang didukung

FRIDA bisa dipindah antar provider **tanpa restart** lewat panel **⚙ Pengaturan → 🤖 AI Provider**:

| Provider | Endpoint | Yang perlu diisi |
|---|---|---|
| **Anthropic (Claude)** | dikunci ke `api.anthropic.com` | API key |
| **OpenAI** | dikunci ke `api.openai.com` | API key |
| **Google Gemini** | dikunci ke `generativelanguage.googleapis.com` | API key |
| **Custom** (OpenAI-compatible: 9Router / Aerolink / proxy lokal) | Base URL bebas | API key + Base URL |

Setiap provider menyimpan key & model-nya **terpisah**, jadi berpindah provider tidak
menghapus key yang sudah Anda isi. Untuk provider resmi, endpoint dikunci di server
(tidak bisa dioverride dari UI); hanya **Custom** yang punya Base URL bebas.

### AgentRouter / gateway ber-proteksi klien

Gateway Anthropic-compatible seperti **AgentRouter** (`https://agentrouter.org`) menolak
klien tak dikenal dengan `unauthorized client detected` — proteksi agar endpoint mereka
hanya dipakai klien resmi (Claude Code). FRIDA mengatasi ini untuk provider **Custom**
dengan mengirim header identitas ala Claude Code, memakai token gateway Anda sendiri.

Pakai: pilih provider **Custom**, isi **Base URL** = `https://agentrouter.org`, **API Key**
= token AgentRouter Anda, Tes Koneksi, pilih model, Simpan. Bila gateway mengubah kriteria
deteksinya, override lewat env `FRIDA_CLIENT_UA` / `FRIDA_ANTHROPIC_BETA` /
`FRIDA_CUSTOM_HEADERS` (lihat `.env.example`).

---

## Cara kerja (singkat)

```
Word (panel add-in)  ──►  https://localhost:3001  ──►  adapter multi-provider  ──►  Anthropic / OpenAI / Gemini / Custom
   Office.js                 server.js (proxy)         rag/aiProvider.js            (pakai API key Anda)
```

`server.js` menyajikan add-in lewat HTTPS sekaligus menyimpan API key. `rag/aiProvider.js`
menerjemahkan request/response internal (format Anthropic Messages) ke/dari format tiap
provider, jadi pemanggil tak perlu tahu provider mana yang aktif. Key tidak pernah masuk
ke dokumen.

---

## Persiapan (sekali saja)

1. **Buka Git Bash / PowerShell** di folder ini:
   ```
   cd C:\Users\iza\claude-word-addin
   ```

2. **Pasang dependency:**
   ```
   npm install
   ```

3. **Pasang sertifikat HTTPS tepercaya** (wajib, kalau tidak Word menolak add-in):
   ```
   npm run cert
   ```
   Akan muncul jendela konfirmasi Windows → klik **Yes**.

4. **(Opsional) siapkan `.env`** — salin `.env.example` menjadi `.env` lalu isi key
   awal. Semua nilai di sini opsional; Anda juga bisa mengisinya belakangan lewat panel
   Pengaturan tanpa restart. Nilai yang diubah dari UI tersimpan di `provider.local.json`
   (gitignored).

---

## Menjalankan

1. **Nyalakan server** (biarkan jendela ini terbuka selama memakai add-in):
   ```
   npm start
   ```
   Muncul:
   ```
   FRIDA berjalan di  https://localhost:3001/taskpane.html
   PUBLIC_URL: (default localhost) — set PUBLIC_URL untuk deploy publik, mis. Railway
   Provider : custom (key ✓)
   Model    : …
   ```

2. **Daftarkan add-in ke Word (sideload)** — cukup sekali:

   a. Buat folder kosong, misalnya `C:\WordAddins`, lalu **share** foldernya
      (klik kanan → Properties → Sharing → Share). Salin `manifest.xml` ke folder itu.

   b. Buka **Word → File → Options → Trust Center → Trust Center Settings →
      Trusted Add-in Catalogs**. Pada *Catalog Url* isi path share tadi
      (mis. `\\NAMA-PC\WordAddins`), klik **Add catalog**, centang
      **Show in Menu**, **OK**, lalu tutup & buka ulang Word.

   c. Di Word: **Insert → My Add-ins (Get Add-ins) → SHARED FOLDER →**
      pilih **FRIDA → Add**.

3. Di tab **Home** muncul tombol untuk membuka panel FRIDA.

---

## Mengatur AI provider (dari dalam add-in)

1. Buka panel FRIDA → tab **⚙ Pengaturan** → kartu **🤖 AI Provider**.
2. Pilih **AI Provider** (Anthropic / OpenAI / Gemini / Custom).
3. Isi **API Key**. Untuk **Custom**, isi juga **Base URL**.
4. Klik **🔌 Tes Koneksi** untuk memvalidasi key (dan memuat daftar model untuk Custom).
5. Pilih **Model**, lalu **💾 Simpan Pengaturan**. Provider yang dipilih langsung menjadi
   aktif **tanpa restart** — berlaku untuk chat/edit maupun generasi Research Copilot.

Field API key selalu tampil kosong demi keamanan; bila sudah ada key tersimpan akan
muncul hint (mis. `sk-a…om`). Mengosongkan field saat menyimpan **tidak** menimpa key lama.

---

## Memakai

1. Buka panel FRIDA, tab **💬 Chat**.
2. Tulis instruksi, mis. *"Perbaiki semua typo dan tata bahasa, buat lebih formal."*
3. (Opsional) blok/seleksi bagian yang dimaksud bila perintah menyebut "ini" / "di sini".
4. Centang **Tinjau** untuk melihat dampak sebelum menerapkan, lalu **Kirim**.

> Tip: setelah diterapkan, Anda tetap bisa **Ctrl+Z** di Word untuk membatalkan.

---

## Konfigurasi lewat `.env` (opsional)

Untuk mengatur nilai awal tanpa UI, salin `.env.example` → `.env`:

- `FRIDA_PROVIDER` — provider aktif saat start: `anthropic` | `openai` | `gemini` | `custom`
- `ANTHROPIC_API_KEY` / `OPENAI_API_KEY` / `GEMINI_API_KEY` — key provider resmi
- `AERO_API_KEY` / `AERO_BASE_URL` / `FRIDA_MODEL` — untuk provider **Custom**
- `FRIDA_ANTHROPIC_MODEL` / `FRIDA_OPENAI_MODEL` / `FRIDA_GEMINI_MODEL` — model default
- `FRIDA_MAX_TOKENS`, `FRIDA_PORT` — batas token & port server lokal
- `EMBED_*` — konfigurasi embeddings untuk Research Copilot / RAG

Prioritas nilai: **env (`.env`) → dioverride `provider.local.json` → dioverride simpan dari UI.**

> Catatan: proyek ini dahulu memakai `config.json` single-provider. Format itu sudah
> digantikan. File `provider.local.json` lama berformat flat (`{ baseUrl, apiKey, model }`)
> masih dibaca dan otomatis dipetakan ke provider **Custom** (setting lama tak hilang).

---

## Deploy publik (URL / Railway) & `PUBLIC_URL`

Secara default add-in disajikan di `https://localhost:3001` (development). Untuk memakainya
dari mana saja, deploy ke host publik (mis. **Railway**) lalu arahkan Word ke URL itu.

Semua URL absolut add-in bersumber dari satu env var **`PUBLIC_URL`**:

- Kosong / tidak di-set → otomatis fallback ke `https://localhost:3001` (dev lokal tetap jalan).
- Di-set (mis. `https://frida-addins-production.up.railway.app`, **tanpa trailing slash**) →
  dipakai untuk `manifest.xml` (SourceLocation, IconUrl, AppDomains, dll) dan log server.
- Di Railway, bila `PUBLIC_URL` kosong, `RAILWAY_PUBLIC_DOMAIN` dibaca otomatis.

`manifest.xml` **di-generate** dari `manifest.template.xml` (token `{{PUBLIC_URL}}`):

```
# set URL publik lalu regen manifest untuk sideload
PUBLIC_URL=https://<app>.up.railway.app  npm run manifest
```

`npm start` juga meregen manifest otomatis (lewat `prestart`). Selain itu server
menyajikan **`GET /manifest.xml` dinamis** — saat di-deploy, cukup arahkan Trusted Catalog /
sideload ke `https://<app>.up.railway.app/manifest.xml`, tak perlu regen manual.

> Frontend (`taskpane.html`, `provider-ui.js`, dll) memakai path relatif untuk semua
> `fetch`, jadi otomatis mengikuti origin tempat add-in disajikan — tidak ada URL yang
> perlu diubah di sisi klien.

Langkah ringkas deploy Railway:
1. Push repo ke GitHub (sudah), hubungkan ke Railway → deploy (`npm start`).
2. Di Railway → Variables, set `PUBLIC_URL` = domain publik app (tanpa `/` di akhir),
   plus API key provider yang dipakai (`ANTHROPIC_API_KEY` / dst).
3. Sideload manifest dari `https://<app>.up.railway.app/manifest.xml` ke Word.

---

## Catatan keamanan

- `.env` dan `provider.local.json` berisi API key Anda — **jangan dibagikan /
  di-screenshot / di-upload**. Keduanya sudah masuk `.gitignore`.
- Isi dokumen yang Anda proses akan dikirim ke server provider yang aktif. Untuk dokumen
  sangat sensitif, pertimbangkan provider dan ini terlebih dahulu.

---

## Masalah umum

| Gejala | Solusi |
|---|---|
| Word bilang add-in tidak aman / sertifikat | Jalankan `npm run cert` lalu restart Word |
| Panel blank / "Gagal" | Pastikan `npm start` masih berjalan di jendela terminal |
| "API key … belum di-set" | Isi & simpan key di ⚙ Pengaturan → AI Provider |
| "API key … tidak valid (401/403)" | Cek ulang key; untuk Custom cek juga Base URL |
| "Model atau endpoint tidak ditemukan (404)" | Periksa nama model di dropdown Model |
| "Respons bukan JSON valid" | Coba lagi, atau perpendek instruksi |
| Tombol add-in tak muncul | Ulangi langkah Trusted Catalog, restart Word |
