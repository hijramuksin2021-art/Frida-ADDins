# Claude Editor — Add-in Word pribadi

Add-in untuk Microsoft Word yang:
- membaca **seluruh isi dokumen** (atau hanya bagian yang Anda seleksi),
- memperbaikinya sesuai instruksi Anda lewat **Claude** (via provider aerolink),
- **langsung menerapkan** hasilnya ke dokumen.

Tidak perlu login akun Anthropic — karena memakai API key provider Anda sendiri (di `config.json`).

---

## Cara kerja (singkat)

```
Word (panel add-in)  ──►  https://localhost:3001  ──►  aerolink  ──►  Claude
   Office.js                 server.js (proxy)        (pakai API key Anda)
```

`server.js` menyajikan add-in lewat HTTPS sekaligus menyimpan API key, jadi key tidak
pernah masuk ke dokumen.

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

---

## Menjalankan

1. **Nyalakan server** (biarkan jendela ini terbuka selama memakai add-in):
   ```
   npm start
   ```
   Muncul: `Add-in Claude berjalan di https://localhost:3001/taskpane.html`

2. **Daftarkan add-in ke Word (sideload)** — cukup sekali:

   a. Buat folder kosong, misalnya `C:\WordAddins`, lalu **share** foldernya
      (klik kanan → Properties → Sharing → Share). Salin `manifest.xml` ke folder itu.

   b. Buka **Word → File → Options → Trust Center → Trust Center Settings →
      Trusted Add-in Catalogs**. Pada *Catalog Url* isi path share tadi
      (mis. `\\NAMA-PC\WordAddins`), klik **Add catalog**, centang
      **Show in Menu**, **OK**, lalu tutup & buka ulang Word.

   c. Di Word: **Insert → My Add-ins (Get Add-ins) → SHARED FOLDER →**
      pilih **Claude Editor → Add**.

3. Di tab **Home** muncul tombol **Claude Editor**. Klik untuk membuka panel.

---

## Memakai

1. Klik tombol **Claude Editor** di tab Home.
2. Tulis instruksi, mis. *"Perbaiki semua typo dan tata bahasa, buat lebih formal."*
3. (Opsional) centang **Hanya bagian yang saya seleksi** untuk membatasi ke teks terpilih.
4. Klik:
   - **Tinjau dulu** → lihat usulan perubahan tanpa mengubah dokumen, atau
   - **Baca & Perbaiki** → langsung terapkan ke dokumen.

> Tip: setelah diterapkan, Anda tetap bisa **Ctrl+Z** di Word untuk membatalkan.

---

## Mengubah pengaturan

Edit `config.json`:
- `apiKey` — API key provider Anda
- `baseUrl` — alamat provider (default aerolink)
- `model` — model yang dipakai (default `claude-opus-4-8`)
- `maxTokens`, `port`

---

## Catatan keamanan

- `config.json` berisi API key Anda — **jangan dibagikan / di-screenshot / di-upload**.
- Seluruh isi dokumen yang Anda proses akan dikirim ke server provider (aerolink).
  Untuk dokumen sangat sensitif, pertimbangkan ini dulu.

---

## Masalah umum

| Gejala | Solusi |
|---|---|
| Word bilang add-in tidak aman / sertifikat | Jalankan `npm run cert` lalu restart Word |
| Panel blank / "Gagal" | Pastikan `npm start` masih berjalan di jendela terminal |
| "Jawaban model bukan JSON valid" | Coba lagi, atau perpendek instruksi |
| Tombol add-in tak muncul | Ulangi langkah Trusted Catalog, restart Word |
