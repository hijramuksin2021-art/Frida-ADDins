# FRIDA Add-in Development Progress Note

## 1. Status Selesai (R0 - R7)
- **R0-R6**: Implementasi Core RAG (Ingestion, Search, Generate Grounded Paragraph), Citation Engine (APA7, MLA, Chicago, Harvard, IEEE), dan UI modern (tabbed navigation).
- **R7 (Guideline Profile)**:
  - Integrasi panduan penulisan per-fakultas/institusi.
  - Implementasi *automatic style enforcement* di backend dan *rule injection* di prompt LLM.
  - Fix bug: Wiring guideline aktif ke general chat prompt (dinamis via `getAgentSystemPrompt`).
  - Fix bug: Deteksi nama guideline otomatis dalam chat (fuzzy matching via `guideline-fuzzy.js`).

## 2. Selesai (Sesi Terakhir)
- **Update Ikon Ribbon**: Ikon lama diganti `new-icon.png` → resize ke `icon-16/32/80.png`. Manifest + cache-bust (`?v=2`) + clear Office Wef cache.
- **Header Task Pane**: Logo dihapus, judul "FRIDA" + subtitle dipindah ke tengah, header dibuat compact.
- **FIX: Batas langkah & efisiensi tool**: `MAX_STEPS`/`AGENT_MAX_STEPS` 12 → 40; prompt minta AI merencanakan & eksekusi dalam BATCH (target `heading`/`whole_document`, gabung properti).
- **FIX: Pembuatan tabel Bab 3**: Alias `insert_table` → `create_table`; prompt mewajibkan pembuatan tabel saat deteksi kata kunci (tabel/instrumen/variabel/rancangan) & Bab 3 Metode Penelitian.
- **FIX: Functional - SearchStringInvalidOrTooLong**: `replace_text` kini memakai strategi fallback anchor (potongan head/tail pendek) saat string pencarian melebihi ~240 char.
- **FIX: Layout - Action History covers input area**: `.audit` diubah `flex: 0 0 auto` → `flex: 0 1 auto` (boleh mengalah/shrink dengan scroll internal, max-height 160px) sehingga tidak pernah mendorong/menutupi `.composer`.

## 3. Task Pending (Bug Fixes)
- (kosong — semua bug fix terjadwal sudah selesai)

---
*Catatan: File ini ditulis sebagai konteks persistensi sebelum restart sesi.*
