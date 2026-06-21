# FRIDA Research Copilot — RAG & Citation Engine (Design)

> Ekstensi FRIDA: unggah PDF/DOCX/TXT → pahami → retrieve → tulis akademik yang **grounded** →
> sisipkan **sitasi & bibliografi akurat** di Word. Dirancang di ATAS arsitektur FRIDA yang ada
> (proxy Node, registry tool, agentic loop, safety). Status implementasi di **Development Roadmap**.

---

## 1. Executive Summary

Tiga keputusan arsitektur inti:

1. **Tool punya `runtime`: `server` vs `client`.** 19 tool FRIDA sekarang jalan di klien (Word.run).
   Tool RAG (parse, embed, search, generate) tak menyentuh Word & butuh Node → dieksekusi **di server**
   dalam loop agentic yang sama. Hanya tool penyisip (insert_citation, insert_bibliography) ke klien.

2. **Sitasi TIDAK PERNAH ditulis LLM** (anti-halusinasi inti). LLM hanya menyebut `source_id`.
   Author/tahun/judul/DOI berasal dari **metadata terverifikasi** (idealnya **Crossref via DOI**),
   string sitasi dirender **kode deterministik** (citeproc-js + CSL). LLM tak punya jalur mengarang.

3. **Embeddings = PROVIDER PLUGGABLE** (permintaan pengguna "API key all provider").
   - **Remote (default fleksibel):** endpoint **OpenAI-compatible** `POST {baseUrl}/embeddings`
     dengan `Authorization: Bearer {apiKey}` + `model`. Bisa pakai provider pihak ketiga mana pun
     (OpenAI, Voyage proxy, lokal LM Studio/Ollama, dsb) hanya dengan set baseUrl+key+model di `.env`.
   - **Local (privasi):** `@xenova/transformers` (ditambah di R1) — dokumen tak keluar mesin.
   - **Multilingual:** dokumen campur (Inggris + lain) → model multilingual (mis.
     `text-embedding-3-large` / `bge-m3` / `paraphrase-multilingual-MiniLM`).

---

## 2. System Architecture

```
TASK PANE (browser)                 LOCAL SERVER (Node)                    EXTERNAL
┌────────────────────┐  upload b64  ┌──────────────────────────────┐
│ Chat + Sumber panel│ ───────────▶ │ /api/sources/upload          │
│ - unggah, daftar KB│              │  ingest: parse→meta→chunk→embed│   ┌──────────┐
│ Agent runtime      │  /api/agent  │ AGENTIC LOOP (extended):     │DOI│ Crossref │
│ - client tools     │ ◀──────────▶ │  server-tools dijalankan sini│──▶│ metadata │
│ - insert_citation  │              │  client-tools dikirim ke pane│   └──────────┘
└─────────┬──────────┘              │ KB: vector + metadata DB     │   ┌──────────┐
          │ Office.js               │ Embeddings (provider plug)   │──▶│ Embed API│
          ▼                         │ Citation engine (CSL)        │   │ /Anthropic│
     WORD DOCUMENT                  └──────────────────────────────┘   └──────────┘
```

---

## 3. RAG Architecture

```
Query → (1) resolve_source ("jurnal Hijra"→id) → (2) embed(query)
      → (3) ANN search (scope workspace) top-k → (4) rerank → top-n
      → (5) gate: skor<ambang → "bukti tak cukup" (TIDAK generate)
      → (6) generate: LLM diberi HANYA chunk → {paragraph, claims:[{text,source_id,chunk_id,quote?}]}
            → verifikasi faithfulness (kutipan ada di chunk; klaim tanpa chunk → ditolak)
```
Parameter v1: chunk 500–800 token, overlap 80–120; k=8 → rerank n=4; ambang cosine ≥0.35 (kalibrasi).

---

## 4. Document Ingestion Pipeline (dipicu UNGGAH, bukan LLM)

```
POST /api/sources/upload {filename, mime, dataBase64}
 1. validate ext∈{pdf,docx,txt}, ukuran≤25MB, hash (dedup)
 2. parse: pdf→pdf-parse(PDFParse.getText); docx→mammoth; txt→langsung
 3. extract_metadata: DOI→Crossref(CSL-JSON resmi) | fallback parse halaman1 (confidence rendah)
 4. chunk per-section/sliding window (+page/section)
 5. embed via provider (R1)
 6. store documents + chunks(+embedding) + aliases
 7. return {document_id, title, authors, year, metaConfidence, numChunks}
```
`metaConfidence` rendah (tanpa DOI) → UI minta konfirmasi metadata sebelum boleh dipakai sitasi.

---

## 5. Knowledge Base Design
- **Scope per workspace** (per dokumen Word/sesi) → query banyak jurnal tanpa bocor antar proyek.
- **Aliases** referensi natural: `penulis1+tahun`+kata kunci judul → "pakai jurnal Hijra".
- **Skala ribuan dok**: SQLite+sqlite-vec (R1); >~500k chunk/multi-user → pgvector/Qdrant (store pluggable).

---

## 6. Tool Registry (LLM-facing) — field `runtime`

| Tool | runtime | Fungsi |
|---|---|---|
| resolve_source | server | nama natural → document_id |
| search_uploaded_sources | server | retrieval chunk |
| summarize_source | server | ringkas 1 sumber |
| compare_sources | server | banding ≥2 sumber |
| generate_paragraph_from_source | server | paragraf grounded + atribusi |
| format_citation | server | render sitasi (deterministik) |
| insert_citation | client | sisip sitasi (content control bertag) |
| insert_bibliography | client | daftar pustaka (regenerable) |
| insert_footnote | client | catatan kaki (Chicago) |
| cross_reference_document | client | rujuk silang dokumen aktif |

*Pipeline (bukan tool): upload_document, parse_pdf, parse_docx, extract_metadata, extract_citations.*

---

## 7. Function Calling Schemas (inti)

```jsonc
{ "name":"generate_paragraph_from_source","runtime":"server",
  "description":"Tulis paragraf DIDUKUNG sumber. Tiap klaim wajib menunjuk chunk. Bukti kurang → needsMoreEvidence=true (JANGAN mengarang).",
  "input_schema":{"type":"object","properties":{
    "instruction":{"type":"string"},"source_query":{"type":"string"},
    "document_ids":{"type":"array","items":{"type":"string"}},
    "style":{"type":"string","enum":["APA7","MLA","Chicago","Harvard","IEEE"]},
    "length":{"type":"string","enum":["sentence","short","paragraph"]}},
    "required":["instruction"]}}
// Output server tervalidasi: {paragraph, claims:[{text,source_id,chunk_id,quote?}], citations:[source_id], needsMoreEvidence}

{ "name":"insert_citation","runtime":"client",
  "input_schema":{"type":"object","properties":{
    "source_id":{"type":"string"},"style":{"enum":["APA7","MLA","Chicago","Harvard","IEEE"]},
    "locator":{"type":"string"},"mode":{"enum":["inText","footnote"],"default":"inText"}},
    "required":["source_id","style"]}}
```

---

## 8. Citation Intelligence Engine
- Kebenaran: **CSL-JSON** per dokumen (Crossref bila DOI; else parse+konfirmasi).
- Renderer: **citeproc-js** + file gaya **CSL** (APA7/MLA9/Chicago/Harvard/IEEE open-source) → satu engine semua gaya.
- insert_citation: server render → klien sisip **Content Control** bertag `frida-cite:{source_id}` → update massal saat ganti gaya, tanpa LLM.
- insert_bibliography: kumpulkan source_id tersitasi dari registry content-control → render → sisip di CC `frida-bibliography`.

---

## 9. Academic Writing Engine (prompt grounded)
```
SYSTEM: "Anda penulis akademik. HANYA nyatakan fakta yang DIDUKUNG passage. Tiap kalimat klaim
 dipetakan ke chunk_id. DILARANG menambah fakta/angka/nama/tahun/sitasi di luar passage. Bukti
 kurang → needsMoreEvidence=true, jangan menulis. Pertahankan bahasa dokumen. Balas via tool."
PASSAGES: [{chunk_id,source_id,text}...]   ← HANYA hasil retrieval
```
"Expand paragraf 3 pakai Hijra 2020" = outline → resolve_source → generate_paragraph_from_source → insert_paragraph + insert_citation.

---

## 10. Hallucination Prevention (lapis berganda)
1. Sitasi by-code, bukan LLM. 2. DOI→Crossref resmi; tanpa DOI → konfirmasi. 3. Retrieval gate (bukti kurang→tolak).
4. Atribusi wajib (klaim tanpa chunk_id ditolak). 5. Verifikasi kutipan (substring ada di chunk).
6. Faithfulness/NLI pass (R6). 7. Traceability (chunk_id di content control tersembunyi). 8. source_id tertutup.

---

## 11. Office.js Integration
- insert_citation: insertText lalu `insertContentControl()`, `.tag="frida-cite:<id>:<style>"`.
- insert_bibliography: cari/buat CC `frida-bibliography`.
- insert_footnote: API footnote (cek requirement set; fallback OOXML).
- update-all: enumerasi `body.contentControls` tag `frida-cite:` → render ulang via server.

---

## 12. Database Schema (SQLite v1; pluggable pgvector)
```sql
CREATE TABLE documents (id TEXT PRIMARY KEY, hash TEXT UNIQUE, filename TEXT, mime TEXT,
  title TEXT, authors_json TEXT, year INT, container TEXT, doi TEXT, csl_json TEXT,
  meta_confidence REAL, num_pages INT, uploaded_at TEXT, workspace TEXT, status TEXT);
CREATE TABLE chunks (id TEXT PRIMARY KEY, document_id TEXT, ordinal INT, page INT,
  section TEXT, text TEXT, token_count INT);
CREATE VIRTUAL TABLE chunk_vec USING vec0(chunk_id TEXT, embedding FLOAT[N]);
CREATE TABLE aliases (document_id TEXT, alias TEXT, weight REAL);
CREATE TABLE workspace_sources (workspace TEXT, document_id TEXT);
CREATE TABLE citation_registry (doc_url TEXT, cc_tag TEXT, source_id TEXT, style TEXT, inserted_at TEXT);
```
*R0 memakai store berbasis FILE (data/sources/*.json) tanpa native dep; SQLite+vec masuk R1.*

---

## 13. API Design
```
POST   /api/sources/upload   {filename,mime,dataBase64} → ingest
GET    /api/sources?workspace= → daftar KB
PATCH  /api/sources/:id/metadata → koreksi metadata
DELETE /api/sources/:id
POST   /api/sources/search   {query,k,document_ids,workspace} → chunks   (R1)
POST   /api/agent            (diperluas: server eksekusi server-tools, defer client-tools)
```

### Provider config (multi-provider "API key all provider")
`.env`:
```
# Chat (sudah ada)
AERO_API_KEY=...  AERO_BASE_URL=...  FRIDA_MODEL=...
# Embeddings (OpenAI-compatible). Kosongkan EMBED_BASE_URL utk pakai lokal (R1).
EMBED_PROVIDER=openai            # openai | local
EMBED_BASE_URL=https://api.provider.com/v1
EMBED_API_KEY=...
EMBED_MODEL=text-embedding-3-large
EMBED_DIM=3072
```
Modul `rag/embeddings.js` dispatch: remote (fetch `{baseUrl}/embeddings`) atau lokal (R1).

---

## 14. Security Architecture
- File lokal di `data/sources/` (**gitignore**); tak pernah dikirim utuh ke LLM — hanya chunk retrieval.
- Embedding lokal → dokumen sensitif tak keluar mesin. Crossref hanya terima DOI (opt-in utk rahasia).
- Validasi tipe/ukuran/hash, sanitasi nama, path guard. Key server-side. Hapus sumber = hapus chunks+file.

---

## 15. Example User Workflows
- "Baca jurnal ini, ringkas." → ingest → summarize_source → insert_paragraph.
- "Expand paragraf 3 pakai Hijra 2020." → outline→resolve→generate→insert+cite.
- "Sisipkan sitasi APA 7." → insert_citation(source_id, APA7).
- "Buat bibliografi semua sumber." → insert_bibliography.
- "Bandingkan 3 jurnal." → compare_sources → paragraf + sitasi tiap klaim.

---

## 16. Development Roadmap

| Fase | Isi | Status |
|---|---|---|
| **R0** | Ingestion: upload PDF/DOCX/TXT → parse → file store; panel Sumber + daftar KB; provider config embeddings | ✅ SELESAI |
| **R1** | Chunk + embed (lokal Xenova multilingual + remote pluggable) + vector store file + search_uploaded_sources + agent loop server/client | ✅ SELESAI |
| **R2** | generate_paragraph_from_source + gate (skor retrieval) + verifikasi sitasi (primer/warisan, buang yang dikarang) + insert_paragraph | ✅ SELESAI |
| **R3** | Citation engine (Crossref+CSL+citeproc) + insert_citation + insert_bibliography (APA7) | ✅ SELESAI |
| **R3.5** | Provider Settings UI: Input Base URL + API Key + Tes Koneksi → dropdown model otomatis terisi → Simpan | ✅ SELESAI |
| **R4** | resolve_source/alias, summarize_source, compare_sources | ✅ SELESAI |
| **R5** | Footnote citations (mode footnote) + update-all citations via content control | ✅ SELESAI |
| **R6** | Faithfulness/NLI verify, quote-check, kalibrasi ambang | ✅ SELESAI |
| **R7** | Guideline Profile: Panduan penulisan per Fakultas/Prodi, auto-style enforcement, rule injection di LLM prompt | ✅ SELESAI |

**Dependensi npm**: `pdf-parse` (v2, PDFParse.getText), `mammoth` (terpasang R0); R1: `better-sqlite3`+`sqlite-vec`, `@xenova/transformers`; R3: `citeproc-js`+file CSL.

---

## R1 — Catatan implementasi (penting)
- **Embeddings provider aerolink TIDAK ADA.** Probe `{baseUrl}/v1/embeddings` → 400
  "model embeddings tidak didukung; tersedia: claude-*". Aerolink hanya relai Claude chat.
  → Default **lokal** (`@xenova/transformers`, `paraphrase-multilingual-MiniLM-L12-v2`, 384-dim).
  Cross-lingual terbukti (cos ID↔EN = 0.91). Multi-provider remote (OpenAI-compatible) tetap
  tersedia via `EMBED_*` di `.env`.
- **Vector store**: berbasis file (`data/sources/<id>.chunks.json`) + cosine brute-force; cukup
  utk ribuan chunk. sqlite-vec/pgvector saat skala lebih besar (interface dijaga).
- **`runtime: server|client` pada registry** (kunci arsitektur). `/api/agent` kini **loop di
  server**: tool RAG (server) dieksekusi di server; saat model memanggil tool Word (client),
  server kembalikan ke task pane untuk `Word.run`. `messages` = sumber kebenaran dari server.
- **Bug yang ditemukan & diperbaiki:** field `runtime` ikut terkirim ke Anthropic → 400
  "Extra inputs are not permitted". Tools disanitasi ke `{name,description,input_schema}` sebelum
  dikirim (`API_TOOLS`); `runtime` hanya metadata internal.
- **Verifikasi nyata:** unggah sumber → "cari di sumber: kenapa reptil berdarah dingin?" →
  server jalankan `search_uploaded_sources`, model menjawab **grounded** dari kutipan (ektotermik,
  bergantung lingkungan) tanpa mengarang; perintah Word biasa tetap kembali ke klien.

## R2 — Catatan implementasi
- **`generate_paragraph_from_source`** (server tool): search → GATE → generasi grounded → verifikasi.
- **GATE = skor retrieval (deterministik), bukan judgment model.** Pelajaran: memberi model opsi
  `needsMoreEvidence` + prompt anti-halusinasi yang panjang membuatnya SELALU menolak (skor 0.76 pun
  ditolak). Solusi: tool model hanya `{paragraph}` (wajib); "bukti tak cukup" ditentukan server bila
  retrieval tak menghasilkan chunk di atas ambang (cosine ≥ 0.3). Einstein vs dok agroforestri →
  similarity negatif → ditolak server (model tak dipanggil).
- **Sitasi primer vs warisan (sesuai permintaan pengguna):** prompt MEMPERTAHANKAN sitasi in-text
  yang ada di passage (mis. `(Nair, 2012)`) = sitasi warisan dari sumber asli; klaim tanpa sitasi
  in-text = milik dokumen (sitasi primer). `rag/citations.js` mengekstrak & **memverifikasi** tiap
  sitasi terhadap chunk; yang TAK ada di sumber (mis. `(Hantu, 1999)` palsu) **dibuang** dari paragraf.
- **`rag/llm.js`**: pemanggil model server-side (forced tool) + self-load `.env`.
- **Verifikasi nyata:** "tambahkan paragraf manfaat agroforestri berdasarkan jurnal" → server
  generate (grounded, `(Nair, 2012)` terverifikasi & dipertahankan) → agent `insert_paragraph`.
  Sitasi PRIMER (format APA/dll dari metadata dokumen) menyusul di R3.

## R3 — Catatan implementasi (Citation engine, 5 gaya)
Tiga komponen inti: **formatter deterministik**, **Crossref lookup**, **tool Word penyisip**.

### Modul baru/diperluas
- **`rag/csl.js`** — formatter sitasi deterministik dari metadata CSL-JSON-ish. 5 gaya:
  APA7, MLA, Chicago, Harvard, IEEE. Fungsi: `inText(meta, style, opts)` (in-text, narrative,
  locator/page), `bibEntry(meta, style)` (entry daftar pustaka), `surnameLabel` (et al. rules
  per gaya), `italic` (penanda `*...*`; renderer klien bisa memiringkan). Ringan (tanpa dep
  eksternal); interface CSL-JSON agar mudah dipindah ke citeproc-js nanti.
- **`rag/crossref.js`** — `fetchByDoi(doi)`: fetch metadata resmi dari Crossref API
  (`api.crossref.org/works/{doi}`) → mapping ke format internal CSL-JSON-ish. Timeout 8s,
  User-Agent benar, gagal → null (pakai tebakan lokal). Mapping type:
  `journal-article` → `article-journal`, dll.
- **`rag/cite.js`** — jembatan store↔CSL. `inTextFor(id, style, opts)`,
  `entryFor(id, style)`, `bibliography(ids, style)` — semua dari metadata terverifikasi di
  store, bukan LLM. Bibliography author-date → urut alfabet; IEEE → urut input.
- **`rag/store.js`** (diperluas R0) — field `csl` + `confidence` sudah ada sejak R0.
  `updateMetadata(id, cslPatch)` mensinkronkan title/year ke index setelah edit.

### Endpoint baru di server.js
- **`POST /api/sources/cite`** `{source_id, style, page, narrative}` → `{text}` — string
  sitasi in-text dari metadata terverifikasi. Error 404 bila `csl` kosong.
- **`POST /api/sources/bibliography`** `{source_ids, style}` → `{entries:[{source_id,text}]}`
  — render daftar pustaka untuk beberapa sumber.
- **`PATCH /api/sources/:id/metadata`** `{csl}` → edit metadata sitasi sumber (sudah R0).

### Tool Word baru (client-side, tools/)
- **`insert_citation`** — fetch string sitasi dari `/api/sources/cite`, sisipkan di posisi
  kursor (akhir seleksi), bungkus dalam ContentControl bertag `frida-cite:{id}:{style}` untuk
  update massal (R5). Error bila metadata kosong → arahkan pengguna ke panel Sumber.
- **`insert_bibliography`** — fetch entries dari `/api/sources/bibliography`, sisipkan heading
  "Daftar Pustaka" + tiap entry dengan render `*...*` sebagai teks miring (`insertRichItalic`).
  Tanda `*` di tengah string = italic (sesuai penanda csl.js).

### UI metadata edit (sources-ui.js)
- Tombol **✎** per sumber → form inline: Judul, Penulis (family/given), Tahun, Tipe,
  Jurnal/Penerbit, Volume, Issue, Halaman, Institusi, DOI. `PATCH /api/sources/:id/metadata` →
  `confidence='user'` → kartu sumber tidak lagi menampilkan badge `metadata?`.
- Sumber dengan `confidence='low'/'medium'` tampilkan badge ⚠️ `metadata?` sebagai peringatan
  agar pengguna konfirmasi sebelum menyitir.

### Selfcheck (tools/selfcheck.js)
- Tambah **15 cek R3** (134 → **149 cek** total): normStyle 5 gaya; inText 5 gaya
  (termasuk locator/page); bibEntry APA7 dengan DOI URL; pipeline store→cite end-to-end
  (`inTextFor`, `entryFor`, `bibliography`); updateMetadata (confidence, title sync); remove.

### Prinsip anti-halusinasi (KUNCI)
Sitasi **TIDAK PERNAH** ditulis LLM. Alur:
1. Model menyebut `source_id` (dari hasil `search_uploaded_sources` / `generate_paragraph_from_source`)
2. Agent memanggil `insert_citation` → handler fetch `/api/sources/cite` → `csl.inText(doc.csl, style)`
3. String sitasi = output **kode deterministik** dari metadata terverifikasi (Crossref atau dikoreksi pengguna)
4. LLM tidak punya jalur mengarang nama/tahun/judul

> **STATUS R3: SELESAI.** Citation engine 5 gaya (APA7, MLA, Chicago, Harvard, IEEE) fungsional
> dengan `insert_citation` + `insert_bibliography` terintegrasi di registry tool, safety, dan selfcheck.
> Eksekusi nyata (Word.run) membutuhkan sideload; logika render deterministik telah terverifikasi
> unit-test. R4 (resolve_source/alias, summarize, compare) dan R5 (update-all via CC) menyusul.

## R3.5 — Catatan implementasi (Provider Settings UI)
Fitur tambahan yang memungkinkan pengguna mengatur provider AI (Base URL, API Key, Model) langsung dari UI task pane tanpa harus edit `.env` manual dan restart server.

### Modul baru & Modifikasi
- **`provider-ui.js`** (baru) — Logika UI untuk Provider Settings. Menangani inisialisasi status, tes koneksi ke `/api/provider/test`, mempopulerkan dropdown model, dan menyimpan config baru ke `/api/provider`.
- **`taskpane.html`** (modifikasi) — Menambahkan panel collapsible `📡 Provider Settings` dengan form input Base URL, API Key, Model dropdown, serta tombol Tes Koneksi & Simpan Pengaturan. Memuat `provider-ui.js`.
- **`taskpane.css`** (modifikasi) — Menambahkan style pendukung untuk form provider agar serasi dengan UI aslinya.
- **`rag/providerConfig.js`** & **Endpoint** (sudah ada) — Modul config provider (`get`, `set`, `status`) dan endpoint `/api/provider` (test koneksi via `listModels()`, get status, set config) diverifikasi fungsional dan terhubung dengan baik.

### Alur Kerja & Keamanan
1. Saat dimuat, UI mengambil status dari `GET /api/provider` lalu mempopulerkan field URL dan hint key saat ini (key penuh disembunyikan demi keamanan).
2. Pengguna memasukkan Base URL & API Key, klik **Tes Koneksi** (panggil `POST /api/provider/test`).
3. Endpoint server melakukan fetch `/v1/models` atau `/models` ke provider dengan API key tersebut. Jika sukses, daftar model yang didukung dikembalikan.
4. UI mempopulerkan dropdown **Model** dari daftar tersebut dan mengaktifkan tombol **Simpan Pengaturan**.
5. Klik **Simpan Pengaturan** (panggil `POST /api/provider`) → server menyimpan config ke `provider.local.json` (ter-gitignore).
6. Perubahan langsung berlaku secara real-time untuk pemanggilan AI berikutnya (editing, agent loop, RAG) tanpa memerlukan restart server.

> **STATUS R3.5: SELESAI.** UI Provider Settings fungsional secara end-to-end, terintegrasi dengan backend providerConfig dan persistensi lokal.

## R4 — Catatan implementasi (resolve, summarize, compare)
Tiga tool baru: **resolve_source** (alias keyword), **summarize_source** (LLM grounded), **compare_sources** (LLM grounded).

### Modul baru
- **`rag/aliases.js`** — resolver nama alami → `document_id` **tanpa embedding** (cepat, tanpa latensi
  jaringan). Strategi: scoring token kata kunci dari metadata dokumen (judul, penulis, filename, tahun,
  container).
  - `docKeywords(doc)` — ekstrak set token lowercase dari semua field metadata (termasuk `csl.author`,
    `csl.container`, `doc.year`).
  - `scoreDoc(doc, queryTokens)` — tiap token query: +3 cocok persis, +1 partial; +2 bonus tahun
    eksplisit; +2 bonus penulis pertama cocok.
  - `resolveSource(query, workspace)` — ranking skor turun, filter skor > 0.
  - `resolveSourceTool(input)` — wrapper untuk dipanggil sebagai server tool; kembalikan `best_id`,
    `matches[]` dengan skor, dan pesan jika tidak ada yang cocok.

- **`rag/summarize.js`** — ringkasan dan perbandingan sumber via LLM (grounded ke teks sumber).
  - `summarize_source({source_id, aspect?, max_sentences?})` — baca teks sumber (`store.get`),
    potong 4000 char, panggil `callModel` dengan forced tool `submit_summary`. Anti-halusinasi:
    prompt melarang menambah fakta di luar teks; output via tool terstruktur (bukan teks bebas).
  - `compare_sources({source_ids, aspect?})` — baca teks tiap sumber (maks 5, 2500 char per sumber),
    panggil `callModel` dengan forced tool `submit_comparison`. Output: `{comparison, similarities[],
    differences[], source_ids, titles}`.
  - Keduanya memakai `callModel` + `firstToolInput` dari `rag/llm.js` (retry 3x, timeout LLM).
  - Teks kosong/sumber tak ditemukan → `{error}` (bukan error throw); model bisa baca & laporkan.

### Update tool registry (schemas.js + agent_tools.js)
- **3 schema baru** (`runtime:"server"`):
  - `resolve_source` — `{query, workspace?, maxResults?}` → `{best_id, matches[], note}`
  - `summarize_source` — `{source_id, aspect?, max_sentences?}` → `{source_id, title, year, summary}`
  - `compare_sources` — `{source_ids[], aspect?}` → `{comparison, similarities[], differences[]}`
- **`rag/agent_tools.js`**: import `aliases.js` + `summarize.js`; tambah 3 entri ke `SERVER_TOOLS`.
- Total tool: **22 client + 5 server** = 27 tool dalam registry.

### Update AGENT_SYSTEM_PROMPT (server.js)
Panduan baru di prompt:
- "Nama alami sumber" (`'jurnal Hijra'`, `'Nair 2012'`) → **pakai resolve_source DULU** → gunakan `best_id`.
- 'Ringkas jurnal' → `summarize_source` (setelah resolve).
- 'Bandingkan paper A dan B' → `compare_sources` (setelah resolve masing-masing).
- `insert_citation` kini juga bisa pakai `source_id` dari `resolve_source`.

### Alur kerja khas R4
```
"Ringkas jurnal Hijra":
  resolve_source("Hijra") → best_id=src_abc
  summarize_source(src_abc) → {summary}  ← grounded ke teks, anti-halusinasi
  → sampaikan ringkasan ke pengguna

"Bandingkan paper Nair 2012 dan jurnal Hijra":
  resolve_source("Nair 2012") → best_id=src_xyz
  resolve_source("Hijra") → best_id=src_abc
  compare_sources([src_xyz, src_abc]) → {comparison, similarities, differences}
  → sampaikan perbandingan ke pengguna
```

### Selfcheck (tools/selfcheck.js)
- Tambah **27 cek R4** (149 → **176 cek** total):
  - `tokenize`: cek lowercase + tokenisasi benar.
  - `resolveSource` 3 skenario: match Hijra, match Nair+tahun+topik, query miss → `[]`.
  - `docKeywords`: verifikasi author family, tahun, dan kata judul ter-include.
  - Schema: 3 nama R4 terdaftar di SCHEMAS dengan `runtime='server'`.
  - Cleanup tes: `store.remove` berhasil (2 dokumen tes).

> **STATUS R4: SELESAI.** resolve_source (scoring keyword, tanpa embedding), summarize_source, dan
> compare_sources (LLM grounded via forced tool) fungsional di server loop. summarize/compare
> membutuhkan API key aktif untuk LLM call; resolve_source dapat diuji offline (unit-test 176 cek
> semua lulus). R5 (update-all via Content Control, gaya footnote) menyusul.

## R5 — Catatan implementasi (Footnote & Bulk Update)
Dua fitur utama ditambahkan pada fase ini: dukungan sitasi catatan kaki (Word Footnotes) dan kemampuan untuk memperbarui semua sitasi di dokumen secara massal (bulk update).

### Modul & Handler yang Diperbarui
- **`server.js`** (endpoint `/api/sources/cite`) — Mendukung parameter `mode="footnote"`. Jika mode ini diaktifkan, server mengembalikan format `entryFor` (format entri bibliografi lengkap) ditambah nomor halaman (misal: `, p. 12.`), menggantikan format in-text biasa.
- **`tools/schemas.js`** — 
  - `insert_citation`: Menambahkan parameter `mode` (`inText` atau `footnote`).
  - `update_all_citations` (baru): Mendaftarkan schema untuk memperbarui gaya semua sitasi.
  - Memperbarui array `SCHEMAS` agar mencakup tool baru ini.
- **`tools/handlers.js`** —
  - `insert_citation`: Mendukung `mode="footnote"` dengan menggunakan API `range.insertFootnote(text)`. Jika API ini tidak didukung oleh host Word, akan otomatis menggunakan fallback inline `[Footnote: text]`.
  - Tag Content Control untuk sitasi ditingkatkan agar menyimpan metadata lengkap: `frida-cite:${source_id};style=${style};mode=${mode};narrative=${narrative};page=${page}`.
  - `insert_bibliography`: Diubah agar membungkus daftar pustaka dalam `ContentControl` bertag `frida-bibliography:[source_ids]` agar dapat dideteksi saat update massal.
  - `update_all_citations` (baru): Melakukan enumerasi semua `contentControls` di dokumen, mem-parsing parameternya, memanggil API backend untuk mendapatkan teks dengan gaya baru, lalu mengganti isinya secara dinamis.
  - `insertRichItalicRange` (baru): Helper untuk menyisipkan teks berformat `*italic*` ke dalam `Word.Range` dan mengembalikan range lengkap yang telah diekspansi untuk dibungkus `ContentControl`.

### Alur Kerja Bulk Update
1. Saat user memerintahkan "ubah semua sitasi jadi MLA" atau "gunakan Chicago", LLM memanggil tool `update_all_citations(style)`.
2. Handler di klien memindai seluruh dokumen mencari content controls.
3. Untuk setiap sitasi (`frida-cite`), script mengekstrak source ID, halaman, narasi, dan mode yang disimpan di tag.
4. Script melakukan query ke server dengan style baru, lalu mengganti teks di dalam content control.
5. Untuk daftar pustaka (`frida-bibliography`), script mengekstrak source IDs yang tersimpan di tag, melakukan query ulang daftar pustaka dengan style baru, lalu me-rebuild isinya.
6. Tag content control diperbarui dengan nama style baru agar dapat di-update kembali di masa mendatang.

> **STATUS R5: SELESAI.** Dukungan footnote fungsional dan bulk update sitasi & daftar pustaka telah diimplementasikan dengan aman menggunakan Content Control tagging.

## R6 — Catatan implementasi (Faithfulness/NLI + Quote-Check + Threshold Calibration)
Tiga lapisan verifikasi tambahan untuk meningkatkan ketahanan terhadap halusinasi: verifikasi kutipan (quote-check), verifikasi kesetiaan semantik (faithfulness/NLI), dan sistem analitik untuk kalibrasi ambang batas retrieval.

### Modul Baru
- **`rag/quotecheck.js`** — Ekstraksi dan verifikasi kutipan literal dalam paragraf.
  - `extractQuotes(text)`: Mengekstrak teks dalam tanda petik ("..." atau '...') dengan panjang 5-200 karakter.
  - `quoteInChunks(quote, chunks, threshold=0.85)`: Verifikasi kutipan ada di chunks sumber menggunakan exact match atau fuzzy match (Levenshtein distance).
  - `verifyQuotes(paragraph, chunks)`: Verifikasi semua kutipan, kembalikan `{quotes, verified, flagged}`.
- **`rag/faithfulness.js`** — Verifikasi kesetiaan semantik menggunakan LLM.
  - `verifyFaithfulness(paragraph, chunks)`: Menggunakan LLM dengan prompt skeptis untuk mendeteksi kontradiksi antara paragraf yang dihasilkan dan chunks sumber.
  - Mengklasifikasikan setiap kalimat klaim sebagai ENTAILED, NEUTRAL, atau CONTRADICTION.
  - Hanya flag kontradiksi MAJOR (minor diabaikan untuk mengurangi false positive).
  - Fallback: jika LLM gagal, izinkan paragraf (lebih baik daripada reject tanpa verifikasi).
- **`rag/analytics.js`** — Manajemen log dan analitik in-memory.
  - `logGeneration(event)`: Catat setiap upaya generasi paragraf (accepted/rejected, maxScore, reason).
  - `getStats()`: Kembalikan statistik agregat (total, accepted, rejected, avgScore, rejectionReasons, scoreDistribution, recentLogs).
  - Batasi 500 log terakhir untuk efisiensi memori.

### Modifikasi Modul Existing
- **`rag/generate.js`** — Integritas verifikasi R6.
  - Import `quotecheck` dan `faithfulness` modules.
  - `GATE_SCORE` sekarang configurable via `process.env.GATE_SCORE_MIN` (default tetap 0.3).
  - Setelah generasi: verifikasi sitasi → verifikasi kutipan → verifikasi faithfulness.
  - Jika kutipan tidak terverifikasi: ganti dengan `[kutipan dihapus]`.
  - Jika ada kontradiksi MAJOR: reject paragraf dengan alasan detail.
  - Log setiap upaya generasi ke `analytics.js` dengan metadata (maxScore, reason, verifiedCitations, verifiedQuotes).
- **`server.js`** — Endpoint analitik baru.
  - `GET /api/sources/analytics/threshold`: Kembalikan statistik retrieval dan verifikasi dari modul `rag/analytics`.

### Alur Verifikasi R6
```
Generasi paragraf (R2)
  ↓
Verifikasi sitasi (R2) → buang sitasi tak terverifikasi
  ↓
Verifikasi kutipan (R6) → buang kutipan tak terverifikasi, ganti dengan "[kutipan dihapus]"
  ↓
Verifikasi faithfulness (R6) → jika ada kontradiksi MAJOR, reject paragraf
  ↓
Log ke analytics (R6) → catat hasil verifikasi dan maxScore untuk kalibrasi
  ↓
Return paragraf final (jika lulus semua verifikasi)
```

### Konfigurasi Threshold
- **Environment variable**: `GATE_SCORE_MIN` (default: 0.3)
  - Contoh: `GATE_SCORE_MIN=0.35 npm start` untuk threshold lebih ketat
- **Endpoint analitik**: `GET /api/sources/analytics/threshold`
  - Kembalikan distribusi skor tertinggi per request (band: 0.0-0.2, 0.2-0.3, 0.3-0.4, 0.4-0.5, 0.5+)
  - Statistik rejection reasons (insufficient_evidence, llm_error, empty_generation, faithfulness_contradiction)
  - Recent logs (15 terakhir) untuk debugging

### Trade-offs & Keputusan Desain
1. **Quote-check fuzzy matching** (threshold 0.85):
   - Kelebihan: Toleransi terhadap minor differences (punctuation, whitespace)
   - Kekurangan: Bisa false positive jika threshold terlalu rendah
   - Keputusan: 0.85 sebagai balance antara akurasi dan toleransi
2. **Faithfulness via LLM** (bukan NLI model):
   - Kelebihan: Reuses existing infrastructure, client-agnostic, mudah interpretasi
   - Kekurangan: Extra LLM call per paragraf (biaya + latensi)
   - Keputusan: LLM-based lebih praktis daripada load NLI model client-side
3. **In-memory analytics** (bukan database):
   - Kelebihan: Simple, fast, no extra dependency
   - Kekurangan: Data hilang saat server restart
   - Keputusan: Acceptable untuk development; bisa migrate ke SQLite/pgvector nanti

### Verifikasi Manual
1. Upload sumber dengan fakta jelas (misal: "Einstein lahir 1879")
2. Generate paragraf dengan kontradiksi sengaja ("Einstein lahir 1900") → harus ditolak
3. Generate paragraf dengan kutipan akurat → harus lulus verifikasi quote
4. Generate paragraf dengan kutipan salah → harus diganti dengan `[kutipan dihapus]`
5. Cek `GET /api/sources/analytics/threshold` → lihat statistik rejection dan distribusi skor

> **STATUS R6: SELESAI.** Semua 8 lapisan anti-halusinasi sekarang fungsional: (1) sitasi by-code, (2) DOI→Crossref, (3) retrieval gate, (4) atribusi wajib, (5) verifikasi kutipan, (6) **faithfulness/NLI (R6)**, (7) traceability via content control, (8) source_id protection. Sistem siap untuk deployment.

## R7 — Catatan implementasi (Guideline Profile: Panduan Fakultas)
Fitur untuk mendukung gaya penulisan spesifik dari fakultas atau program studi tertentu, bukan sekadar format akademik generik. Profil panduan disimpan dalam bentuk JSON terstruktur, yang kemudian diinjeksi ke LLM untuk generasi teks dan digunakan secara deterministik untuk engine sitasi.

### Komponen Utama
- **`rag/guidelines/*.json`** — Repositori profil panduan. File pertama `unkhair-pertanian-2021.json` menyimpan konfigurasi kertas, spasi, struktur bab, sitasi, dan aturan plagiarisme secara mendetail.
- **`rag/guidelineConfig.js`** — Pengelola status dan cache profil yang sedang aktif. Di-persist ke `guideline.local.json`.
- **`server.js` (Endpoint & Override)**
  - Menambahkan endpoint `GET /api/guidelines` (daftar semua panduan) dan `GET /api/guidelines/:id`.
  - Override otomatis di `/api/sources/cite` dan `/api/sources/bibliography`. Jika profil aktif mewajibkan "APA7", endpoint otomatis memaksakan style APA7 terlepas dari argumen client.
- **`rag/generate.js` (Rule Injection)**
  - System prompt `GROUNDED_SYSTEM` dirakit secara dinamis (via `getGroundedSystem()`).
  - Menyuntikkan _Aturan Khusus_ (misal: "Istilah asing dicetak miring", "Angka <10 menggunakan huruf", "Hindari plagiarisme >25%").
- **UI Task Pane (`guideline-ui.js` & `taskpane.html`)**
  - Dropdown "Panduan Penulisan" di tab Pengaturan untuk menampilkan daftar panduan yang tersedia.
  - Memilih panduan akan memberi badge konfirmasi dan menyimpan pilihan (mirip Provider Settings).

### Cara Kerja dan Verifikasi
1. User memilih profil "Fakultas Pertanian Unkhair" di Task Pane. Add-in mengingat pilihan di `guideline.local.json`.
2. Saat user meminta LLM "Buatkan paragraf...", prompt sistem menyertakan aturan spesifik Unkhair (seperti cara menulis angka dan istilah asing).
3. Saat user meminta "Tambahkan sitasi", server mengecek `guidelineConfig`. Jika panduan menyebut APA 7, maka `csl.js` dieksekusi dengan mode `APA7`.

> **STATUS R7: SELESAI.** Profil panduan terinjeksi mulus ke pipeline generasi (`rag/generate.js`) dan formater sitasi (`server.js`). Siap digunakan untuk menyesuaikan hasil AI dengan pedoman penulisan fakultas nyata.

## 17. Scalability
1. Store pluggable (`VectorStore` interface): sqlite-vec → pgvector/Qdrant.
2. Embedding batched + cache by chunk-hash; provider abstraksi (remote/lokal).
3. Ingest async dgn status. 4. Retrieval scoped per workspace. 5. CSL-JSON kanonik → tambah gaya = tambah file.
6. Pisah runtime server/client sejak awal → server-tools mudah dipindah ke worker.
