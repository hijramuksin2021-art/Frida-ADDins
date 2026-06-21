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
| **R4** | resolve_source/alias, summarize_source, compare_sources | ✅ SELESAI |
| **R5** | Gaya MLA/Chicago/Harvard/IEEE, footnote, update-all via content control | ⬜ |
| **R6** | Faithfulness/NLI verify, quote-check, kalibrasi ambang | ⬜ |

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

## 17. Scalability
1. Store pluggable (`VectorStore` interface): sqlite-vec → pgvector/Qdrant.
2. Embedding batched + cache by chunk-hash; provider abstraksi (remote/lokal).
3. Ingest async dgn status. 4. Retrieval scoped per workspace. 5. CSL-JSON kanonik → tambah gaya = tambah file.
6. Pisah runtime server/client sejak awal → server-tools mudah dipindah ke worker.
