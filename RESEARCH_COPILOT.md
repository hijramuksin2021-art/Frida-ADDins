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
| **R2** | generate_paragraph_from_source + gate + atribusi + traceability | ⬜ |
| **R3** | Citation engine (Crossref+CSL+citeproc) + insert_citation + insert_bibliography (APA7) | ⬜ |
| **R4** | resolve_source/alias, summarize_source, compare_sources | ⬜ |
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

## 17. Scalability
1. Store pluggable (`VectorStore` interface): sqlite-vec → pgvector/Qdrant.
2. Embedding batched + cache by chunk-hash; provider abstraksi (remote/lokal).
3. Ingest async dgn status. 4. Retrieval scoped per workspace. 5. CSL-JSON kanonik → tambah gaya = tambah file.
6. Pisah runtime server/client sejak awal → server-tools mudah dipindah ke worker.
