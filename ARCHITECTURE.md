# FRIDA Agent — Production AI Word Copilot (Architecture)

> Dokumen desain untuk mengevolusi **FRIDA v1** (single-shot, satu tool `apply_changes`)
> menjadi **FRIDA Agent v2** (multi-tool, agentic loop, transaksi + rollback + audit).
> Ditulis sebagai evolusi dari kode yang ada (`server.js`, `taskpane.js`), bukan greenfield.
>
> Status implementasi terlacak di bagian **Development Roadmap** di bawah.

---

## 1. Executive Summary

**Kondisi saat ini (FRIDA v1).** Arsitektur sudah sehat: task pane → server proxy lokal
(HTTPS) → provider Anthropic, dengan satu *tool* terstruktur `apply_changes` yang membawa
`paragraphOps` + `tableOps`. Model mengembalikan rencana, `taskpane.js` menerapkannya lewat
Office.js. Pola "proxy menyimpan key, model menjawab via tool, klien mengeksekusi" sudah benar
dan jadi fondasi yang tepat.

**Keterbatasan yang menghambat skala.** Satu tool monolitik hanya bisa replace/insert teks dan
edit sel tabel. Tidak ada formatting, header/footer, gambar, ToC, track changes, dll. Eksekusi
langsung tanpa transaksi/rollback. Tidak ada multi-step planning (model harus muat seluruh
perubahan dalam satu respons). Tidak ada audit log atau permission.

**Target (FRIDA Agent v2).** Sistem agentic **multi-tool, multi-turn** di mana LLM memanggil
tool granular dari registry yang dapat diperluas, klien mengeksekusinya di dalam **boundary
transaksi Office.js** dengan **preview → confirm → execute → verify**, dan setiap aksi tercatat
untuk rollback dan audit. Kuncinya: pindah dari "model mengarang satu paket perubahan besar"
ke "model mengorkestrasi tool kecil yang teruji satu per satu".

**Tiga keputusan arsitektur inti:**
1. **Tool registry deklaratif** — satu sumber kebenaran untuk schema (dikirim ke LLM) +
   implementasi (dijalankan klien). Tambah kapabilitas Word = tambah satu entri, bukan ubah core.
2. **Agentic loop di server** — `tool_use` → klien eksekusi → `tool_result` → model lanjut,
   sampai model berhenti. Bukan one-shot.
3. **Transaction Manager di klien** — semua mutasi dalam `Word.run` dengan snapshot OOXML
   sebelum-sesudah, sehingga rollback = restore snapshot, bukan menebak invers.

---

## 2. System Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│  WORD HOST (Desktop / Web / Mac)                                  │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │  TASK PANE  (taskpane.html / .js)                          │   │
│  │  ┌────────────────┐  ┌──────────────────────────────────┐ │   │
│  │  │ UI Layer       │  │ Agent Client Runtime             │ │   │
│  │  │ - chat input   │  │ - conversation state             │ │   │
│  │  │ - plan preview │  │ - tool dispatcher                │ │   │
│  │  │ - diff/confirm │  │ - Transaction Manager (snapshot) │ │   │
│  │  │ - audit panel  │  │ - permission gate                │ │   │
│  │  └────────────────┘  └──────────────────────────────────┘ │   │
│  │                              │ Office.js (Word.run)         │   │
│  └──────────────────────────────┼──────────────────────────────┘   │
└─────────────────────────────────┼──────────────────────────────────┘
          │ HTTPS /api/agent       │ OOXML / Word Object Model
          ▼                        ▼
┌──────────────────────────────┐   ┌──────────────────────────────┐
│ LOCAL PROXY SERVER (server.js)│   │  WORD DOCUMENT (in-memory)   │
│ - holds API key (env)         │   │  body, paragraphs, tables,   │
│ - AGENTIC LOOP orchestrator   │   │  sections, contentControls…  │
│ - tool-schema injection       │   └──────────────────────────────┘
│ - rate limit / audit sink     │
└────────┬──────────────────────┘
         ▼
┌──────────────────────────────┐
│  LLM PROVIDER (Anthropic API) │
│  Messages API + tool_use loop │
└──────────────────────────────┘
```

**Pembagian tanggung jawab per layer:**

| Layer | Lokasi | Tanggung jawab | Status |
|---|---|---|---|
| **UI** | `taskpane.html/css` | Chat, preview diff, tombol confirm, panel audit | Ada (single-shot form) → perlu chat + preview |
| **AI / LLM** | Provider | Reasoning, planning, memilih tool | Ada (`callOnce`) |
| **Function Calling** | `server.js` | Inject tool schema, jalankan agentic loop, relay `tool_use`/`tool_result` | Parsial (one-shot, single tool) → perlu loop |
| **Tool Execution** | `taskpane.js` | Dispatch nama tool → handler Office.js | Parsial (`applyParagraphOps`/`applyTableOps`) → perlu registry |
| **Office.js Integration** | `taskpane.js` | `Word.run`, `context.sync`, proxy objects | Ada (`gatherContext`) |
| **Document Interaction** | Word host | Object model sebenarnya | N/A |

**Mengapa tool dieksekusi di klien, bukan server?** Office.js objects (`paragraph`, `table`,
`range`) hanya hidup di dalam `Word.run` di task pane — tidak bisa diserialisasi ke server. Jadi
server **hanya** mengorkestrasi percakapan + memegang key; **eksekusi tool wajib di klien**.
Ini sudah benar di FRIDA dan harus dipertahankan.

---

## 3. Tool Design

### Prinsip
1. **Granular tapi tidak atomik berlebihan.** `format_text(range,{bold,italic,…})` — bukan tool
   terpisah per properti. Satu tool per *kapabilitas Word*, banyak parameter.
2. **Range-addressable.** Hampir semua tool menerima `target` (selector range) seragam.
3. **Read tools + Write tools.** Agent harus bisa *membaca* sebelum *menulis*.
4. **Idempoten bila mungkin & verifiable.** Tiap write tool mengembalikan state hasil.

### Range Selector (kontrak bersama)
Mengatasi masalah indeks paragraf bergeser (lihat `taskpane.js` yang mengurut dari indeks besar
ke kecil). Selector deklaratif, klien meresolusi → `Word.Range[]`.

```jsonc
{
  "type": "object",
  "properties": {
    "mode": { "enum": ["selection","whole_document","paragraph_index",
                        "search","heading","style","range_id"] },
    "value": { "type": "string" },
    "index": { "type": "integer" },
    "occurrence": { "enum": ["first","all","nth"], "default": "all" },
    "n": { "type": "integer" }
  },
  "required": ["mode"]
}
```

### Tool catalog

| # | Tool | Domain | R/W |
|---|---|---|---|
| 1 | `get_document_outline` | Struktur | R |
| 2 | `search_text` | Search | R |
| 3 | `format_text` | Text formatting | W |
| 4 | `format_paragraph` | Paragraph | W |
| 5 | `apply_style` | Styles | W |
| 6 | `replace_text` | Search & replace | W |
| 7 | `insert_paragraph` | Konten | W |
| 8 | `manage_header_footer` | Header/footer | W |
| 9 | `set_page_numbers` | Page numbering | W |
| 10 | `insert_break` | Page/section break | W |
| 11 | `create_table` / `edit_table` | Tabel | W |
| 12 | `insert_image` | Gambar | W |
| 13 | `format_list` | List | W |
| 14 | `insert_toc` | Table of contents | W |
| 15 | `manage_comments` | Comments | W |
| 16 | `set_track_changes` | Track changes | W |
| 17 | `set_page_layout` | Layout | W |
| 18 | `insert_cover_page` | Composite | W |

### 3.1 Detail tool representatif

#### `get_document_outline` (read; landasan "understand document")
```jsonc
{
  "name": "get_document_outline",
  "description": "Baca struktur dokumen: paragraf (indeks, style, level heading, preview). Panggil INI DULU sebelum mengubah apa pun.",
  "input_schema": { "type":"object", "properties": {
    "include_text": { "type":"boolean", "default":false }
  }}
}
```
Output: `{ paragraphs:[{i,style,level,preview,isHeading}], sections, hasTables }`
```javascript
async function get_document_outline(context, args) {
  const paras = context.document.body.paragraphs;
  paras.load("items/text,items/styleBuiltIn,items/style");
  const tables = context.document.body.tables; tables.load("items");
  const sections = context.document.sections; sections.load("items");
  await context.sync();
  const cut = args.include_text ? 1e9 : 80;
  return {
    paragraphs: paras.items.map((p, i) => ({
      i, style: p.style,
      isHeading: /Heading/i.test(p.styleBuiltIn || ""),
      level: (p.styleBuiltIn.match(/Heading(\d)/i) || [])[1] || null,
      preview: (p.text || "").slice(0, cut),
    })),
    sections: sections.items.length,
    hasTables: tables.items.length,
  };
}
```

#### `format_text` (bold/italic/underline/font/color/size)
```jsonc
{
  "name": "format_text",
  "description": "Terapkan format karakter ke range target.",
  "input_schema": { "type":"object", "properties": {
    "target": { "$ref":"#/$defs/target" },
    "bold": {"type":"boolean"}, "italic": {"type":"boolean"},
    "underline": {"enum":["None","Single","Double","Thick","Dotted","Wavy"]},
    "fontName": {"type":"string"}, "fontSize": {"type":"number"},
    "color": {"type":"string"}, "highlightColor": {"type":"string"}
  }, "required":["target"] }
}
```
```javascript
async function format_text(context, args) {
  const ranges = await resolveTarget(context, args.target);
  ranges.forEach(r => {
    const f = r.font;
    if (args.bold      !== undefined) f.bold = args.bold;
    if (args.italic    !== undefined) f.italic = args.italic;
    if (args.underline !== undefined) f.underline = args.underline;
    if (args.fontName) f.name = args.fontName;
    if (args.fontSize) f.size = args.fontSize;
    if (args.color)    f.color = args.color;
    if (args.highlightColor) f.highlightColor = args.highlightColor;
  });
  await context.sync();
  return { applied: ranges.length };
}
```

#### `replace_text` (search & replace via Body.search)
```javascript
async function replace_text(context, args) {
  const results = context.document.body.search(args.find, {
    matchCase: args.matchCase, matchWholeWord: args.wholeWord
  });
  results.load("items"); await context.sync();
  results.items.forEach(r => r.insertText(args.replace, Word.InsertLocation.replace));
  await context.sync();
  return { replaced: results.items.length };
}
```

#### `set_page_numbers` (PAGE field via OOXML)
```javascript
async function set_page_numbers(context, args) {
  const footer = context.document.sections.getFirst()
                  .getFooter(Word.HeaderFooterType.primary);
  footer.clear();
  const p = footer.insertParagraph("", Word.InsertLocation.start);
  p.alignment = args.alignment || "Centered";
  const ooxml = args.format === "page_x_of_y"
    ? pageField("PAGE") + " of " + pageField("NUMPAGES")
    : pageField("PAGE");
  p.insertOoxml(wrapOoxml(ooxml), Word.InsertLocation.replace);
  await context.sync();
  return { ok: true };
}
```

#### `create_table` (data 2D atau dari seleksi)
```javascript
async function create_table(context, args) {
  let data = args.data, insertPoint;
  if (args.fromSelection) {
    const sel = context.document.getSelection();
    sel.load("text"); await context.sync();
    data = sel.text.split(args.rowDelimiter || "\n").filter(Boolean)
              .map(row => row.split(args.colDelimiter || "\t"));
    insertPoint = sel;
  } else {
    insertPoint = context.document.body.getRange(Word.RangeLocation.end);
  }
  const rows = data.length, cols = Math.max(...data.map(r => r.length));
  const table = insertPoint.insertTable(rows, cols,
    args.fromSelection ? Word.InsertLocation.replace : Word.InsertLocation.end, data);
  table.style = args.style || "Grid Table 4 - Accent 1";
  if (args.headerRow) table.getRow(0).font.bold = true;
  await context.sync();
  return { rows, cols };
}
```

### 3.2 Ringkas sisanya (pola identik)

| Tool | Office.js API inti | Catatan |
|---|---|---|
| `search_text` | `body.search(q, opts)` | Read-only |
| `format_paragraph` | `paragraph.alignment/.spaceBefore/.lineSpacing/.leftIndent` | pt |
| `apply_style` | `paragraph.style` / `styleBuiltIn` | bawaan vs custom |
| `insert_paragraph` | `body.insertParagraph` / `range.insertParagraph` | generalisasi `paragraphOps` |
| `manage_header_footer` | `section.getHeader(type).insertText/insertOoxml` | primary/firstPage/even |
| `insert_break` | `range.insertBreak(BreakType.page\|sectionNext, loc)` | section break |
| `edit_table` | `getCell(r,c)`, `addRows`, `deleteRow` | superset `applyTableOps` |
| `insert_image` | `range.insertInlinePictureFromBase64(b64, loc)` | server fetch+encode via jimp |
| `format_list` | `paragraph.startNewList()` / `attachToList(id,level)` | bullet & numbered |
| `insert_toc` | `body.insertOoxml(<TOC field>, loc)` | field |
| `manage_comments` | `range.insertComment(text)`, `comment.reply/.delete` | OfficeJS 1.4+ |
| `set_track_changes` | `document.changeTrackingMode` | off/trackAll/trackMineOnly |
| `set_page_layout` | `section.pageSetup` via OOXML | margin/orientasi/ukuran |
| `insert_cover_page` | composite (break+heading+styled+image) | tingkat tinggi |

---

## 4. Function Calling Schema (server orchestration)

Transformasi inti `server.js`: dari **one-shot single-tool** → **agentic loop multi-tool**.

```javascript
const TOOL_SCHEMAS = require("./tools/schemas"); // array {name,description,input_schema}

const SYSTEM_PROMPT = [
  "Nama Anda FRIDA, agen penyunting yang mengendalikan Microsoft Word lewat tool.",
  "ALUR WAJIB: (1) panggil get_document_outline dulu untuk memahami struktur.",
  "(2) Susun rencana minimal. (3) Panggil tool write satu per satu.",
  "(4) Untuk aksi destruktif/masif, jelaskan dulu lalu tunggu tool_result.",
  "Jangan mengubah yang tidak diminta. Pertahankan bahasa dokumen.",
].join("\n");

async function runAgentTurn(messages) {
  const transcript = [...messages];
  for (let step = 0; step < MAX_STEPS; step++) {
    const data = await callProvider(transcript, TOOL_SCHEMAS); // tool_choice: auto
    transcript.push({ role: "assistant", content: data.content });
    const toolUses = data.content.filter(b => b.type === "tool_use");
    if (toolUses.length === 0) return { done: true, transcript, finalText: textOf(data) };
    // server TIDAK eksekusi — kembalikan ke klien (Word.run)
    return { done: false, transcript, pendingToolUses: toolUses };
  }
}
```

**Protokol klien ⇆ server:**
```
POST /api/agent
  → { messages:[...], pendingResults?:[{tool_use_id, content}] }
  ← { done:false, assistantMsg, toolUses:[{id,name,input}] }
  ← { done:true, finalText, summary }
```

---

## 5. Office.js Implementation — Tool Dispatcher (klien)

```javascript
const HANDLERS = {
  get_document_outline, search_text, format_text, format_paragraph,
  apply_style, replace_text, insert_paragraph, manage_header_footer,
  set_page_numbers, insert_break, create_table, edit_table,
  insert_image, format_list, insert_toc, manage_comments,
  set_track_changes, set_page_layout, insert_cover_page,
};

async function executeToolUses(toolUses) {
  const results = [];
  await Word.run(async (context) => {
    const tx = await TransactionManager.begin(context); // snapshot OOXML
    try {
      for (const tu of toolUses) {
        const handler = HANDLERS[tu.name];
        if (!handler) { results.push(err(tu, "unknown tool")); continue; }
        if (!Permissions.allow(tu.name, tu.input)) { results.push(err(tu,"blocked")); continue; }
        const out = await handler(context, tu.input);
        AuditLog.record(tu, out);
        results.push({ tool_use_id: tu.id, content: JSON.stringify(out) });
      }
      await tx.commit();
    } catch (e) { await tx.rollback(); throw e; }
  });
  return results;
}

async function resolveTarget(context, target) {
  const body = context.document.body;
  switch (target.mode) {
    case "whole_document": return [body.getRange()];
    case "selection":      return [context.document.getSelection()];
    case "search": {
      const res = body.search(target.value, { matchCase:false });
      res.load("items"); await context.sync();
      return target.occurrence === "first" ? res.items.slice(0,1) : res.items;
    }
    case "heading": {
      const ps = body.paragraphs; ps.load("items/styleBuiltIn"); await context.sync();
      return ps.items.filter(p => /Heading/i.test(p.styleBuiltIn)).map(p => p.getRange());
    }
    case "paragraph_index": {
      const ps = body.paragraphs; ps.load("items"); await context.sync();
      return [ps.items[target.index].getRange()];
    }
  }
}
```

---

## 6. Agent Workflow

```
USER perintah
  → [1 UNDERSTAND] get_document_outline()        (read, auto, tanpa konfirmasi)
  → [2 PLAN]       model susun urutan tool_use + narasi rencana
  → [3 PREVIEW]    klien DRY-RUN: resolveTarget, hitung dampak, render diff (tanpa sync final)
  → [4 CONFIRM]    auto utk aman; WAJIB utk whole_document/replace_all/delete/track-off/>N range
  → [5 EXECUTE]    executeToolUses() dalam Word.run + TransactionManager
  → [6 VERIFY]     model terima tool_result → re-baca outline → "selesai" / koreksi
  → DONE + audit entry + tombol Undo
```

**Klasifikasi risiko → confirm:**
```javascript
function riskScore(toolUse) {
  const t = toolUse.name, a = toolUse.input; let s = 0;
  if (/whole_document/.test(JSON.stringify(a.target))) s += 3;
  if (t === "replace_text" && a.target?.mode === "whole_document") s += 2;
  if (/delete|clear/.test(t)) s += 4;
  if (t === "set_track_changes" && a.mode === "off") s += 3;
  return s; // >=3 → wajib konfirmasi
}
```

Dry-run memanfaatkan sifat Office.js: perubahan tidak permanen sampai `context.sync()` sukses.

---

## 7. Security Design

### 7.1 Kredensial (DONE — Fase 0)
- API key dibaca dari `process.env` (`.env`), bukan `config.json` ter-commit.
- `.env` + `config.json` ada di `.gitignore`; server menolak menyajikannya via HTTP.
- Key tidak pernah menyeberang ke klien/dokumen (pola proxy dipertahankan).

```javascript
const cfg = {
  apiKey:  process.env.AERO_API_KEY || fileCfg.apiKey,
  baseUrl: process.env.AERO_BASE_URL || fileCfg.baseUrl || "https://capi.aerolink.lat/",
  model:   process.env.FRIDA_MODEL   || fileCfg.model   || "claude-opus-4-8",
};
if (!cfg.apiKey) { console.error("Set AERO_API_KEY di .env"); process.exit(1); }
```

### 7.2 Permission controls (per-tool policy)
```javascript
const Permissions = {
  policy: { "set_track_changes":"confirm", "replace_text":"confirm_if_whole_doc",
            "delete_*":"confirm", "*":"allow" },
  allow(name, input) {
    const rule = this.policy[name] ?? this.policy["*"];
    if (rule === "deny") return false;
    if (rule === "confirm" || (rule==="confirm_if_whole_doc" && isWholeDoc(input)))
      return UI.requestConfirm(name, input);
    return true;
  }
};
```

### 7.3 Transaction Manager + Rollback (snapshot OOXML)
```javascript
const TransactionManager = {
  async begin(context) {
    const ooxml = context.document.body.getOoxml(); await context.sync();
    const snapshot = ooxml.value;
    return {
      async commit() {},
      async rollback() {
        context.document.body.clear();
        context.document.body.insertOoxml(snapshot, Word.InsertLocation.replace);
        await context.sync();
      }
    };
  }
};
```

### 7.4 Undo (dua lapis)
1. **Native** — jalankan satu batch dalam satu `Word.run`/`sync` → satu langkah undo Word.
2. **Snapshot** — tombol "Undo FRIDA" memanggil `rollback()` ke kondisi sebelum perintah.

### 7.5 Error handling
- Tiap handler dibungkus: gagal satu tool → `tool_result {error}` (model bisa koreksi).
- Error transaksional → rollback penuh.
- Map `RichApi.Error` (ItemNotFound/InvalidArgument) ke pesan ramah.
- Retry hanya untuk error jaringan provider (sudah ada di `callClaude`), JANGAN untuk eksekusi Word.

### 7.6 Audit logging
```javascript
const AuditLog = {
  entries: [],
  record(toolUse, output) {
    this.entries.push({ ts: new Date().toISOString(), tool: toolUse.name,
      input: redact(toolUse.input), result: summarize(output) });
    // enterprise: POST /api/audit (append-only, retensi)
  }
};
```

---

## 8. Sample Commands → tool calls

| Perintah | Tool call |
|---|---|
| "Make all headings blue and size 18" | `format_text({target:{mode:"heading",occurrence:"all"},color:"#0000FF",fontSize:18})` |
| "Add page numbers to every page" | `set_page_numbers({alignment:"Centered",format:"plain"})` |
| "Replace all Company A with Company B" | `replace_text({find:"Company A",replace:"Company B",target:{mode:"whole_document"}})` → **konfirmasi** |
| "Convert selected text into a table" | `create_table({fromSelection:true,colDelimiter:"\t",headerRow:true})` |
| "Create a professional cover page" | `insert_cover_page({title,subtitle,author,date,logo?})` |
| "Format as a professional business proposal" | `get_document_outline()` → `apply_style` → `set_page_layout` → `insert_cover_page` → `insert_toc` → `set_page_numbers` → `format_text` |

```javascript
async function insert_cover_page(context, a) {
  const body = context.document.body;
  const start = body.getRange(Word.RangeLocation.start);
  start.insertBreak(Word.BreakType.page, Word.InsertLocation.before);
  const title = body.insertParagraph(a.title, Word.InsertLocation.start);
  title.styleBuiltIn = Word.BuiltInStyleName.title; title.alignment = "Centered";
  if (a.subtitle) { const s = title.insertParagraph(a.subtitle, "After");
    s.styleBuiltIn = Word.BuiltInStyleName.subtitle; s.alignment = "Centered"; }
  body.insertParagraph(`${a.author||""}\n${a.date||""}`, "After");
  if (a.logoBase64) start.insertInlinePictureFromBase64(a.logoBase64, "Before");
  await context.sync();
  return { ok:true };
}
```

---

## 9. Scalability Recommendations

1. **Single source of truth untuk tool.** `{name,description,input_schema,handler}` sekali;
   generate schema-array untuk LLM + map handler klien dari objek yang sama.
2. **Lazy/grouped tool exposure.** >30 tool → tool router: model pilih kategori dulu,
   schema kategori itu yang dikirim (hemat token).
3. **OOXML sebagai escape hatch.** Fitur tanpa API tinggi (page-setup, fields, ToC, watermark)
   → `insertOoxml` + pustaka snippet.
4. **Batching `context.sync`.** Satu sync per batch, bukan per tool (sync mahal di Word web).
5. **Streaming + progres.** Stream narasi rencana; progress per tool ("3/7 selesai").
6. **Session & multi-dokumen.** State + audit per `document.url`; enterprise pakai session store (Redis).
7. **Cost control.** Outline default tanpa teks penuh; teks penuh hanya utk paragraf yang diubah;
   dokumen besar → chunk + retrieval.
8. **Testing.** Tiap handler diuji dgn harness + dokumen fixture; validasi schema vs input contoh.

---

## 10. Development Roadmap (status)

| Fase | Status | Deliverable | Basis kode |
|---|---|---|---|
| **0 — Hardening** | ✅ SELESAI | Key → env, `.env`/`.env.example`, `.gitignore`, path guard, git init | `config.json`, `server.js` |
| **1 — Tool registry** | ✅ SELESAI | Registry deklaratif `tools/`; 3 tool pertama (`get_document_outline`, `format_text`, `replace_text`); `resolveTarget`; selfcheck parity | `tools/`, `server.js` |
| **2 — Agentic loop** | ✅ SELESAI | `/api/agent` relay multi-turn; loop `tool_use`↔`tool_result` di klien; dispatcher via `resolveHandler`; chat UI; read auto / write konfirmasi | `server.js`, `taskpane.*`, `tools/handlers.js` |
| **3 — Safety core** | ✅ SELESAI | `tools/safety.js`: TransactionManager (snapshot OOXML) + rollback, Undo FRIDA, permission gate per-tool, `riskScore`/`needsConfirm`, AuditLog + panel | `tools/safety.js`, `taskpane.*` |
| **4 — Tool breadth** | ✅ SELESAI | **16 tool**. b1: set_page_layout, format_paragraph, apply_style, insert_break. b2: create_table, format_list, manage_header_footer, set_page_numbers, insert_image. b3: insert_toc, manage_comments, set_track_changes, edit_table | `tools/` |
| **5 — Preview/diff** | ⬜ | Dry-run preview + diff visual | `taskpane.js` |
| **6 — Composite & polish** | ⬜ | `insert_cover_page`, "business proposal", tool router, audit panel, streaming | baru |
| **7 — Enterprise** | ⬜ | Audit sink server, session store, policy per-tenant, sideload→AppSource | baru |

### Catatan Fase 0 (apa yang berubah)
- `server.js`: loader `.env` mini (tanpa dependency baru), `cfg` dari env → fallback `config.json`,
  exit bila `AERO_API_KEY` kosong, peringatan bila `config.json` masih memuat `apiKey`.
- `server.js` `serveStatic`: path guard pakai `path.resolve` + batas `path.sep` (bukan prefix string);
  tolak `.env`/`config.json` via HTTP.
- `config.json`: `apiKey` DIHAPUS (kini hanya nilai non-rahasia). File ini di-gitignore juga
  (karena masih bisa memuat `apiKey` sbg fallback); template ada di `config.example.json`.
- `.env`: berisi key asli (gitignored). `.env.example`: template tanpa nilai rahasia.
- `.gitignore`: `.env`, `config.json`, `node_modules/`, log, `*.pem`, dll.

### Catatan Fase 1 (apa yang berubah)
- **`tools/schemas.js`** — sumber kebenaran schema. Ekspor dual-mode: `module.exports` (Node) +
  `window.FRIDA_SCHEMAS` (browser). Selektor `target` didefinisikan sekali (`targetSchema`) lalu
  di-**inline** ke tiap tool via `withTarget()` — bukan `$ref/$defs`, karena banyak provider
  tool-calling tidak andal memproses `$ref`.
- **`tools/handlers.js`** — implementasi Office.js + `resolveTarget`. Dual-mode (`window.FRIDA_HANDLERS`).
  `resolveTarget` mendukung mode: selection, whole_document, paragraph_index, search (first/all/nth),
  heading, style.
- **`tools/selfcheck.js`** + `npm run check` — uji invarian: parity nama schema↔handler dua arah,
  bentuk schema valid, `resolveTarget` ada. Exit≠0 bila dilanggar (siap untuk CI/pre-commit).
- **`server.js`** — memuat `TOOL_SCHEMAS` dari registry & melaporkan jumlah tool saat start.
  `/api/edit` lama SENGAJA dibiarkan utuh (masih dipakai `taskpane.js` saat ini); endpoint agentic
  `/api/agent` yang memakai registry menyusul di Fase 2.
- **Belum** mengubah `taskpane.js`/UI — itu bagian Fase 2 (dispatcher klien + chat). Registry sudah
  bisa di-load browser via `window.FRIDA_*` begitu `taskpane.html` menyertakan kedua file `tools/`.

### Catatan Fase 2 (apa yang berubah)
- **`server.js`** — endpoint baru **`/api/agent`**: relay tipis & stateless. Klien mengirim
  SELURUH `messages` (termasuk `tool_result`), server memanggil provider sekali dengan
  `tools: TOOL_SCHEMAS` + `tool_choice: auto`, lalu mengembalikan `{stop_reason, content}` apa
  adanya. `AGENT_SYSTEM_PROMPT` mengarahkan "outline dulu → tool write satu per satu → jawaban
  teks akhir". `/api/edit` lama **tetap ada** (tidak ada regresi).
- **`taskpane.js`** — ditulis ulang jadi **loop agentic di klien** (`runAgentLoop`): kirim →
  terima → jika ada `tool_use`, eksekusi via `HANDLERS` di dalam satu `Word.run` → balas
  `tool_result` → ulangi; berhenti saat balasan tanpa `tool_use`. `MAX_STEPS=12` sbg batas aman.
- **Jembatan keamanan (sebelum Fase 3):** tool READ (`get_document_outline`, `search_text`)
  jalan otomatis; batch yang memuat tool WRITE **ditahan** dan butuh klik "Jalankan" (`confirmBatch`).
  Batal → kirim `tool_result is_error` agar model berhenti rapi.
- **`taskpane.html` / `taskpane.css`** — UI chat (gelembung user/assistant + kartu tool).
  HTML memuat `tools/schemas.js` & `tools/handlers.js` sebelum `taskpane.js`
  (registry tersedia via `window.FRIDA_*`).
- **Bug nyata yang ditemukan & diperbaiki:** provider/proxy me-RENAME nama tool di respons
  (mis. `get_document_outline` → `CompatGetDocumentOutline375718` — persis peringatan lama di
  `/api/edit`). Karena dispatcher multi-tool mencocokkan nama, ditambahkan **`resolveHandler`**
  (handlers.js): cocokkan via bentuk kanonik + containment, ambil yang terpanjang. Dipakai untuk
  dispatch DAN klasifikasi read/write di UI. Regresi ini diuji permanen di `selfcheck`.
- **Verifikasi end-to-end:** server boot OK; `/tools/*.js` tersaji (200), `.env` ditolak (403);
  `/api/agent` validasi input (400) dan panggilan nyata mengembalikan `stop_reason: tool_use`
  dengan model memilih `get_document_outline` lebih dulu (alur outline-first bekerja).

### Catatan Fase 3 (apa yang berubah)
- **`tools/safety.js`** (modul baru, dual-mode). Bagian PURE (diuji di Node):
  - `riskScore(toolUse)` — skor risiko; whole_document/replace-all/delete/track-off menaikkan skor.
    Implicit-whole-doc (mis. `replace_text` tanpa `target`) dihitung sbg seluruh dokumen.
  - `needsConfirm(toolUse)` — gabungan kebijakan + skor (ambang ≥3).
  - `Permissions` — kebijakan per-tool (`allow`/`confirm`/`deny`/`auto`) + pola `prefix_*` & `*`.
  - `makeAuditLog()` — pencatat append-only (tool, input, ok, ringkas hasil).
  Bagian butuh Word: `TransactionManager.begin(context)` ambil snapshot OOXML (`body.getOoxml`);
  `rollback()`/`restore(snapshot)` mengembalikan via `clear` + `insertOoxml`. Di Node = no-op.
- **`taskpane.js`** — terintegrasi safety:
  - Konfirmasi kini **berbasis risiko**, bukan "semua write". Read & write aman jalan otomatis;
    hanya aksi `needsConfirm` yang menahan + menampilkan banner ⚠️ aksi berisiko.
  - Batch write dieksekusi di atas **snapshot** (`TransactionManager.begin`). Bila ADA kegagalan
    di tengah batch → **rollback otomatis** seluruh batch (dokumen tak setengah jadi).
  - Batch sukses → snapshot disimpan sbg titik **Undo FRIDA**; tombol Undo memanggil `restore`.
  - Setiap aksi tercatat di **AuditLog**; panel "Riwayat aksi" menampilkan 8 terakhir.
  - Kebijakan `deny` memblokir aksi & membalas `tool_result is_error`.
- **`taskpane.html` / `.css`** — tombol "↩️ Undo perubahan terakhir" (muncul setelah write
  sukses) + panel `#audit`. Memuat `tools/safety.js` sebelum `taskpane.js`.
- **`tools/selfcheck.js`** — tambah uji regresi safety (riskScore read=0, format-selection aman,
  whole_document berisiko, replace-all perlu konfirmasi, kebijakan deny, AuditLog mencatat).
  Total **26 cek lulus**.
- **Verifikasi:** semua file `node -c` OK; globals browser (`FRIDA_SAFETY`) ter-set; `begin()` di
  Node no-op; server boot & menyajikan `/tools/safety.js` (200). Eksekusi rollback/Undo nyata
  butuh host Word (tak bisa diuji headless) — logikanya lurus: snapshot OOXML lalu `insertOoxml`.

### Catatan Fase 4 — batch 1 (apa yang berubah)
Registry tumbuh 3 → **7 tool**. Menjawab kebutuhan "ganti posisi halaman" yang sempat memblok.
- **`set_page_layout`** — orientasi (portrait/landscape), ukuran kertas (A4/Letter/Legal/A3/A5),
  margin (preset normal/narrow/moderate/wide atau `marginCm` manual). Office.js JS API TIDAK
  mengekspos `pageSetup`, jadi handler mengubah **OOXML `sectPr`** (`pgSz` utk ukuran/orientasi,
  `pgMar` utk margin; satuan twips, 1cm=567twip) lewat `getOoxml` → patch regex → `insertOoxml`.
  Karena ini replace OOXML seluruh body → `riskScore` +3 (wajib konfirmasi + ter-snapshot/rollback).
- **`format_paragraph`** — alignment, spaceBefore/After, lineSpacing, leftIndent, firstLineIndent.
  Butuh objek `Paragraph` (bukan Range), maka ditambah helper **`resolveTargetParagraphs`**.
- **`apply_style`** — set `paragraph.style` ke style bawaan (Heading 1/2, Title, Normal, …).
- **`insert_break`** — page break / sectionNext / sectionContinuous, posisi before/after target.
  Section break = prasyarat bila ingin orientasi berbeda antar bagian.
- **Verifikasi:** parity 7/7 (selfcheck **47 cek lulus**); unit-test patcher OOXML (landscape A4
  menukar dimensi; 1.27cm→720 twips); panggilan agent NYATA "ubah ke landscape" → model memilih
  `set_page_layout {orientation:"landscape"}`. Eksekusi OOXML di dokumen nyata perlu dites saat
  sideload (tak bisa headless).

> **CATATAN sideload:** `set_page_layout` mengganti OOXML seluruh body. Bila dokumen punya banyak
> section, patch berlaku ke SEMUA `sectPr`. Sudah ter-cover snapshot/rollback Fase 3 (ada tombol
> Undo), tapi mohon perhatikan hasilnya pada dokumen multi-section saat pengujian.

### Catatan Fase 4 — batch 2 (apa yang berubah)
Registry 7 → **12 tool**.
- **`create_table`** — dari `data` 2D atau **konversi teks terseleksi** (baris=baris-baru,
  kolom=tab/koma). Meratakan lebar kolom; header row ditebalkan; style opsional.
- **`format_list`** — bullet/number via `startNewList` + `attachToList`.
- **`manage_header_footer`** — tulis teks ke header/footer (section pertama).
- **`set_page_numbers`** — `PAGE` field (+`NUMPAGES` utk "x of y") di footer via OOXML
  (`fldSimple`), karena Office.js tak punya API field langsung.
- **`insert_image`** — `insertInlinePictureFromBase64`; untuk alur internal/komposit (model
  jarang mengirim base64).
- **Risiko:** tool batch-2 bersifat aditif/terlokalisasi (sisip tabel, beri list, tulis footer) &
  sudah ter-cover snapshot+Undo Fase 3 → dibiarkan jalan otomatis (tanpa friksi konfirmasi).
- **Verifikasi:** parity 12/12 (selfcheck **72 cek lulus**); agent call NYATA: "tambahkan nomor
  halaman" → `set_page_numbers`; "ubah teks seleksi jadi tabel" → `create_table {fromSelection:true}`.
  Eksekusi di dokumen nyata perlu dites saat sideload.

### Catatan Fase 4 — batch 3 (TUNTAS, 16 tool)
Registry 12 → **16 tool**. Fase 4 selesai.
- **`insert_toc`** — Daftar Isi via OOXML TOC field (`fldChar`+`instrText TOC`) ke BODY (andal,
  beda dgn footer). Pengguna klik kanan → Update Field utk mengisi.
- **`manage_comments`** — `Range.insertComment` pada target.
- **`set_track_changes`** — `document.changeTrackingMode` (trackAll/trackMineOnly/off). `off` → risiko +3.
- **`edit_table`** — pengganti `tableOps` lama FRIDA v1: edit sel, `addRows`, `deleteRowIndices`
  (hapus dari indeks besar ke kecil). Hapus baris → risiko +3.
- **Verifikasi:** parity 16/16 (selfcheck **95 cek lulus**); agent call nyata: "aktifkan track
  changes" → `set_track_changes {trackAll}`; "komentar: tolong revisi" → `manage_comments`;
  "daftar isi" → outline dulu lalu insert_toc.

### Catatan Fase 4 — nomor halaman di header (atas)
Dari pengujian: FRIDA menolak "nomor halaman di tengah atas" karena `set_page_numbers` di-hardcode
ke footer & header hanya bisa teks statis.
- **`set_page_numbers` + param `position` (top/bottom).** Field PAGE jalan sama baik di header,
  jadi `position:"top"` menaruh nomor BERJALAN (1,2,3…) di header — bukan teks statis. `bottom` = default.
- **System prompt** diperjelas: tool bisa atas/bawah + rata kiri/tengah/kanan; "tengah atas" →
  `position=top, alignment=Centered`; larang menolak permintaan nomor di atas.
- **Verifikasi:** "nomor halaman di tengah atas" → `set_page_numbers {position:"top",alignment:"Centered"}`.

### Catatan Fase 4 — perbaikan (bug dari pengujian sideload)
Ditemukan saat pengguna mencoba: `set_page_numbers` melempar **GeneralException** berulang, dan
"ubah posisi halaman" salah dirutekan ke `set_page_numbers`.
- **`set_page_numbers` ditulis ulang:** `insertOoxml` ke paragraf footer ternyata rapuh
  (GeneralException di Word desktop). Diganti **`Range.insertField(Word.FieldType.page)`** (+ `numPages`
  utk "x of y") — API resmi yang andal. Ada fallback teks bila host tak mendukung field. Helper
  OOXML field lama (`fldSimple`/`wrapRunOoxml`) dihapus.
- **Pelaporan error diperbaiki:** `officeErr()` di `taskpane.js` mengekstrak `.code` & `.debugInfo`
  dari `RichApi.Error`, jadi pesan bukan lagi sekadar "GeneralException" tetapi menyertakan
  kode & lokasi — berguna utk model maupun pengguna.
- **Routing tool diperjelas** di `AGENT_SYSTEM_PROMPT`: posisi/orientasi/ukuran/margin → `set_page_layout`;
  nomor halaman → `set_page_numbers`; + larangan mengulang tool yang sama saat error.
- **Verifikasi agent nyata:** "ubah posisi halaman" (ambigu) → model minta klarifikasi (bukan salah
  tool); "ganti orientasi jadi landscape" → `set_page_layout`; "beri nomor halaman" → `set_page_numbers`.

> **PENTING (rotasi key):** key lama sempat tersimpan plaintext di `config.json`. Karena belum
> pernah ter-commit ke git (repo baru di-init bersih), risikonya terbatas pada mesin lokal. Tetap
> disarankan **merotasi/regenerate API key** di dashboard provider bila file ini pernah tersalin
> ke tempat lain (backup, cloud sync, dsb).
