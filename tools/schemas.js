// tools/schemas.js — SUMBER KEBENARAN schema tool (ARCHITECTURE §9.1).
// File ini dipakai DUA tempat:
//   - server.js (Node)  -> dikirim ke LLM sebagai daftar `tools`.
//   - taskpane (browser) -> tidak wajib, tapi tersedia via window.FRIDA_SCHEMAS.
// Pasangan handler-nya ada di tools/handlers.js (nama tool HARUS sama).
//
// Catatan: schema `target` di-INLINE ke tiap tool (bukan $ref/$defs) karena
// banyak provider tool-calling tidak andal memproses $ref. Definisinya tetap
// satu (targetSchema) lalu disisipkan via helper withTarget().

(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api; // Node
  if (typeof window !== "undefined") window.FRIDA_SCHEMAS = api;             // Browser
})(this, function () {
  // ---- Selektor range bersama: dipakai hampir semua write-tool ----
  // Klien (handlers.resolveTarget) menerjemahkan ini -> Word.Range[].
  const targetSchema = {
    type: "object",
    description:
      "Penunjuk lokasi yang akan dikenai aksi. Pilih 'mode' yang sesuai instruksi pengguna.",
    properties: {
      mode: {
        type: "string",
        enum: [
          "selection",        // teks yang sedang diblok pengguna
          "whole_document",   // seluruh isi dokumen
          "paragraph_index",  // satu paragraf berdasarkan indeks (lihat get_document_outline)
          "search",           // semua/seq. kemunculan teks 'value'
          "heading",          // semua paragraf bergaya Heading*
          "style",            // semua paragraf dengan nama style = 'value'
        ],
        description:
          "selection=blok aktif; whole_document=seluruh dokumen; " +
          "paragraph_index=pakai 'index'; search=pakai 'value'; " +
          "heading=semua heading; style=pakai 'value' sbg nama style.",
      },
      value: { type: "string", description: "Kata kunci (mode=search) atau nama style (mode=style)." },
      index: { type: "integer", description: "Indeks paragraf (mode=paragraph_index)." },
      occurrence: {
        type: "string",
        enum: ["first", "all", "nth"],
        default: "all",
        description: "Untuk mode=search: kemunculan mana yang dikenai.",
      },
      n: { type: "integer", description: "Nomor kemunculan jika occurrence=nth (mulai 1)." },
    },
    required: ["mode"],
  };

  // helper: bangun input_schema dgn 'target' tersisip + properti lain
  function withTarget(props, required) {
    return {
      type: "object",
      properties: Object.assign({ target: targetSchema }, props),
      required: required || ["target"],
    };
  }

  // ---- 3 tool pertama (Fase 1) ----

  const get_document_outline = {
    name: "get_document_outline",
    description:
      "Baca struktur dokumen: daftar paragraf {indeks, style, level heading, preview teks}, " +
      "jumlah section, dan jumlah tabel. PANGGIL INI DULU sebelum mengubah apa pun agar tahu " +
      "indeks/target yang benar. Secara default teks dipangkas 80 karakter agar hemat token.",
    input_schema: {
      type: "object",
      properties: {
        include_text: {
          type: "boolean",
          default: false,
          description: "true = sertakan teks penuh tiap paragraf (mahal token). Default hanya 80 char.",
        },
      },
    },
  };

  const format_text = {
    name: "format_text",
    description:
      "Terapkan format KARAKTER ke range target: bold, italic, underline, nama font, " +
      "ukuran (pt), warna teks (hex #RRGGBB), dan warna highlight. Hanya set properti yang relevan.",
    input_schema: withTarget({
      bold: { type: "boolean" },
      italic: { type: "boolean" },
      underline: {
        type: "string",
        enum: ["None", "Single", "Double", "Thick", "Dotted", "Wavy"],
      },
      fontName: { type: "string", description: "Nama font, mis. 'Calibri'." },
      fontSize: { type: "number", description: "Ukuran dalam pt, mis. 18." },
      color: { type: "string", description: "Warna teks hex #RRGGBB." },
      highlightColor: { type: "string", description: "Warna highlight hex #RRGGBB." },
    }),
  };

  const replace_text = {
    name: "replace_text",
    description:
      "Cari dan ganti teks. Default di seluruh dokumen. Dukung match-case dan whole-word. " +
      "Gunakan untuk perintah seperti 'ganti semua A menjadi B'.",
    input_schema: {
      type: "object",
      properties: {
        find: { type: "string", description: "Teks yang dicari." },
        replace: { type: "string", description: "Teks pengganti." },
        matchCase: { type: "boolean", default: false },
        wholeWord: { type: "boolean", default: false },
        target: Object.assign({}, targetSchema, {
          description: "Batasi area cari-ganti. Default whole_document bila diabaikan.",
        }),
      },
      required: ["find", "replace"],
    },
  };

  // Daftar final (urutan = urutan yang dikirim ke LLM).
  const SCHEMAS = [get_document_outline, format_text, replace_text];

  return { SCHEMAS, targetSchema, byName: indexByName(SCHEMAS) };

  function indexByName(list) {
    const m = {};
    list.forEach((s) => (m[s.name] = s));
    return m;
  }
});
