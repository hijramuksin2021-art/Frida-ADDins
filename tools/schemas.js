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

  // ---- Tool tambahan (Fase 4) ----

  const set_page_layout = {
    name: "set_page_layout",
    description:
      "Atur tata letak halaman dokumen: orientasi (portrait/landscape), ukuran kertas " +
      "(A4/Letter/Legal), dan/atau margin (preset normal/narrow/moderate/wide, atau angka cm). " +
      "Gunakan untuk perintah seperti 'ubah ke landscape', 'ganti kertas ke A4', 'perkecil margin'. " +
      "Hanya set properti yang diminta; yang lain dibiarkan.",
    input_schema: {
      type: "object",
      properties: {
        orientation: {
          type: "string",
          enum: ["portrait", "landscape"],
          description: "Orientasi halaman.",
        },
        paperSize: {
          type: "string",
          enum: ["A4", "Letter", "Legal", "A3", "A5"],
          description: "Ukuran kertas.",
        },
        marginPreset: {
          type: "string",
          enum: ["normal", "narrow", "moderate", "wide"],
          description: "Preset margin: normal=2.54cm, narrow=1.27cm, moderate=2.54/1.91, wide=2.54/5.08.",
        },
        marginCm: {
          type: "object",
          description: "Margin manual dalam cm (override preset). Set field yang perlu saja.",
          properties: {
            top: { type: "number" }, bottom: { type: "number" },
            left: { type: "number" }, right: { type: "number" },
          },
        },
      },
    },
  };

  const format_paragraph = {
    name: "format_paragraph",
    description:
      "Atur format PARAGRAF pada range target: perataan (alignment), spasi sebelum/sesudah (pt), " +
      "jarak baris, dan indentasi kiri/baris-pertama (pt). Untuk format karakter (bold/warna) pakai format_text.",
    input_schema: withTarget({
      alignment: {
        type: "string",
        enum: ["Left", "Centered", "Right", "Justified"],
      },
      spaceBefore: { type: "number", description: "Spasi sebelum paragraf (pt)." },
      spaceAfter: { type: "number", description: "Spasi sesudah paragraf (pt)." },
      lineSpacing: { type: "number", description: "Jarak antar baris (pt)." },
      leftIndent: { type: "number", description: "Indentasi kiri (pt)." },
      firstLineIndent: { type: "number", description: "Indentasi baris pertama (pt)." },
    }),
  };

  const apply_style = {
    name: "apply_style",
    description:
      "Terapkan STYLE bawaan Word ke paragraf target (mis. ubah jadi 'Heading 1', 'Title', 'Normal'). " +
      "Gunakan untuk perintah seperti 'jadikan ini Heading 2' atau 'buat semua judul jadi Heading 1'.",
    input_schema: withTarget({
      styleName: {
        type: "string",
        description:
          "Nama style bawaan, mis. 'Heading 1','Heading 2','Title','Subtitle','Normal','Quote'.",
      },
    }, ["target", "styleName"]),
  };

  const insert_break = {
    name: "insert_break",
    description:
      "Sisipkan pemisah: ganti halaman (page break) atau ganti bagian (section break). " +
      "Lokasi mengikuti range target (default: sebelum range). Section break diperlukan bila " +
      "ingin orientasi/ukuran berbeda antar bagian.",
    input_schema: withTarget({
      breakType: {
        type: "string",
        enum: ["page", "sectionNext", "sectionContinuous"],
        default: "page",
        description: "page=ganti halaman; sectionNext=section baru di halaman berikut; sectionContinuous=section baru tanpa ganti halaman.",
      },
      position: {
        type: "string",
        enum: ["before", "after"],
        default: "before",
        description: "Sisipkan sebelum atau sesudah range target.",
      },
    }, ["target"]),
  };

  // ---- Tool tambahan (Fase 4 batch 2) ----

  const create_table = {
    name: "create_table",
    description:
      "Buat tabel dari data 2D, ATAU konversi teks yang sedang diseleksi menjadi tabel " +
      "(baris dipisah baris-baru, kolom dipisah tab/koma). Gunakan untuk 'ubah teks ini jadi tabel' " +
      "atau 'buatkan tabel ...'.",
    input_schema: {
      type: "object",
      properties: {
        data: {
          type: "array",
          description: "Data tabel sbg array baris; tiap baris array string sel.",
          items: { type: "array", items: { type: "string" } },
        },
        fromSelection: {
          type: "boolean",
          default: false,
          description: "true = konversi teks terseleksi jadi tabel (abaikan 'data').",
        },
        colDelimiter: { type: "string", default: "\t", description: "Pemisah kolom saat fromSelection (default tab; bisa ',')." },
        headerRow: { type: "boolean", default: true, description: "Tebalkan baris pertama sbg header." },
        style: { type: "string", description: "Nama style tabel bawaan, mis. 'Grid Table 4 - Accent 1'." },
      },
    },
  };

  const format_list = {
    name: "format_list",
    description:
      "Jadikan paragraf target sbg daftar berbutir (bullet) atau bernomor (numbered). " +
      "Gunakan untuk 'jadikan poin-poin' atau 'beri penomoran'.",
    input_schema: withTarget({
      listType: {
        type: "string",
        enum: ["bullet", "number"],
        default: "bullet",
        description: "bullet = berbutir; number = bernomor.",
      },
    }, ["target"]),
  };

  const manage_header_footer = {
    name: "manage_header_footer",
    description:
      "Atur isi header atau footer dokumen. Gunakan untuk 'tambahkan judul di header' atau " +
      "'tulis nama perusahaan di footer'. Untuk nomor halaman otomatis pakai set_page_numbers.",
    input_schema: {
      type: "object",
      properties: {
        area: { type: "string", enum: ["header", "footer"], description: "Bagian yang diatur." },
        text: { type: "string", description: "Teks yang ditulis (mengganti isi lama)." },
        alignment: { type: "string", enum: ["Left", "Centered", "Right"], default: "Left" },
      },
      required: ["area", "text"],
    },
  };

  const set_page_numbers = {
    name: "set_page_numbers",
    description:
      "Tambahkan nomor halaman OTOMATIS yang berjalan (1,2,3,…) di header (atas) atau footer " +
      "(bawah) setiap halaman. Gunakan untuk 'beri nomor halaman', 'nomor halaman di tengah atas', " +
      "'page number di kanan bawah', dll. Ini nomor field sungguhan, bukan teks statis.",
    input_schema: {
      type: "object",
      properties: {
        position: {
          type: "string",
          enum: ["top", "bottom"],
          default: "bottom",
          description: "top = header (atas halaman); bottom = footer (bawah halaman).",
        },
        alignment: { type: "string", enum: ["Left", "Centered", "Right"], default: "Centered",
          description: "Perataan: 'tengah atas' -> position=top, alignment=Centered." },
        format: {
          type: "string",
          enum: ["plain", "page_x_of_y"],
          default: "plain",
          description: "plain = '1'; page_x_of_y = '1 of 10'.",
        },
      },
    },
  };

  const insert_image = {
    name: "insert_image",
    description:
      "Sisipkan gambar dari data base64 ke lokasi target (default akhir dokumen). " +
      "Catatan: model biasanya tidak punya data base64; tool ini untuk dipakai alur internal/komposit.",
    input_schema: withTarget({
      base64: { type: "string", description: "Data gambar base64 (tanpa prefix data:)." },
      width: { type: "number", description: "Lebar opsional (pt)." },
    }, ["base64"]),
  };

  // Daftar final (urutan = urutan yang dikirim ke LLM).
  const SCHEMAS = [
    get_document_outline, format_text, replace_text,
    set_page_layout, format_paragraph, apply_style, insert_break,
    create_table, format_list, manage_header_footer, set_page_numbers, insert_image,
  ];

  return { SCHEMAS, targetSchema, byName: indexByName(SCHEMAS) };

  function indexByName(list) {
    const m = {};
    list.forEach((s) => (m[s.name] = s));
    return m;
  }
});
