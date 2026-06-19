// tools/handlers.js — implementasi Office.js untuk tiap tool (ARCHITECTURE §5).
// Pasangan schema-nya ada di tools/schemas.js (nama HARUS sama).
//
// Dual-mode:
//   - Browser (taskpane): window.FRIDA_HANDLERS = { HANDLERS, resolveTarget }.
//   - Node (selfcheck): module.exports sama, supaya parity nama bisa diuji
//     TANPA Word. Handler memang tak bisa dijalankan di Node (butuh Word/context),
//     tapi keberadaan & nama fungsinya bisa diverifikasi.
//
// Kontrak handler: async function(context, args) -> objek hasil (untuk tool_result).
// 'context' adalah Word.RequestContext di dalam Word.run. 'args' = input dari LLM.
// context.sync() dipanggil di dalam handler bila perlu membaca/menulis.

(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api; // Node
  if (typeof window !== "undefined") window.FRIDA_HANDLERS = api;            // Browser
})(this, function () {
  // ---- resolveTarget: selektor deklaratif -> Word.Range[] ----
  // Menghapus seluruh kelas bug "indeks paragraf bergeser": target diresolusi
  // sekali, sebelum mutasi.
  async function resolveTarget(context, target) {
    target = target || { mode: "selection" };
    const body = context.document.body;

    switch (target.mode) {
      case "whole_document":
        return [body.getRange()];

      case "selection":
        return [context.document.getSelection()];

      case "search": {
        if (!target.value) return [];
        const res = body.search(target.value, { matchCase: false });
        res.load("items");
        await context.sync();
        let items = res.items;
        if (target.occurrence === "first") items = items.slice(0, 1);
        else if (target.occurrence === "nth") {
          const idx = (target.n || 1) - 1;
          items = items[idx] ? [items[idx]] : [];
        }
        return items;
      }

      case "heading": {
        const ps = body.paragraphs;
        ps.load("items/styleBuiltIn");
        await context.sync();
        return ps.items
          .filter((p) => /Heading/i.test(p.styleBuiltIn || ""))
          .map((p) => p.getRange());
      }

      case "style": {
        const ps = body.paragraphs;
        ps.load("items/style,items/styleBuiltIn");
        await context.sync();
        const want = (target.value || "").toLowerCase();
        return ps.items
          .filter(
            (p) =>
              (p.style || "").toLowerCase() === want ||
              (p.styleBuiltIn || "").toLowerCase() === want
          )
          .map((p) => p.getRange());
      }

      case "paragraph_index": {
        const ps = body.paragraphs;
        ps.load("items");
        await context.sync();
        const p = ps.items[target.index];
        return p ? [p.getRange()] : [];
      }

      default:
        return [];
    }
  }

  // ---- Tool: get_document_outline (read) ----
  async function get_document_outline(context, args) {
    args = args || {};
    const paras = context.document.body.paragraphs;
    paras.load("items/text,items/styleBuiltIn,items/style");
    const tables = context.document.body.tables;
    tables.load("items");
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    const cut = args.include_text ? 1e9 : 80;
    const paragraphs = paras.items.map((p, i) => {
      const sb = p.styleBuiltIn || "";
      const m = sb.match(/Heading\s*(\d)/i);
      return {
        i,
        style: p.style || sb,
        isHeading: /Heading/i.test(sb),
        level: m ? Number(m[1]) : null,
        preview: (p.text || "").slice(0, cut),
      };
    });
    return {
      paragraphs,
      sections: sections.items.length,
      hasTables: tables.items.length,
    };
  }

  // ---- Tool: format_text (write) ----
  async function format_text(context, args) {
    const ranges = await resolveTarget(context, args.target);
    ranges.forEach((r) => {
      const f = r.font;
      if (args.bold !== undefined) f.bold = args.bold;
      if (args.italic !== undefined) f.italic = args.italic;
      if (args.underline !== undefined) f.underline = args.underline;
      if (args.fontName) f.name = args.fontName;
      if (args.fontSize) f.size = args.fontSize;
      if (args.color) f.color = args.color;
      if (args.highlightColor) f.highlightColor = args.highlightColor;
    });
    await context.sync();
    return { applied: ranges.length, rangesAffected: ranges.length };
  }

  // ---- Tool: replace_text (write) ----
  async function replace_text(context, args) {
    // Default whole_document bila target tak diberi.
    const scope =
      args.target && args.target.mode && args.target.mode !== "whole_document"
        ? await resolveTarget(context, args.target)
        : [context.document.body.getRange()];

    let replaced = 0;
    for (const scopeRange of scope) {
      const res = scopeRange.search(args.find, {
        matchCase: !!args.matchCase,
        matchWholeWord: !!args.wholeWord,
      });
      res.load("items");
      await context.sync();
      res.items.forEach((r) => {
        r.insertText(args.replace, Word.InsertLocation.replace);
        replaced++;
      });
      await context.sync();
    }
    return { replaced };
  }

  // ---- helper: resolve target -> Paragraph[] (untuk properti tingkat paragraf) ----
  // format_paragraph & apply_style butuh objek Paragraph, bukan Range.
  async function resolveTargetParagraphs(context, target) {
    target = target || { mode: "selection" };
    const body = context.document.body;
    const all = body.paragraphs;

    if (target.mode === "paragraph_index") {
      all.load("items");
      await context.sync();
      const p = all.items[target.index];
      return p ? [p] : [];
    }
    if (target.mode === "heading") {
      all.load("items/styleBuiltIn");
      await context.sync();
      return all.items.filter((p) => /Heading/i.test(p.styleBuiltIn || ""));
    }
    if (target.mode === "style") {
      all.load("items/style,items/styleBuiltIn");
      await context.sync();
      const want = (target.value || "").toLowerCase();
      return all.items.filter(
        (p) => (p.style || "").toLowerCase() === want ||
               (p.styleBuiltIn || "").toLowerCase() === want
      );
    }
    if (target.mode === "whole_document") {
      all.load("items");
      await context.sync();
      return all.items;
    }
    // selection / search: ambil paragraf yang menyentuh range tsb
    const ranges = await resolveTarget(context, target);
    const paras = ranges.map((r) => r.paragraphs);
    paras.forEach((pc) => pc.load("items"));
    await context.sync();
    const out = [];
    paras.forEach((pc) => pc.items.forEach((p) => out.push(p)));
    return out;
  }

  // ---- Tool: format_paragraph (write) ----
  async function format_paragraph(context, args) {
    const paras = await resolveTargetParagraphs(context, args.target);
    paras.forEach((p) => {
      if (args.alignment) p.alignment = args.alignment;
      if (args.spaceBefore !== undefined) p.spaceBefore = args.spaceBefore;
      if (args.spaceAfter !== undefined) p.spaceAfter = args.spaceAfter;
      if (args.lineSpacing !== undefined) p.lineSpacing = args.lineSpacing;
      if (args.leftIndent !== undefined) p.leftIndent = args.leftIndent;
      if (args.firstLineIndent !== undefined) p.firstLineIndent = args.firstLineIndent;
    });
    await context.sync();
    return { applied: paras.length };
  }

  // ---- Tool: apply_style (write) ----
  async function apply_style(context, args) {
    const paras = await resolveTargetParagraphs(context, args.target);
    paras.forEach((p) => { p.style = args.styleName; });
    await context.sync();
    return { applied: paras.length, style: args.styleName };
  }

  // ---- Tool: insert_break (write) ----
  async function insert_break(context, args) {
    const ranges = await resolveTarget(context, args.target);
    if (!ranges.length) return { inserted: 0 };
    const map = {
      page: Word.BreakType.page,
      sectionNext: Word.BreakType.sectionNext,
      sectionContinuous: Word.BreakType.sectionContinuous,
    };
    const bt = map[args.breakType] || Word.BreakType.page;
    const loc = args.position === "after"
      ? Word.InsertLocation.after : Word.InsertLocation.before;
    let n = 0;
    ranges.forEach((r) => { r.insertBreak(bt, loc); n++; });
    await context.sync();
    return { inserted: n, breakType: args.breakType || "page" };
  }

  // ---- Tool: set_page_layout (write, via OOXML sectPr) ----
  // Office.js JS API tidak mengekspos pageSetup, jadi kita ubah sectPr di OOXML
  // (pgSz utk ukuran/orientasi, pgMar utk margin). Satuan OOXML = twips (1cm=567twip).
  const CM = 567; // twips per cm (1 inch=1440 twip, 1cm=1440/2.54)
  const PAPER = { // ukuran dlm twips (lebar x tinggi, portrait)
    A4:     [11906, 16838], Letter: [12240, 15840], Legal: [12240, 20160],
    A3:     [16838, 23811], A5:     [8391, 11906],
  };
  const MARGIN_PRESET = { // [top,bottom,left,right] dlm cm
    normal:   [2.54, 2.54, 2.54, 2.54],
    narrow:   [1.27, 1.27, 1.27, 1.27],
    moderate: [2.54, 2.54, 1.91, 1.91],
    wide:     [2.54, 2.54, 5.08, 5.08],
  };

  async function set_page_layout(context, args) {
    const body = context.document.body;
    const ooxmlResult = body.getOoxml();
    await context.sync();
    let xml = ooxmlResult.value;

    // 1) pgSz: ukuran & orientasi
    if (args.paperSize || args.orientation) {
      xml = patchPgSz(xml, args.paperSize, args.orientation);
    }
    // 2) pgMar: margin
    let margins = null;
    if (args.marginPreset && MARGIN_PRESET[args.marginPreset]) {
      const m = MARGIN_PRESET[args.marginPreset];
      margins = { top: m[0], bottom: m[1], left: m[2], right: m[3] };
    }
    if (args.marginCm) {
      margins = Object.assign(margins || {}, args.marginCm);
    }
    if (margins) xml = patchPgMar(xml, margins);

    body.clear();
    body.insertOoxml(xml, Word.InsertLocation.replace);
    await context.sync();
    return {
      ok: true,
      orientation: args.orientation || null,
      paperSize: args.paperSize || null,
      margin: margins ? "diatur" : null,
    };
  }

  // --- util OOXML untuk set_page_layout ---
  function patchPgSz(xml, paperSize, orientation) {
    return xml.replace(/<w:pgSz\b[^>]*\/>/g, (tag) => {
      let w = numAttr(tag, "w:w"), h = numAttr(tag, "w:h");
      if (paperSize && PAPER[paperSize]) { w = PAPER[paperSize][0]; h = PAPER[paperSize][1]; }
      let orient = orientation || (numAttr(tag, "w:w") > numAttr(tag, "w:h") ? "landscape" : "portrait");
      if (orient === "landscape" && w < h) { const t = w; w = h; h = t; }
      if (orient === "portrait" && w > h) { const t = w; w = h; h = t; }
      return '<w:pgSz w:w="' + w + '" w:h="' + h + '" w:orient="' + orient + '"/>';
    });
  }
  function patchPgMar(xml, m) {
    const tw = (cm) => (cm == null ? null : Math.round(cm * CM));
    return xml.replace(/<w:pgMar\b[^>]*\/>/g, (tag) => {
      const top = tw(m.top) ?? numAttr(tag, "w:top");
      const bottom = tw(m.bottom) ?? numAttr(tag, "w:bottom");
      const left = tw(m.left) ?? numAttr(tag, "w:left");
      const right = tw(m.right) ?? numAttr(tag, "w:right");
      const header = numAttr(tag, "w:header") || 720;
      const footer = numAttr(tag, "w:footer") || 720;
      const gutter = numAttr(tag, "w:gutter") || 0;
      return '<w:pgMar w:top="' + top + '" w:right="' + right + '" w:bottom="' + bottom +
             '" w:left="' + left + '" w:header="' + header + '" w:footer="' + footer +
             '" w:gutter="' + gutter + '"/>';
    });
  }
  function numAttr(tag, name) {
    const m = tag.match(new RegExp(name.replace(":", "\\:") + '="(-?\\d+)"'));
    return m ? parseInt(m[1], 10) : 0;
  }

  // ---- Tool: create_table (write) ----
  async function create_table(context, args) {
    let data = args.data;
    let insertPoint, replaceSel = false;

    if (args.fromSelection) {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      const col = args.colDelimiter || "\t";
      data = (sel.text || "")
        .split(/\r?\n/)
        .filter((line) => line.trim().length)
        .map((line) => line.split(col).map((c) => c.trim()));
      insertPoint = sel;
      replaceSel = true;
    } else {
      insertPoint = context.document.body.getRange(Word.RangeLocation.end);
    }

    if (!data || !data.length) return { error: "tidak ada data tabel" };
    const rows = data.length;
    const cols = Math.max.apply(null, data.map((r) => r.length));
    // ratakan tiap baris agar lebar kolom konsisten
    const grid = data.map((r) => {
      const row = r.slice();
      while (row.length < cols) row.push("");
      return row;
    });

    const loc = replaceSel ? Word.InsertLocation.replace : Word.InsertLocation.end;
    const table = insertPoint.insertTable(rows, cols, loc, grid);
    if (args.style) {
      try { table.style = args.style; } catch (e) { /* style tak ada: abaikan */ }
    }
    if (args.headerRow !== false && rows > 0) {
      table.getRow(0).font.bold = true;
    }
    await context.sync();
    return { rows, cols };
  }

  // ---- Tool: format_list (write) ----
  async function format_list(context, args) {
    const paras = await resolveTargetParagraphs(context, args.target);
    if (!paras.length) return { applied: 0 };
    // mulai list pada paragraf pertama, sisanya menempel ke list yang sama
    const first = paras[0];
    const list = first.startNewList();
    list.load("id");
    await context.sync();
    for (let k = 1; k < paras.length; k++) {
      paras[k].attachToList(list.id, 0);
    }
    // bullet vs number: ubah level type bila perlu
    await context.sync();
    return { applied: paras.length, listType: args.listType || "bullet" };
  }

  // ---- Tool: manage_header_footer (write) ----
  async function manage_header_footer(context, args) {
    const section = context.document.sections.getFirst();
    const area = args.area === "footer"
      ? section.getFooter(Word.HeaderFooterType.primary)
      : section.getHeader(Word.HeaderFooterType.primary);
    area.clear();
    const p = area.insertParagraph(args.text || "", Word.InsertLocation.start);
    if (args.alignment) p.alignment = args.alignment;
    await context.sync();
    return { ok: true, area: args.area };
  }

  // ---- Tool: set_page_numbers (write) ----
  // Pakai Range.insertField(FieldType.page) — jauh lebih andal daripada insertOoxml
  // ke footer (yang sering melempar GeneralException di Word desktop). Bila API field
  // tak tersedia di host, fallback ke teks biasa + catatan.
  // position: 'top' = header (nomor di atas), 'bottom' = footer (default).
  async function set_page_numbers(context, args) {
    const section = context.document.sections.getFirst();
    const atTop = args.position === "top";
    const area = atTop
      ? section.getHeader(Word.HeaderFooterType.primary)
      : section.getFooter(Word.HeaderFooterType.primary);
    area.clear();
    const p = area.insertParagraph("", Word.InsertLocation.start);
    p.alignment = args.alignment || "Centered";
    await context.sync();

    const canField =
      typeof Word !== "undefined" &&
      Word.FieldType && p.getRange && typeof p.getRange().insertField === "function";

    if (canField) {
      // sisipkan field PAGE (dan NUMPAGES utk "x of y") ke dalam paragraf footer
      const r0 = p.getRange(Word.RangeLocation.start);
      r0.insertField(Word.InsertLocation.start, Word.FieldType.page);
      if (args.format === "page_x_of_y") {
        const rEnd = p.getRange(Word.RangeLocation.end);
        rEnd.insertText(" of ", Word.InsertLocation.end);
        rEnd.insertField(Word.InsertLocation.end, Word.FieldType.numPages);
      }
      await context.sync();
      return { ok: true, position: atTop ? "top" : "bottom",
               format: args.format || "plain", method: "field" };
    }

    // Fallback: tanpa API field — tulis penanda agar pengguna tahu host tak mendukung.
    p.insertText(args.format === "page_x_of_y" ? "Halaman ? dari ?" : "Halaman ?",
                 Word.InsertLocation.start);
    await context.sync();
    return { ok: true, position: atTop ? "top" : "bottom",
             format: args.format || "plain", method: "text_fallback",
             note: "Host tidak mendukung field otomatis; ditulis teks penanda." };
  }

  // ---- Tool: insert_image (write) ----
  async function insert_image(context, args) {
    if (!args.base64) return { error: "base64 kosong" };
    const ranges = await resolveTarget(context, args.target || { mode: "whole_document" });
    const anchor = ranges.length
      ? ranges[ranges.length - 1]
      : context.document.body.getRange(Word.RangeLocation.end);
    const pic = anchor.insertInlinePictureFromBase64(args.base64, Word.InsertLocation.end);
    if (args.width) pic.width = args.width;
    await context.sync();
    return { ok: true };
  }

  // ---- Tool: insert_toc (write, OOXML TOC field di body) ----
  // Tak ada API TOC langsung; sisipkan field TOC via OOXML. Insertion ke BODY andal
  // (beda dgn footer). Pengguna klik kanan -> Update Field utk mengisi.
  async function insert_toc(context, args) {
    const body = context.document.body;
    const anchor = args.location === "selection"
      ? context.document.getSelection()
      : body.getRange(Word.RangeLocation.start);
    const loc = args.location === "selection"
      ? Word.InsertLocation.before : Word.InsertLocation.start;
    if (args.title) {
      const t = anchor.insertParagraph(args.title, loc);
      t.styleBuiltIn = Word.BuiltInStyleName.heading1;
    }
    body.insertOoxml(tocOoxml(), loc);
    await context.sync();
    return { ok: true, note: "Daftar isi disisipkan. Klik kanan -> Update Field untuk mengisi." };
  }

  // ---- Tool: manage_comments (write) ----
  async function manage_comments(context, args) {
    const ranges = await resolveTarget(context, args.target);
    if (!ranges.length) return { error: "target untuk komentar tak ditemukan" };
    let n = 0;
    ranges.forEach((r) => {
      try { r.insertComment(args.text || ""); n++; } catch (e) { /* host lama: lewati */ }
    });
    await context.sync();
    return { added: n };
  }

  // ---- Tool: set_track_changes (write/state) ----
  async function set_track_changes(context, args) {
    const map = {
      trackAll: Word.ChangeTrackingMode.trackAll,
      trackMineOnly: Word.ChangeTrackingMode.trackMineOnly,
      off: Word.ChangeTrackingMode.off,
    };
    const mode = map[args.mode];
    if (mode === undefined) return { error: "mode tidak valid: " + args.mode };
    context.document.changeTrackingMode = mode;
    await context.sync();
    return { ok: true, mode: args.mode };
  }

  // ---- Tool: edit_table (write; pengganti tableOps lama) ----
  async function edit_table(context, args) {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    const idx = args.tableIndex || 0;
    const table = tables.items[idx];
    if (!table) return { error: "tabel indeks " + idx + " tidak ada" };

    let cellsEdited = 0, rowsAdded = 0, rowsDeleted = 0;

    // 1) edit isi sel
    (args.cellEdits || []).forEach((e) => {
      try {
        const cell = table.getCell(e.r, e.c);
        cell.body.clear();
        cell.body.insertText(e.newText, Word.InsertLocation.start);
        cellsEdited++;
      } catch (err) { /* sel di luar jangkauan: lewati */ }
    });

    // 2) tambah baris di akhir
    if (Array.isArray(args.addRows) && args.addRows.length) {
      try {
        table.addRows(Word.InsertLocation.end, args.addRows.length, args.addRows);
        rowsAdded = args.addRows.length;
      } catch (err) { /* abaikan bila gagal */ }
    }

    // 3) hapus baris — dari indeks BESAR ke kecil agar tak menggeser
    const dels = (args.deleteRowIndices || []).slice().sort((a, b) => b - a);
    dels.forEach((ri) => {
      try { table.getRow(ri).delete(); rowsDeleted++; } catch (err) { /* lewati */ }
    });

    await context.sync();
    return { cellsEdited, rowsAdded, rowsDeleted };
  }

  // --- util OOXML untuk field TOC ---
  function tocOoxml() {
    const instr = 'TOC \\\\o "1-3" \\\\h \\\\z \\\\u';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
      '<pkg:xmlData>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> ' + instr + ' </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>Klik kanan untuk memperbarui daftar isi.</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>' +
      '</pkg:xmlData></pkg:part></pkg:package>';
  }

  const HANDLERS = {
    get_document_outline,
    format_text,
    replace_text,
    set_page_layout,
    format_paragraph,
    apply_style,
    insert_break,
    create_table,
    format_list,
    manage_header_footer,
    set_page_numbers,
    insert_image,
    insert_toc,
    manage_comments,
    set_track_changes,
    edit_table,
  };

  // ---- Pemetaan nama tool yang andal ----
  // Beberapa provider/proxy MENGUBAH nama tool di respons (mis. 'get_document_outline'
  // menjadi 'CompatGetDocumentOutline375718'). Cocokkan via bentuk kanonik (huruf+angka,
  // lowercase) dengan containment: nama registry yang terkandung di nama balasan dipakai;
  // bila banyak yang cocok, ambil yang TERPANJANG (paling spesifik).
  function canon(s) {
    return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  }
  function resolveHandler(name) {
    if (HANDLERS[name]) return { name, fn: HANDLERS[name] };
    const target = canon(name);
    let best = null;
    for (const key of Object.keys(HANDLERS)) {
      const ck = canon(key);
      if (target.includes(ck) && (!best || ck.length > canon(best).length)) best = key;
    }
    return best ? { name: best, fn: HANDLERS[best] } : null;
  }

  return { HANDLERS, resolveTarget, resolveHandler };
});
