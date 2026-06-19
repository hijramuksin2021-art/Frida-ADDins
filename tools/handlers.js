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

  const HANDLERS = {
    get_document_outline,
    format_text,
    replace_text,
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
