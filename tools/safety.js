// tools/safety.js — lapisan keselamatan (ARCHITECTURE §7). Dual-mode Node/browser.
//
// Berisi:
//   - riskScore(toolUse)         : skor risiko aksi (PURE — diuji di Node).
//   - Permissions                : kebijakan per-tool (PURE).
//   - makeAuditLog()             : pencatat aksi append-only (PURE).
//   - TransactionManager         : snapshot OOXML + rollback (butuh Word; no-op di Node).
//
// Bagian PURE sengaja dipisah agar bisa diuji tanpa Word (lihat selfcheck.js).

(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api; // Node
  if (typeof window !== "undefined") window.FRIDA_SAFETY = api;              // Browser
})(this, function () {
  // Nama tool yang hanya MEMBACA (tak pernah memicu konfirmasi/rollback).
  const READ_TOOLS = new Set(["get_document_outline", "search_text"]);

  // ---- riskScore: seberapa berbahaya satu tool_use ----
  // >=3 => wajib konfirmasi eksplisit pengguna. Skala selaras ARCHITECTURE §6.
  function riskScore(toolUse) {
    const name = (toolUse && toolUse.name) || "";
    const a = (toolUse && toolUse.input) || {};
    const targetStr = JSON.stringify(a.target || {});
    let s = 0;

    if (READ_TOOLS.has(name)) return 0; // read selalu aman

    // cari-ganti default = seluruh dokumen (target boleh kosong/tanpa mode)
    const impliesWholeDoc =
      /whole_document/.test(targetStr) ||
      (name.indexOf("replace_text") >= 0 && (!a.target || !a.target.mode));
    if (impliesWholeDoc) s += 3;                            // menyasar seluruh dokumen
    if (name.indexOf("replace_text") >= 0 && impliesWholeDoc) s += 1; // ganti masal = ekstra hati2
    if (/delete|clear|remove/i.test(name)) s += 4;         // destruktif
    if (name.indexOf("set_track_changes") >= 0 && a.mode === "off") s += 3;
    if (a.target && a.target.occurrence === "all" && a.target.mode === "search") s += 1;

    return s;
  }

  // Apakah satu tool_use butuh konfirmasi eksplisit?
  function needsConfirm(toolUse) {
    if (READ_TOOLS.has((toolUse && toolUse.name) || "")) return false;
    const rule = Permissions.ruleFor(toolUse.name);
    if (rule === "deny") return true;     // ditangani sbg blokir, tetap minta perhatian
    if (rule === "allow") return false;   // dipercaya tanpa konfirmasi
    if (rule === "confirm") return true;
    // default ("auto"): tergantung skor risiko
    return riskScore(toolUse) >= 3;
  }

  // ---- Permissions: kebijakan per-tool (bisa di-override enterprise) ----
  const Permissions = {
    // "allow" | "confirm" | "deny" | "auto"(default: berdasar riskScore)
    policy: {
      get_document_outline: "allow",
      search_text: "allow",
      set_track_changes: "confirm",
      "*": "auto",
    },
    ruleFor(name) {
      // cocokkan nama persis dulu, lalu pola "prefix_*", lalu wildcard.
      if (this.policy[name]) return this.policy[name];
      for (const key of Object.keys(this.policy)) {
        if (key.endsWith("*") && name.indexOf(key.slice(0, -1)) === 0) return this.policy[key];
      }
      return this.policy["*"] || "auto";
    },
    isBlocked(name) { return this.ruleFor(name) === "deny"; },
  };

  // ---- AuditLog: catatan append-only ----
  function makeAuditLog() {
    const entries = [];
    return {
      entries,
      record(toolUse, output, isError) {
        entries.push({
          ts: new Date().toISOString(),
          tool: (toolUse && toolUse.name) || "?",
          input: redact(toolUse && toolUse.input),
          ok: !isError,
          result: summarize(output),
        });
        return entries[entries.length - 1];
      },
      clear() { entries.length = 0; },
    };
  }

  function redact(input) {
    if (!input || typeof input !== "object") return input;
    // saat ini tak ada field rahasia di input tool; hook utk masa depan.
    return input;
  }
  function summarize(out) {
    if (out == null) return "";
    if (out.error) return "error: " + String(out.error).slice(0, 120);
    if (out.replaced != null) return out.replaced + " diganti";
    if (out.applied != null) return out.applied + " range";
    if (out.paragraphs != null) return out.paragraphs.length + " paragraf dibaca";
    return "ok";
  }

  // ---- TransactionManager: snapshot OOXML + rollback (butuh Word) ----
  // begin() dipanggil di dalam Word.run SEBELUM batch write. Bila batch gagal /
  // pengguna menekan Undo, rollback() mengembalikan body ke snapshot.
  const TransactionManager = {
    async begin(context) {
      if (typeof Word === "undefined") {
        // lingkungan Node (uji): kembalikan transaksi no-op.
        return { snapshot: null, async rollback() {} };
      }
      const ooxml = context.document.body.getOoxml();
      await context.sync();
      const snapshot = ooxml.value;
      return {
        snapshot,
        async rollback() {
          context.document.body.clear();
          context.document.body.insertOoxml(snapshot, Word.InsertLocation.replace);
          await context.sync();
        },
      };
    },

    // Restore snapshot mentah dalam Word.run terpisah (dipakai tombol Undo FRIDA).
    async restore(snapshot) {
      if (typeof Word === "undefined" || snapshot == null) return false;
      await Word.run(async (context) => {
        context.document.body.clear();
        context.document.body.insertOoxml(snapshot, Word.InsertLocation.replace);
        await context.sync();
      });
      return true;
    },
  };

  return { riskScore, needsConfirm, Permissions, makeAuditLog, TransactionManager, READ_TOOLS };
});
