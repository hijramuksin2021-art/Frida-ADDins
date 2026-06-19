// tools/selfcheck.js — uji invarian registry tanpa butuh Word.
// Jalankan: npm run check
// Memastikan "single source of truth" tetap konsisten:
//   1) Setiap schema punya handler bernama sama, dan sebaliknya.
//   2) Tiap schema berbentuk valid (name, description, input_schema.type === object).
//   3) Tiap handler adalah fungsi.
// Exit code != 0 bila ada pelanggaran (cocok untuk CI / pre-commit nanti).

const { SCHEMAS } = require("./schemas");
const { HANDLERS, resolveTarget, resolveHandler } = require("./handlers");
const { riskScore, needsConfirm, Permissions, makeAuditLog } = require("./safety");

const problems = [];
const ok = [];

function check(cond, msg, okMsg) {
  if (cond) ok.push(okMsg || msg);
  else problems.push(msg);
}

// 1) bentuk schema
for (const s of SCHEMAS) {
  check(typeof s.name === "string" && s.name.length > 0, `schema tanpa name: ${JSON.stringify(s).slice(0, 60)}`);
  check(typeof s.description === "string" && s.description.length > 10, `schema '${s.name}' deskripsi terlalu pendek`);
  check(s.input_schema && s.input_schema.type === "object", `schema '${s.name}' input_schema.type harus 'object'`);
}

// 2) parity nama: schema -> handler
const schemaNames = SCHEMAS.map((s) => s.name);
const handlerNames = Object.keys(HANDLERS);

for (const n of schemaNames) {
  check(typeof HANDLERS[n] === "function", `schema '${n}' TIDAK punya handler`, `handler ada untuk '${n}'`);
}
// 3) parity nama: handler -> schema
for (const n of handlerNames) {
  check(schemaNames.includes(n), `handler '${n}' TIDAK punya schema`, `schema ada untuk '${n}'`);
}

// 4) resolveTarget tersedia
check(typeof resolveTarget === "function", "resolveTarget hilang dari handlers");

// 5) resolveHandler tahan nama yang di-rename provider (regresi Fase 2)
check(typeof resolveHandler === "function", "resolveHandler hilang dari handlers");
if (typeof resolveHandler === "function") {
  const r1 = resolveHandler("CompatGetDocumentOutline375718");
  check(r1 && r1.name === "get_document_outline",
    "resolveHandler gagal memetakan nama ter-rename ke get_document_outline",
    "resolveHandler memetakan nama ter-rename");
  check(resolveHandler("totally_unknown_xyz") === null,
    "resolveHandler salah memetakan nama yang tak dikenal", "resolveHandler menolak nama tak dikenal");
}

// 6) safety: riskScore & permissions (regresi Fase 3)
check(typeof riskScore === "function", "riskScore hilang dari safety");
if (typeof riskScore === "function") {
  // read = 0 risiko
  check(riskScore({ name: "get_document_outline", input: {} }) === 0,
    "get_document_outline seharusnya risiko 0", "read tool risiko 0");
  // format pada seleksi = aman (tak perlu konfirmasi)
  check(!needsConfirm({ name: "format_text", input: { target: { mode: "selection" }, bold: true } }),
    "format_text pada selection seharusnya tidak perlu konfirmasi",
    "format selection tanpa konfirmasi");
  // format seluruh dokumen = berisiko (>=3)
  check(riskScore({ name: "format_text", input: { target: { mode: "whole_document" } } }) >= 3,
    "format_text whole_document seharusnya risiko >=3", "whole_document berisiko");
  // replace_text tanpa target (default seluruh dok) = perlu konfirmasi
  check(needsConfirm({ name: "replace_text", input: { find: "A", replace: "B" } }),
    "replace_text seluruh dokumen seharusnya perlu konfirmasi", "replace-all perlu konfirmasi");
  // set_page_layout = replace OOXML penuh -> berisiko
  check(needsConfirm({ name: "set_page_layout", input: { orientation: "landscape" } }),
    "set_page_layout seharusnya perlu konfirmasi", "set_page_layout berisiko");
  // matikan track changes -> berisiko
  check(needsConfirm({ name: "set_track_changes", input: { mode: "off" } }),
    "set_track_changes off seharusnya perlu konfirmasi", "track-off berisiko");
  // hapus baris tabel -> berisiko
  check(needsConfirm({ name: "edit_table", input: { deleteRowIndices: [2] } }),
    "edit_table delete row seharusnya perlu konfirmasi", "hapus baris berisiko");
  // edit sel tabel biasa -> aman (auto)
  check(!needsConfirm({ name: "edit_table", input: { cellEdits: [{ r: 0, c: 0, newText: "x" }] } }),
    "edit_table cellEdits seharusnya tidak perlu konfirmasi", "edit sel aman");
  // kebijakan deny
  Permissions.policy["__danger_test"] = "deny";
  check(Permissions.isBlocked("__danger_test"), "kebijakan deny tidak berlaku");
  delete Permissions.policy["__danger_test"];
}
// audit log mencatat
const a = makeAuditLog();
a.record({ name: "format_text", input: {} }, { applied: 3 }, false);
check(a.entries.length === 1 && a.entries[0].ok === true, "AuditLog gagal mencatat entri");

// laporan
console.log(`FRIDA tool registry selfcheck`);
console.log(`  tools terdaftar : ${schemaNames.length} (${schemaNames.join(", ")})`);
console.log(`  cek lulus       : ${ok.length}`);
if (problems.length) {
  console.error(`  MASALAH (${problems.length}):`);
  problems.forEach((p) => console.error("   - " + p));
  process.exit(1);
}
console.log("  semua invarian OK.");
