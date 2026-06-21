// tools/selfcheck.js — uji invarian registry tanpa butuh Word.
// Jalankan: npm run check
// Memastikan "single source of truth" tetap konsisten:
//   1) Setiap schema punya handler bernama sama, dan sebaliknya.
//   2) Tiap schema berbentuk valid (name, description, input_schema.type === object).
//   3) Tiap handler adalah fungsi.
// Exit code != 0 bila ada pelanggaran (cocok untuk CI / pre-commit nanti).

const { SCHEMAS } = require("./schemas");
const { HANDLERS, resolveTarget, resolveHandler, previewTool } = require("./handlers");
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

// 2) parity nama: schema CLIENT -> handler (tool server/RAG dieksekusi di server, tak punya handler klien)
const clientSchemas = SCHEMAS.filter((s) => (s.runtime || "client") === "client");
const schemaNames = clientSchemas.map((s) => s.name);
const handlerNames = Object.keys(HANDLERS);

for (const n of schemaNames) {
  check(typeof HANDLERS[n] === "function", `schema '${n}' TIDAK punya handler`, `handler ada untuk '${n}'`);
}
// 3) parity nama: handler -> schema
for (const n of handlerNames) {
  check(schemaNames.includes(n), `handler '${n}' TIDAK punya schema`, `schema ada untuk '${n}'`);
}

// 4) resolveTarget + previewTool tersedia
check(typeof resolveTarget === "function", "resolveTarget hilang dari handlers");
check(typeof previewTool === "function", "previewTool hilang dari handlers (Fase 5)");

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
  // composite -> berisiko (perlu konfirmasi)
  check(needsConfirm({ name: "format_business_proposal", input: {} }),
    "format_business_proposal seharusnya perlu konfirmasi", "format_business_proposal berisiko");
  check(needsConfirm({ name: "insert_cover_page", input: { title: "X" } }),
    "insert_cover_page seharusnya perlu konfirmasi", "insert_cover_page berisiko");
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

// 7) R3: CSL citation engine — unit-test deterministik (tanpa butuh Word/server)
const csl = require("../rag/csl");
const store = require("../rag/store");
const cite = require("../rag/cite");

const r3meta = {
  type: "article-journal",
  title: "Selfcheck R3 Citation",
  author: [{ family: "Doe", given: "J." }, { family: "Smith", given: "A." }],
  issued: { year: 2024 },
  container: "Test Journal",
  volume: "1", issue: "2", page: "10-20",
  DOI: "10.1234/test.2024",
};
// normStyle
check(csl.normStyle("apa7") === "APA7", "normStyle 'apa7' → APA7 gagal", "normStyle APA7 OK");
check(csl.normStyle("IEEE") === "IEEE", "normStyle 'IEEE' gagal", "normStyle IEEE OK");
// inText — semua 5 gaya
const it_apa = csl.inText(r3meta, "APA7", {});
check(it_apa === "(Doe & Smith, 2024)", "inText APA7 format salah: " + it_apa, "inText APA7 OK");
const it_mla = csl.inText(r3meta, "MLA", {});
check(it_mla.includes("Doe"), "inText MLA harus mengandung 'Doe'", "inText MLA OK");
const it_chi = csl.inText(r3meta, "Chicago", {});
check(it_chi.includes("2024"), "inText Chicago harus mengandung tahun", "inText Chicago OK");
const it_har = csl.inText(r3meta, "Harvard", { page: "p.5" });
check(it_har.includes("p. 5") || it_har.includes("p.5"), "inText Harvard+page format salah: " + it_har, "inText Harvard+page OK");
const it_ieee = csl.inText(r3meta, "IEEE", { number: 7 });
check(it_ieee === "[7]", "inText IEEE format salah: " + it_ieee, "inText IEEE OK");
// bibEntry
const bib_apa = csl.bibEntry(r3meta, "APA7");
check(bib_apa.includes("doi.org"), "bibEntry APA7 harus mengandung DOI URL", "bibEntry APA7 DOI OK");
check(bib_apa.includes("(2024)"), "bibEntry APA7 harus mengandung tahun", "bibEntry APA7 tahun OK");

// store + cite pipeline (tulis & hapus doc tes)
const r3doc = store.save({
  filename: "selfcheck-r3.pdf", ext: "pdf", mime: "application/pdf",
  hash: "selfcheck-r3-" + Date.now(),
  title: r3meta.title, year: 2024, confidence: "user", csl: r3meta, chars: 100,
});
const r3id = r3doc.id;
const citeIT = cite.inTextFor(r3id, "APA7", {});
check(citeIT === "(Doe & Smith, 2024)", "cite.inTextFor APA7 gagal: " + citeIT, "cite.inTextFor APA7 OK");
const citeEntry = cite.entryFor(r3id, "MLA");
check(typeof citeEntry === "string" && citeEntry.length > 5, "cite.entryFor MLA kosong", "cite.entryFor MLA OK");
const citeBibs = cite.bibliography([r3id], "APA7");
check(citeBibs.length === 1 && citeBibs[0].source_id === r3id, "cite.bibliography gagal", "cite.bibliography OK");

// updateMetadata lalu verifikasi confidence = 'user'
store.updateMetadata(r3id, { title: "R3 Updated" });
const r3updated = store.get(r3id);
check(r3updated && r3updated.confidence === "user", "updateMetadata confidence harus 'user'", "updateMetadata confidence OK");
check(r3updated && r3updated.title === "R3 Updated", "updateMetadata title tidak tersinkron", "updateMetadata title OK");

// cleanup
store.remove(r3id);
check(store.get(r3id) === null, "remove gagal: dokumen masih ada", "store.remove OK");

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
