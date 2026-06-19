// tools/selfcheck.js — uji invarian registry tanpa butuh Word.
// Jalankan: npm run check
// Memastikan "single source of truth" tetap konsisten:
//   1) Setiap schema punya handler bernama sama, dan sebaliknya.
//   2) Tiap schema berbentuk valid (name, description, input_schema.type === object).
//   3) Tiap handler adalah fungsi.
// Exit code != 0 bila ada pelanggaran (cocok untuk CI / pre-commit nanti).

const { SCHEMAS } = require("./schemas");
const { HANDLERS, resolveTarget } = require("./handlers");

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
