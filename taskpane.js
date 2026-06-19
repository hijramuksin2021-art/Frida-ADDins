/* FRIDA — penyunting cerdas di dalam Word.
   Alur: kumpulkan konteks (semua paragraf + seleksi + tabel terseleksi) -> kirim ke
   server lokal (/api/edit) bersama instruksi -> terima rencana perubahan ->
   terapkan ke Word: perbaiki (replace), sisip paragraf (insertAfter/append),
   atau rapikan sel tabel (tableOps). */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = () => process(true);
    document.getElementById("preview").onclick = () => process(false);
  } else {
    setStatus("FRIDA hanya untuk Microsoft Word.", "err");
  }
});

function setStatus(msg, cls) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status" + (cls ? " " + cls : "");
}

function busy(on) {
  document.getElementById("run").disabled = on;
  document.getElementById("preview").disabled = on;
}

// Ambil seluruh paragraf dokumen + teks/indeks seleksi + grid tabel (jika ada).
async function gatherContext(context) {
  const body = context.document.body;
  const allParas = body.paragraphs;
  allParas.load("items/text");

  const sel = context.document.getSelection();
  const selParas = sel.paragraphs;
  selParas.load("items/text");
  sel.load("text");

  // tabel yang menyentuh seleksi
  const selTables = sel.tables;
  selTables.load("items");

  await context.sync();

  const paragraphs = [];
  allParas.items.forEach((p, i) => {
    if (p.text && p.text.trim().length) paragraphs.push({ i, text: p.text });
  });

  // cocokkan paragraf terseleksi ke indeks dokumen (berdasarkan teks)
  const selIndices = [];
  const selTexts = selParas.items.map((p) => p.text);
  allParas.items.forEach((p, i) => {
    if (selTexts.includes(p.text) && p.text.trim().length) selIndices.push(i);
  });
  const selection = sel.text && sel.text.trim().length
    ? { text: sel.text, indices: selIndices }
    : null;

  // baca grid tabel pertama yang diseleksi
  let table = null;
  let tableObj = null;
  if (selTables.items.length > 0) {
    tableObj = selTables.items[0];
    tableObj.load("values");
    await context.sync();
    table = { rows: tableObj.values };
  }

  return { paragraphs, selection, table, allParas, tableObj, body };
}

async function process(apply) {
  const instruction = document.getElementById("instruction").value.trim();
  if (!instruction) { setStatus("Tulis dulu perintahnya untuk FRIDA.", "err"); return; }
  document.getElementById("result").innerHTML = "";
  busy(true);
  setStatus(apply ? "FRIDA membaca dokumen…" : "FRIDA menyiapkan tinjauan…");

  try {
    await Word.run(async (context) => {
      const ctx = await gatherContext(context);
      if (ctx.paragraphs.length === 0 && !ctx.table) {
        setStatus("Dokumen kosong.", "err"); return;
      }

      const resp = await fetch("/api/edit", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          instruction,
          paragraphs: ctx.paragraphs,
          selection: ctx.selection,
          table: ctx.table,
        }),
      });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error || ("HTTP " + resp.status));

      const pOps = Array.isArray(data.paragraphOps) ? data.paragraphOps : [];
      const tOps = Array.isArray(data.tableOps) ? data.tableOps : [];
      renderResult(data.summary, pOps, tOps, ctx.paragraphs);

      if (!apply) {
        setStatus((pOps.length + tOps.length) + " usulan perubahan (belum diterapkan).", "ok");
        return;
      }
      if (pOps.length === 0 && tOps.length === 0) {
        setStatus("Tidak ada yang perlu diubah.", "ok"); return;
      }

      let count = applyParagraphOps(ctx.allParas, ctx.body, pOps);
      count += applyTableOps(ctx.tableObj, tOps);
      await context.sync();
      setStatus("FRIDA menerapkan " + count + " perubahan ke dokumen.", "ok");
    });
  } catch (err) {
    setStatus("Gagal: " + (err.message || err), "err");
  } finally {
    busy(false);
  }
}

// Terapkan operasi paragraf. Penting: kerjakan dari indeks BESAR ke kecil
// agar penyisipan tidak menggeser indeks paragraf berikutnya.
function applyParagraphOps(allParas, body, ops) {
  let n = 0;
  const items = allParas.items;
  // kerjakan dari indeks besar ke kecil agar penyisipan tidak menggeser indeks lain
  const sorted = ops.slice().sort((a, b) => (b.i ?? 1e9) - (a.i ?? 1e9));
  sorted.forEach((op) => {
    if (typeof op.newText !== "string") return;
    const chunks = op.newText.split("\n").map((s) => s.trim()).filter((s) => s.length > 0);

    if (op.action === "append") {
      chunks.forEach((t) => body.insertParagraph(t, Word.InsertLocation.end));
      n++;
    } else if (op.action === "insertAfter") {
      const target = items[op.i];
      if (!target) return;
      let anchor = target;
      chunks.forEach((t) => { anchor = anchor.insertParagraph(t, Word.InsertLocation.after); });
      n++;
    } else { // replace
      const target = items[op.i];
      if (!target) return;
      target.insertText(op.newText, Word.InsertLocation.replace);
      n++;
    }
  });
  return n;
}

function applyTableOps(tableObj, ops) {
  if (!tableObj || ops.length === 0) return 0;
  let n = 0;
  ops.forEach((op) => {
    try {
      const cell = tableObj.getCell(op.r, op.c);
      cell.body.clear();
      cell.body.insertText(op.newText, Word.InsertLocation.start);
      n++;
    } catch (e) { /* sel di luar jangkauan: lewati */ }
  });
  return n;
}

function renderResult(summary, pOps, tOps, paragraphs) {
  const byIndex = {};
  paragraphs.forEach((p) => (byIndex[p.i] = p.text));
  const box = document.getElementById("result");
  let html = "";
  if (summary) html += '<div class="summary">' + escapeHtml(summary) + "</div>";

  pOps.forEach((op) => {
    const label =
      op.action === "append" ? "➕ Tambah di akhir" :
      op.action === "insertAfter" ? "➕ Sisip paragraf baru" : "✏️ Perbaiki";
    html += '<div class="edit">';
    html += '<div class="tag">' + label + (op.reason ? " — " + escapeHtml(op.reason) : "") + "</div>";
    if (op.action === "replace" && byIndex[op.i] != null) {
      html += '<div class="old">' + escapeHtml(byIndex[op.i]) + "</div>";
    }
    html += '<div class="new">' + escapeHtml(op.newText) + "</div>";
    html += "</div>";
  });

  if (tOps.length) {
    html += '<div class="edit"><div class="tag">📊 Rapikan tabel (' + tOps.length + " sel)</div>";
    tOps.forEach((op) => {
      html += '<div class="cell">[' + op.r + "," + op.c + "] → " + escapeHtml(op.newText) + "</div>";
    });
    html += "</div>";
  }
  box.innerHTML = html;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
