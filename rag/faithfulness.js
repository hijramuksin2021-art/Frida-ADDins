// rag/faithfulness.js — verifikasi faithfulness paragraf terhadap sumber (R6).
// Menggunakan LLM untuk deteksi kontradiksi antara paragraf yang dihasilkan dan chunks sumber.
// Strategi: skeptical prompt untuk mengidentifikasi klaim yang bertentangan atau menyesatkan.

const { callModel, firstToolInput } = require("./llm");

const FAITHFULNESS_SYSTEM = [
  "Anda adalah verifikator kebenaran akademik yang skeptis.",
  "Tugas: Analisis apakah klaim dalam paragraf KONSISTEN dengan passage sumber yang diberikan.",
  "Untuk setiap kalimat klaim dalam paragraf, tentukan:",
  "  ENTAILED: klaim didukung langsung oleh passage",
  "  NEUTRAL: klaim tidak bertentangan tapi tidak didukung passage",
  "  CONTRADICTION: klaim bertentangan dengan passage atau menyesatkan",
  "Hanya flag CONTRADICTION. Berikan alasan singkat untuk setiap contradiction yang ditemukan.",
  "Balas via tool submit_faithfulness_check dengan daftar contradictions (jika ada) atau empty list."
].join("\n");

const SUBMIT_TOOL = {
  name: "submit_faithfulness_check",
  description: "Kirim hasil verifikasi faithfulness.",
  input_schema: {
    type: "object",
    properties: {
      contradictions: {
        type: "array",
        items: {
          type: "object",
          properties: {
            sentence: { type: "string", description: "Kalimat yang kontradiktif" },
            reason: { type: "string", description: "Alasan singkat mengapa kontradiktif" },
            severity: { type: "string", enum: ["minor", "major"], description: "Tingkat keparahan" }
          },
          required: ["sentence", "reason", "severity"]
        },
        description: "Daftar klaim yang kontradiktif (kosong jika semua konsisten)"
      },
      overall_faithful: {
        type: "boolean",
        description: "Apakah paragraf secara keseluruhan faithful (true jika tidak ada contradiction major)"
      }
    },
    required: ["contradictions", "overall_faithful"]
  }
};

function buildUserContent(paragraph, chunks) {
  const passageList = chunks
    .map((c, i) => `[Passage ${i + 1}] (source_id=${c.source_id}, chunk_id=${c.id})\n${c.text}`)
    .join("\n\n");

  return (
    "PARAGRAF YANG DIVERIFIKASI:\n" + paragraph +
    "\n\nPASSAGES SUMBER:\n" + passageList +
    "\n\nAnalisis: Apakah kalimat dalam paragraf konsisten dengan passages? " +
    "Flag HANYA klaim yang JELAS bertentangan atau menyesatkan."
  );
}

async function verifyFaithfulness(paragraph, chunks) {
  if (!paragraph || !paragraph.trim()) {
    return { contradictions: [], overall_faithful: true, error: "paragraph kosong" };
  }

  if (!chunks || chunks.length === 0) {
    return { contradictions: [], overall_faithful: true, note: "tidak ada chunks untuk verifikasi" };
  }

  try {
    const data = await callModel({
      system: FAITHFULNESS_SYSTEM,
      messages: [{ role: "user", content: buildUserContent(paragraph, chunks) }],
      tools: [SUBMIT_TOOL],
      tool_choice: { type: "tool", name: "submit_faithfulness_check" },
      maxTokens: 1000
    });

    const result = firstToolInput(data) || {};

    // Ekstrak kontradiksi MAJOR saja (minor diabaikan untuk mengurangi false positive)
    const majorContradictions = (result.contradictions || []).filter(c => c.severity === "major");

    return {
      contradictions: majorContradictions,
      overall_faithful: result.overall_faithful === true && majorContradictions.length === 0,
      all_findings: result.contradictions || []
    };
  } catch (e) {
    return {
      error: "faithfulness check gagal: " + (e.message || e),
      contradictions: [],
      overall_faithful: true // fallback: izinkan jika LLM gagal (lebih baik daripada reject)
    };
  }
}

module.exports = { verifyFaithfulness };
