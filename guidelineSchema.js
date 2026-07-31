// guidelineSchema.js — skema validasi untuk file panduan penulisan (guidelines)

function validateGuideline(data) {
  const errors = [];

  if (!data.id || typeof data.id !== "string") errors.push("Field 'id' wajib berupa string.");
  if (!data.nama || typeof data.nama !== "string") errors.push("Field 'nama' wajib berupa string.");
  if (!data.type || typeof data.type !== "string") errors.push("Field 'type' wajib berupa string.");
  
  if (data.formatting) {
    if (typeof data.formatting !== "object") errors.push("Field 'formatting' harus berupa objek.");
    // validasi field formatting
    if (data.formatting.font && typeof data.formatting.font !== "string") errors.push("formatting.font harus berupa string.");
    if (data.formatting.fontSize && typeof data.formatting.fontSize !== "number") errors.push("formatting.fontSize harus berupa number.");
  }

  if (data.structure) {
    if (!Array.isArray(data.structure)) errors.push("Field 'structure' harus berupa array dari nama-nama bab.");
  }

  if (data.citationStyle) {
    if (typeof data.citationStyle !== "string") errors.push("Field 'citationStyle' harus berupa string.");
  }

  if (data.rules) {
    if (!Array.isArray(data.rules)) {
      errors.push("Field 'rules' harus berupa array of strings.");
    } else {
      if (data.rules.some(r => typeof r !== "string")) {
        errors.push("Isi dari 'rules' harus berupa string.");
      }
    }
  }

  return {
    valid: errors.length === 0,
    errors
  };
}

module.exports = { validateGuideline };
