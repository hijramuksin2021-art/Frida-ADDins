// rag/guidelineConfig.js — manajemen konfigurasi panduan penulisan (R7).
// Menyimpan id panduan penulisan aktif ke file guideline.local.json secara lokal.

const fs = require("fs");
const path = require("path");

const FILE = path.join(__dirname, "..", "guideline.local.json");
const GUIDELINES_DIR = path.join(__dirname, "guidelines");

let activeId = ""; // Kosong berarti "Tidak ada / Generik"

function loadPersisted() {
  try {
    const data = JSON.parse(fs.readFileSync(FILE, "utf8"));
    return data.activeId || "";
  } catch (_) {
    return "";
  }
}

function init() {
  activeId = loadPersisted();
}

function getActiveId() {
  if (activeId === undefined) init();
  return activeId;
}

function setActiveId(id) {
  activeId = String(id || "").trim();
  persist();
  return status();
}

function persist() {
  try {
    fs.writeFileSync(FILE, JSON.stringify({ activeId }, null, 2));
  } catch (_) {}
}

// Memuat data lengkap guideline aktif dari file JSON
function getActiveGuideline() {
  const id = getActiveId();
  if (!id) return null;
  try {
    const p = path.join(GUIDELINES_DIR, `${id}.json`);
    if (fs.existsSync(p)) {
      return JSON.parse(fs.readFileSync(p, "utf8"));
    }
  } catch (_) {}
  return null;
}

function status() {
  const id = getActiveId();
  const guideline = getActiveGuideline();
  return {
    activeId: id,
    activeName: guideline ? guideline.nama : "Tidak ada / Generik",
    activeFakultas: guideline ? guideline.fakultas : null,
    activeUniversitas: guideline ? guideline.universitas : null,
    gayaSitasi: guideline && guideline.sitasi ? guideline.sitasi.gaya : null,
    hasActive: !!id
  };
}

module.exports = { getActiveId, setActiveId, getActiveGuideline, status, init };
