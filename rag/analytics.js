// rag/analytics.js — manajemen log & analitik ambang batas generasi (R6).
// Berjalan in-memory untuk mencatat histori keputusan gate retrieval & hasil verifikasi.

const logs = [];

function logGeneration(event) {
  logs.push({
    timestamp: new Date().toISOString(),
    ...event
  });
  if (logs.length > 500) logs.shift(); // Batasi hanya menyimpan 500 log terakhir
}

function getStats() {
  const total = logs.length;
  const accepted = logs.filter((l) => l.accepted).length;
  const rejected = total - accepted;

  const rejectionReasons = {};
  logs.forEach((l) => {
    if (!l.accepted && l.reason) {
      rejectionReasons[l.reason] = (rejectionReasons[l.reason] || 0) + 1;
    }
  });

  const scores = logs.map((l) => l.maxScore).filter((s) => s != null);
  const avgScore = scores.length
    ? Number((scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(3))
    : 0;

  // Distribusi skor tertinggi
  const scoreDistribution = { "0.0-0.2": 0, "0.2-0.3": 0, "0.3-0.4": 0, "0.4-0.5": 0, "0.5+": 0 };
  scores.forEach((s) => {
    if (s < 0.2) scoreDistribution["0.0-0.2"]++;
    else if (s < 0.3) scoreDistribution["0.2-0.3"]++;
    else if (s < 0.4) scoreDistribution["0.3-0.4"]++;
    else if (s < 0.5) scoreDistribution["0.4-0.5"]++;
    else scoreDistribution["0.5+"]++;
  });

  return {
    total,
    accepted,
    rejected,
    avgScore,
    rejectionReasons,
    scoreDistribution,
    gateScoreMin: Number(process.env.GATE_SCORE_MIN || 0.3),
    recentLogs: logs.slice(-15).reverse() // 15 log terbaru
  };
}

module.exports = { logGeneration, getStats };
