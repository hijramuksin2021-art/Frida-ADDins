// rag/llm.js — pemanggil model untuk tugas RAG di server (generate_paragraph_from_source, dll).
// Kini hanya wrapper tipis di atas rag/aiProvider.js supaya generasi RAG ikut multi-provider:
// ganti provider/model lewat panel Pengaturan langsung berlaku untuk generasi RAG juga.

const aiProvider = require("./aiProvider");

// callModel({ system, messages, tools, tool_choice, maxTokens }) -> { content, stop_reason }
// (format internal Anthropic). Retry ditangani di aiProvider (kecuali error 4xx key/model).
async function callModel(opts) {
  return aiProvider.callMessagesRetry({
    system: opts.system,
    messages: opts.messages,
    tools: opts.tools,
    tool_choice: opts.tool_choice,
    maxTokens: opts.maxTokens,
  });
}

// Ambil input dari blok tool_use pertama (abaikan nama; provider kadang me-rename).
function firstToolInput(data) {
  const b = (data.content || []).find((x) => x.type === "tool_use" && x.input);
  return b ? b.input : null;
}

module.exports = { callModel, firstToolInput };
