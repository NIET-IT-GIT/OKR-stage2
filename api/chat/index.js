import Anthropic from "@anthropic-ai/sdk";
import { requireUser, sendAuthError } from "../_lib/auth.js";

const MAX_QUESTION = 4000;
const MAX_SYSTEM = 20000;
const MAX_CONTEXT = 800000;
const MODEL = "claude-opus-5-5";
const MAX_TOKENS = 8000;
const DEFAULT_SYSTEM = "You are an OKR analytics assistant. Answer questions based on the provided data. Be concise and use specific numbers. Reply in the same language as the question.";

// Always added after the (admin-editable) system prompt, so charts work even with a custom prompt
const CHART_GUIDE = `Charts: when a trend over time, a comparison across entities or a ranking is easier to see than to read, you may add up to 2 charts to your answer. Skip charts for simple facts, single numbers or text-only answers.
Write each chart as its own fenced block, on its own lines, exactly like this:
\`\`\`chart
{"type":"line","title":"Cash collected by RTO","unit":"$","labels":["Jul 2026","Aug 2026"],"series":[{"name":"NIET","values":[412300,398100]}]}
\`\`\`
- type: "line" for trends over time, "bar" to compare a few entities (optionally several series side by side), "hbar" for rankings or long names.
- unit: "$" for money, "%" for percentages, "" for counts or ratings.
- labels: up to 24. series: 1 to 6, each with exactly one value per label (use null when a value is missing).
- Values are plain numbers in full units (412300, not "412k" or "$412,300"; 66.7 for 66.7%).
- Use only numbers that appear in the data. Never add, average or estimate a value for a chart; each value is checked against the data and unmatched values are flagged to the reader.
- Still state the key numbers in the text; the chart supports the answer, it does not replace it.`;

const TOOLS = [{
  name: "propose_bulk_action",
  description: "Call this tool when the user wants to approve or reject OKR submissions. Never perform or describe the approval in text — always use this tool so the admin can review a full submission summary before confirming. Use this for any approve/reject/bulk-approve/bulk-reject request.",
  input_schema: {
    type: "object",
    properties: {
      type: {
        type: "string",
        enum: ["approve", "reject"],
        description: "The action type"
      },
      filters: {
        type: "object",
        description: "Criteria to select which pending submissions to act on. Omit a field to not filter on it.",
        properties: {
          deptName: { type: "string", description: "Filter by department name (partial match ok). Omit for all departments." },
          memberName: { type: "string", description: "Filter by member name (partial match ok). Omit for all members." },
          periodKey: { type: "string", description: "Filter by period key prefix, e.g. '2026-08' for August 2026. Omit to default to current calendar month." },
          allPeriods: { type: "boolean", description: "If true, include submissions from all time periods, not just current month." }
        }
      },
      message: {
        type: "string",
        description: "Brief message shown as the card header, e.g. 'Pending submissions for IT Department — August 2026'"
      }
    },
    required: ["type", "message"]
  }
}];

const ALLOWED_ORIGINS = ["https://okr.nietgroup.com.au", "http://localhost:5173", "http://localhost:4173"];

export default async function handler(req, res) {
  const origin = req.headers.origin || "";
  const allowedOrigin = ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
  res.setHeader("Access-Control-Allow-Origin", allowedOrigin);
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Vary", "Origin");
  if (req.method === "OPTIONS") return res.status(204).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try { await requireUser(req); } catch (err) { return sendAuthError(res, err); }

  let parsed = req.body;
  if (typeof parsed === "string") { try { parsed = JSON.parse(parsed); } catch { parsed = {}; } }
  if (!parsed || typeof parsed !== "object") parsed = {};
  const { question, systemPrompt, contextData } = parsed;

  if (!question || typeof question !== "string") return res.status(400).json({ error: "missing question" });
  if (question.length > MAX_QUESTION) return res.status(413).json({ error: "Your question is too long. Shorten it and try again." });
  if ((systemPrompt && String(systemPrompt).length > MAX_SYSTEM) || (contextData && String(contextData).length > MAX_CONTEXT)) {
    return res.status(413).json({ error: "Too much data was sent with this question." });
  }

  const apiKey = process.env.CLAUDE_API_KEY;
  if (!apiKey) { console.error("[chat] CLAUDE_API_KEY not configured"); return res.status(500).json({ error: "The assistant is not configured." }); }

  // One attempt within the function's 60-second limit, so a slow answer ends with a clear message, not a platform timeout
  const client = new Anthropic({ apiKey, timeout: 55000, maxRetries: 0 });
  try {
    const response = await client.messages.create({
      model: MODEL,
      // Opus 5.5 always thinks first and the thinking counts towards this limit, so leave room for long answers
      max_tokens: MAX_TOKENS,
      output_config: { effort: "medium" },
      system: `${systemPrompt || DEFAULT_SYSTEM}\n\n${CHART_GUIDE}`,
      tools: TOOLS,
      messages: [{ role: "user", content: contextData ? `${contextData}\n\nQuestion: ${question}` : question }],
    });
    return res.status(200).json(toResult(response, contextData));
  } catch (err) {
    if (err instanceof Anthropic.APIError) console.error("[chat] API error", err.status, err.message);
    else console.error("[chat] error", err.message);
    const timedOut = err instanceof Anthropic.APIConnectionTimeoutError;
    return res.status(timedOut ? 504 : 502).json({ error: timedOut
      ? "The assistant took too long to answer. Ask about one part at a time and try again."
      : "The assistant could not answer right now. Try again shortly." });
  }
}

export const CUT_OFF_NOTE = "*The answer was cut off because it was too long. Ask about one part at a time for the full detail.*";

// Turns a Messages API response into what the Pilot shows: an approval card, or the full answer text
export function toResult(response, contextData = "") {
  if (response.stop_reason === "refusal") {
    console.warn("[chat] refusal:", response.stop_details?.category);
    return { answer: "The assistant can't help with this request." };
  }
  const toolUse = response.content.find(b => b.type === "tool_use" && b.name === "propose_bulk_action");
  // A tool call cut off by the length limit has incomplete input, so it is never turned into an approval card
  if (toolUse && response.stop_reason === "tool_use") return { action: toolUse.input };
  const text = checkCharts(response.content.filter(b => b.type === "text").map(b => b.text).join("\n\n").trim(), contextData);
  if (response.stop_reason === "max_tokens") {
    console.warn("[chat] answer hit max_tokens");
    return { answer: text ? `${text}\n\n${CUT_OFF_NOTE}` : CUT_OFF_NOTE };
  }
  return { answer: text || "No response." };
}

// Every number written in the data the assistant was given (full units; "412k" / "1.2m" also as 412000 / 1200000)
function dataNumbers(contextData) {
  const nums = [];
  const re = /(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)\s*([kKmM])?(?![\w])/g;
  let m;
  while ((m = re.exec(String(contextData || "")))) {
    const n = parseFloat(m[1].replace(/,/g, ""));
    nums.push(n);
    if (m[2]) nums.push(n * (/k/i.test(m[2]) ? 1e3 : 1e6));
  }
  return Float64Array.from(nums).sort();
}

// True when the data contains a number within rounding distance of v (whole units, or 0.5% for large values)
function inData(sorted, v) {
  const a = Math.abs(v), tol = Math.max(0.5, a * 0.005);
  let lo = 0, hi = sorted.length;
  while (lo < hi) { const mid = (lo + hi) >> 1; if (sorted[mid] < a - tol) lo = mid + 1; else hi = mid; }
  return lo < sorted.length && sorted[lo] <= a + tol;
}

// Adds a "checked" summary to each chart block: how many values there are and which aren't in the data.
// The browser shows unmatched values under the chart. Blocks that aren't valid JSON are left for the browser to reject.
export function checkCharts(text, contextData) {
  if (!text.includes("```chart")) return text;
  let sorted = null;
  return text.replace(/```chart[^\S\n]*\n([\s\S]*?)\n?```/g, (block, body) => {
    let chart;
    try { chart = JSON.parse(body); } catch { return block; }
    if (!chart || typeof chart !== "object" || !Array.isArray(chart.series)) return block;
    sorted = sorted || dataNumbers(contextData);
    const values = chart.series.flatMap(s => (Array.isArray(s?.values) ? s.values : [])).filter(v => typeof v === "number" && Number.isFinite(v));
    const unmatched = values.filter(v => !inData(sorted, v));
    chart.checked = { total: values.length, unmatched: [...new Set(unmatched)].slice(0, 12) , unmatchedCount: unmatched.length };
    return "```chart\n" + JSON.stringify(chart) + "\n```";
  });
}
