import https from "https";
import { requireUser, sendAuthError } from "../_lib/auth.js";

const MAX_QUESTION = 4000;
const MAX_SYSTEM = 20000;
const MAX_CONTEXT = 800000;

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

  try {
    const result = await callClaude(apiKey, systemPrompt, contextData, question);
    res.status(200).json(result);
  } catch (err) {
    console.error("[chat] Claude API error:", err.message);
    res.status(502).json({ error: "The assistant could not answer right now. Try again shortly." });
  }
}

function callClaude(apiKey, systemPrompt, contextData, question) {
  const system = systemPrompt || "You are an OKR analytics assistant. Answer questions based on the provided data. Be concise and use specific numbers. Reply in the same language as the question.";
  const userMessage = contextData ? `${contextData}\n\nQuestion: ${question}` : question;

  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      model: "claude-opus-5-5",
      max_tokens: 2048,
      system,
      tools: TOOLS,
      messages: [{ role: "user", content: userMessage }],
    });

    const options = {
      hostname: "api.anthropic.com",
      path: "/v1/messages",
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "Content-Length": Buffer.byteLength(body),
      },
    };

    const req = https.request(options, r => {
      let data = "";
      r.on("data", chunk => { data += chunk; });
      r.on("end", () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.error) return reject(new Error(parsed.error.message));
          // Check for tool_use block first
          const toolUse = parsed.content?.find(b => b.type === "tool_use" && b.name === "propose_bulk_action");
          if (toolUse) {
            resolve({ action: toolUse.input });
          } else {
            const textBlock = parsed.content?.find(b => b.type === "text");
            resolve({ answer: textBlock?.text || "No response." });
          }
        } catch (e) { reject(e); }
      });
    });
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}
