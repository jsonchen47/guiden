import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(__dirname, "..");
const port = Number(process.env.GUIDEN_LLM_PORT || 8787);
const provider = process.env.GUIDEN_LLM_PROVIDER || "ollama";
const model =
  process.env.GUIDEN_LLM_MODEL ||
  process.env.OLLAMA_MODEL ||
  (provider === "gemini" ? "gemini-3.5-flash" : "llama3.2:3b");
const ollamaUrl = process.env.OLLAMA_URL || "http://127.0.0.1:11434";

const patternSchema = {
  type: "OBJECT",
  properties: {
    summary: { type: "STRING" },
    patterns: {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          label: { type: "STRING" },
          description: { type: "STRING" },
          count: { type: "INTEGER" },
          percent: { type: "INTEGER" },
          examples: {
            type: "ARRAY",
            items: { type: "STRING" }
          },
          evidence: { type: "STRING" }
        },
        required: ["label", "description", "count", "percent", "examples", "evidence"]
      }
    }
  },
  required: ["summary", "patterns"]
};

function sendJson(res, status, payload) {
  res.writeHead(status, {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Content-Type": "application/json"
  });
  res.end(JSON.stringify(payload));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => {
      body += chunk;
      if (body.length > 1_000_000) {
        reject(new Error("Request body is too large."));
        req.destroy();
      }
    });
    req.on("end", () => resolve(body));
    req.on("error", reject);
  });
}

async function loadStories() {
  const source = await readFile(join(projectRoot, "src/data/stories.ts"), "utf8");
  const assignment = source.indexOf("= [");
  const start = assignment === -1 ? -1 : source.indexOf("[", assignment);
  const end = source.lastIndexOf("];");
  if (start === -1 || end === -1) {
    throw new Error("Could not find story data in src/data/stories.ts.");
  }
  return JSON.parse(source.slice(start, end + 1));
}

const stories = await loadStories();

function tokenize(value) {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter((term) => term.length > 2);
}

function selectRelevantStories(query, storyIds = []) {
  const requested = new Set(Array.isArray(storyIds) ? storyIds : []);
  const terms = new Set(tokenize(query));
  const scored = stories.map((story) => {
    const text = [
      story.title,
      story.category,
      story.categories?.join(" "),
      story.tags?.join(" "),
      story.whatHelped?.join(" "),
      story.excerpt,
      story.story
    ].join(" ");
    const searchable = new Set(tokenize(text));
    const termScore = [...terms].filter((term) => searchable.has(term)).length;
    const requestedBoost = requested.has(story.id) ? 15 : 0;
    return { story, score: termScore + requestedBoost };
  });

  return scored
    .sort((a, b) => b.score - a.score)
    .slice(0, 8)
    .map(({ story }) => ({
      id: story.id,
      title: story.title,
      categories: story.categories,
      tags: story.tags,
      whatHelped: story.whatHelped,
      excerpt: story.excerpt,
      story: String(story.story).slice(0, 320)
    }));
}

function extractOpenAIOutputText(payload) {
  if (typeof payload.output_text === "string") return payload.output_text;
  const chunks = [];
  for (const item of payload.output || []) {
    for (const content of item.content || []) {
      if (typeof content.text === "string") chunks.push(content.text);
    }
  }
  return chunks.join("\n");
}

function extractGeminiOutputText(payload) {
  if (typeof payload.output_text === "string") return payload.output_text;
  const chunks = [];

  function visit(value) {
    if (!value) return;
    if (typeof value === "string") {
      chunks.push(value);
      return;
    }
    if (Array.isArray(value)) {
      value.forEach(visit);
      return;
    }
    if (typeof value === "object") {
      for (const nested of Object.values(value)) {
        visit(nested);
      }
    }
  }

  for (const item of payload.output || []) {
    if (typeof item.text === "string") chunks.push(item.text);
    for (const content of item.content || []) {
      if (typeof content.text === "string") chunks.push(content.text);
    }
  }
  for (const candidate of payload.candidates || []) {
    for (const part of candidate.content?.parts || []) {
      if (typeof part.text === "string") chunks.push(part.text);
    }
  }
  visit(payload.response);
  visit(payload.outputText);
  visit(payload.text);
  return chunks.join("\n");
}

function findStructuredPatternPayload(value) {
  if (!value) return null;
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findStructuredPatternPayload(item);
      if (found) return found;
    }
    return null;
  }
  if (typeof value !== "object") return null;
  if (Array.isArray(value.patterns) || typeof value.summary === "string") {
    return value;
  }
  for (const nested of Object.values(value)) {
    const found = findStructuredPatternPayload(nested);
    if (found) return found;
  }
  return null;
}

function buildAnalysisMessages({ query, selectedStories }) {
  const prompt = {
    query,
    sampleSize: selectedStories.length,
    stories: selectedStories
  };

  return [
    {
      role: "system",
      content:
        "Return only JSON. You analyze lived-experience story records. Find concrete repeated behaviors that helped people in similar situations. Turn the strongest repeated patterns into cautious suggestions. Do not give medical advice. Count only the provided records. Do not assume the user has conditions, behaviors, substance use, trauma, medication needs, or diagnoses they did not mention."
    },
    {
      role: "user",
      content:
        "User query: " +
        query +
        "\n\nReturn JSON exactly like this: {\"summary\":\"one sentence recommending the strongest evidence-backed next idea\", \"patterns\":[{\"label\":\"reach out to friends more often\", \"description\":\"why this concrete action may help, based on these stories\", \"count\":3, \"percent\":21, \"examples\":[\"support system\"], \"evidence\":\"short paraphrase\"}]}. Return 3 to 5 concrete recommendations ranked by evidence count and relevance to the query. Avoid unrelated issues; do not mention sobriety, medication changes, trauma work, or diagnosis-specific treatment unless the user mentioned it or the stories overwhelmingly support it. Count only the records below. sampleSize=" +
        selectedStories.length +
        "\n\nStories:\n" +
        JSON.stringify(prompt)
    }
  ];
}

function parseModelJson(text) {
  if (text && typeof text === "object") {
    return {
      summary: String(text.summary || ""),
      patterns: normalizePatterns(text.patterns)
    };
  }

  const trimmed = String(text || "").trim();
  const jsonStart = trimmed.indexOf("{");
  const jsonEnd = trimmed.lastIndexOf("}");
  if (jsonStart === -1 || jsonEnd === -1) {
    throw new Error(`The LLM response did not include JSON. Preview: ${trimmed.slice(0, 240) || "[empty]"}`);
  }

  const parsed = JSON.parse(trimmed.slice(jsonStart, jsonEnd + 1));
  return {
    summary: String(parsed.summary || ""),
    patterns: normalizePatterns(parsed.patterns)
  };
}

function normalizePatterns(patterns) {
  if (!Array.isArray(patterns)) return [];

  const flattened = patterns.flatMap((pattern) => {
    if (pattern && Array.isArray(pattern.patterns)) {
      return pattern.patterns;
    }
    return pattern;
  });

  return flattened
    .filter((pattern) => pattern && typeof pattern === "object" && typeof pattern.label === "string")
    .map((pattern) => {
      const label = pattern.label.trim();
      const description =
        typeof pattern.description === "string" &&
        pattern.description.trim() &&
        pattern.description.trim().toLowerCase() !== "concrete explanation"
          ? pattern.description.trim()
          : `Similar stories mentioned ${label} as part of what helped.`;
      return {
        label,
        description,
        count: Number.isFinite(Number(pattern.count)) ? Number(pattern.count) : 1,
        percent: Number.isFinite(Number(pattern.percent)) ? Number(pattern.percent) : 0,
        examples: Array.isArray(pattern.examples) ? pattern.examples.map(String).slice(0, 3) : [],
        evidence: typeof pattern.evidence === "string" && pattern.evidence !== "short paraphrase" ? pattern.evidence : ""
      };
    })
    .filter((pattern) => pattern.label.trim())
    .slice(0, 5);
}

async function callOllama({ query, selectedStories }) {
  const response = await fetch(`${ollamaUrl}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model,
      messages: buildAnalysisMessages({ query, selectedStories }),
      stream: false,
      format: "json",
      options: {
        temperature: 0.2,
        num_ctx: 8192
      }
    })
  });

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error || `Ollama request failed with ${response.status}`);
  }

  const parsed = parseModelJson(payload.message?.content || payload.response || "");
  return {
    ...parsed,
    sampleSize: selectedStories.length,
    model,
    provider: "ollama"
  };
}

async function callOpenAI({ query, selectedStories }) {
  if (!process.env.OPENAI_API_KEY) {
    throw new Error("OPENAI_API_KEY is missing. Set GUIDEN_LLM_PROVIDER=ollama for the free local model, or provide OPENAI_API_KEY.");
  }

  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model,
      temperature: 0.2,
      input: buildAnalysisMessages({ query, selectedStories })
    })
  });

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error?.message || `OpenAI request failed with ${response.status}`);
  }

  const parsed = parseModelJson(extractOpenAIOutputText(payload));
  return {
    ...parsed,
    sampleSize: selectedStories.length,
    model,
    provider: "openai"
  };
}

async function callGemini({ query, selectedStories }) {
  if (!process.env.GEMINI_API_KEY) {
    throw new Error("GEMINI_API_KEY is missing. Create a key in Google AI Studio, then run with GUIDEN_LLM_PROVIDER=gemini.");
  }

  const messages = buildAnalysisMessages({ query, selectedStories });
  const systemInstruction = messages.find((message) => message.role === "system")?.content || "";
  const input = messages
    .filter((message) => message.role !== "system")
    .map((message) => message.content)
    .join("\n\n");
  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(process.env.GEMINI_API_KEY)}`,
    {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      systemInstruction: {
        parts: [{ text: systemInstruction }]
      },
      contents: [
        {
          role: "user",
          parts: [{ text: input }]
        }
      ],
      generationConfig: {
        temperature: 0.2,
        response_mime_type: "application/json",
        response_schema: patternSchema
      }
    })
  }
  );

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error?.message || `Gemini request failed with ${response.status}`);
  }

  const structuredPayload = findStructuredPatternPayload(payload);
  const parsed = parseModelJson(structuredPayload || extractGeminiOutputText(payload));
  return {
    ...parsed,
    sampleSize: selectedStories.length,
    model,
    provider: "gemini"
  };
}

async function callModel({ query, selectedStories }) {
  if (provider === "openai") {
    return callOpenAI({ query, selectedStories });
  }
  if (provider === "gemini") {
    return callGemini({ query, selectedStories });
  }
  return callOllama({ query, selectedStories });
}

createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    sendJson(res, 200, {});
    return;
  }

  if (req.method === "GET" && req.url === "/api/health") {
    sendJson(res, 200, { ok: true, provider, model });
    return;
  }

  if (req.method !== "POST" || req.url !== "/api/search") {
    sendJson(res, 404, { error: "Use POST /api/search", provider, model });
    return;
  }

  try {
    const body = JSON.parse(await readBody(req));
    const query = String(body.query || "").trim();
    if (!query) {
      sendJson(res, 400, { error: "query is required" });
      return;
    }

    const selectedStories = selectRelevantStories(query, body.storyIds);
    const result = await callModel({ query, selectedStories });
    sendJson(res, 200, result);
  } catch (error) {
    sendJson(res, 500, {
      error: error instanceof Error ? error.message : "Unknown server error",
      provider,
      model
    });
  }
}).listen(port, "127.0.0.1", () => {
  console.log(`Guiden LLM search server listening on http://127.0.0.1:${port}`);
  console.log(`Using provider: ${provider}`);
  console.log(`Using model: ${model}`);
});
