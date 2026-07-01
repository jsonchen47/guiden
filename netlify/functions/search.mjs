import { readFile } from "node:fs/promises";
import { join } from "node:path";

const model = process.env.GUIDEN_LLM_MODEL || "gemini-3.5-flash";

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

let storyCache;

function json(statusCode, payload) {
  return {
    statusCode,
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Headers": "Content-Type",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  };
}

async function loadStories() {
  if (storyCache) return storyCache;

  const source = await readFile(join(process.cwd(), "src/data/stories.ts"), "utf8");
  const assignment = source.indexOf("= [");
  const start = assignment === -1 ? -1 : source.indexOf("[", assignment);
  const end = source.lastIndexOf("];");
  if (start === -1 || end === -1) {
    throw new Error("Could not find story data in src/data/stories.ts.");
  }

  storyCache = JSON.parse(source.slice(start, end + 1));
  return storyCache;
}

function tokenize(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter((term) => term.length > 2);
}

function selectRelevantStories(stories, query, storyIds = []) {
  const requested = new Set(Array.isArray(storyIds) ? storyIds : []);
  const terms = new Set(tokenize(query));
  return stories
    .map((story) => {
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
    })
    .sort((a, b) => b.score - a.score)
    .slice(0, 14)
    .map(({ story }) => ({
      id: story.id,
      title: story.title,
      categories: story.categories,
      tags: story.tags,
      whatHelped: story.whatHelped,
      excerpt: story.excerpt,
      story: String(story.story).slice(0, 550)
    }));
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
        "\n\nReturn JSON exactly like this: {\"summary\":\"one sentence recommending the strongest evidence-backed next idea\", \"patterns\":[{\"label\":\"reach out to friends more often\", \"description\":\"why this concrete action may help, based on these stories\", \"count\":3, \"percent\":21, \"examples\":[\"support system\"], \"evidence\":\"short paraphrase\"}]}. Return 3 to 5 recommendations, ranked by how many provided stories support them AND how relevant they are to the user's exact query. Labels must be concrete actions that complete the sentence 'you could try to ...'. Avoid vague labels like support, mindset, routine, or treatment. Avoid recommendations that introduce unrelated issues. For example, do not recommend sobriety, medication changes, trauma work, or diagnosis-specific treatment unless the user mentioned that topic or the matching stories overwhelmingly and directly support it; if included, make the label conditional, like 'if substance use is part of this, reduce alcohol or cannabis'. Count only the records below. sampleSize=" +
        selectedStories.length +
        "\n\nStories:\n" +
        JSON.stringify(prompt)
    }
  ];
}

function extractGeminiOutputText(payload) {
  const chunks = [];
  for (const candidate of payload.candidates || []) {
    for (const part of candidate.content?.parts || []) {
      if (typeof part.text === "string") chunks.push(part.text);
    }
  }
  return chunks.join("\n");
}

function normalizePatterns(patterns) {
  if (!Array.isArray(patterns)) return [];

  return patterns
    .filter((pattern) => pattern && typeof pattern === "object" && typeof pattern.label === "string")
    .map((pattern) => ({
      label: pattern.label.trim(),
      description:
        typeof pattern.description === "string" && pattern.description.trim()
          ? pattern.description.trim()
          : `Similar stories mentioned ${pattern.label} as part of what helped.`,
      count: Number.isFinite(Number(pattern.count)) ? Number(pattern.count) : 1,
      percent: Number.isFinite(Number(pattern.percent)) ? Number(pattern.percent) : 0,
      examples: Array.isArray(pattern.examples) ? pattern.examples.map(String).slice(0, 3) : [],
      evidence: typeof pattern.evidence === "string" ? pattern.evidence : ""
    }))
    .filter((pattern) => pattern.label)
    .slice(0, 5);
}

function parseModelJson(text) {
  const trimmed = String(text || "").trim();
  const jsonStart = trimmed.indexOf("{");
  const jsonEnd = trimmed.lastIndexOf("}");
  if (jsonStart === -1 || jsonEnd === -1) {
    throw new Error(`The Gemini response did not include JSON. Preview: ${trimmed.slice(0, 180) || "[empty]"}`);
  }

  const parsed = JSON.parse(trimmed.slice(jsonStart, jsonEnd + 1));
  return {
    summary: String(parsed.summary || ""),
    patterns: normalizePatterns(parsed.patterns)
  };
}

async function callGemini({ query, selectedStories }) {
  if (!process.env.GEMINI_API_KEY) {
    throw new Error("GEMINI_API_KEY is missing from Netlify environment variables.");
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
      headers: { "Content-Type": "application/json" },
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

  const parsed = parseModelJson(extractGeminiOutputText(payload));
  return {
    ...parsed,
    sampleSize: selectedStories.length,
    model,
    provider: "gemini"
  };
}

export async function handler(event) {
  if (event.httpMethod === "OPTIONS") {
    return json(200, {});
  }

  if (event.httpMethod !== "POST") {
    return json(404, { error: "Use POST /.netlify/functions/search", provider: "gemini", model });
  }

  try {
    const body = JSON.parse(event.body || "{}");
    const query = String(body.query || "").trim();
    if (!query) {
      return json(400, { error: "query is required", provider: "gemini", model });
    }

    const stories = await loadStories();
    const selectedStories = selectRelevantStories(stories, query, body.storyIds);
    const result = await callGemini({ query, selectedStories });
    return json(200, result);
  } catch (error) {
    return json(500, {
      error: error instanceof Error ? error.message : "Unknown server error",
      provider: "gemini",
      model
    });
  }
}
