# Guiden

Guiden is a cross-platform iOS and Android prototype for sharing and discovering mental-health recovery testimonies.

The current website prototype includes:

- A Figma-inspired Home page with featured stories and topic browsing.
- A searchable testimony library seeded from the Reddit prototype workbook.
- An AI Search page that ranks stories by themes, tags, tone, and helpful actions.
- Story detail pages with background, what helped, timeline, tags, and related stories.
- An anonymous story submission form that stores submissions in local app state.
- Safety framing around peer support, moderation, and crisis support.

## Run locally

Install dependencies, then start Expo:

```sh
npm install
npm run start
```

From Expo you can open the app on iOS Simulator, Android Emulator, Expo Go, or web.

## LLM search

The search UI calls a local API server. Keep API keys in this server terminal, not in the browser code.

### Gemini setup

Gemini is the easiest hosted/free-tier option for this prototype.

1. Go to [Google AI Studio](https://aistudio.google.com/app/apikey) and create an API key.
2. In one terminal, run:

```sh
GUIDEN_LLM_PROVIDER=gemini GEMINI_API_KEY=your_key_here npm run llm-search
```

The default Gemini model is `gemini-3.5-flash`. To choose another model:

```sh
GUIDEN_LLM_PROVIDER=gemini GUIDEN_LLM_MODEL=gemini-3.1-flash-lite GEMINI_API_KEY=your_key_here npm run llm-search
```

In another terminal:

```sh
npm run web
```

The app posts searches to `http://127.0.0.1:8787/api/search`. If the server is not running, the UI shows a labeled local fallback summary instead of pretending an LLM answered.

### Ollama setup

Ollama is the free local option, but it is usually slower than Gemini.

First install Ollama and pull the model:

```sh
ollama pull llama3.2:3b
```

In one terminal:

```sh
npm run llm-search
```

In another terminal:

```sh
npm run web
```

Optional: to use OpenAI instead of Ollama, run:

```sh
GUIDEN_LLM_PROVIDER=openai GUIDEN_LLM_MODEL=gpt-4.1-mini OPENAI_API_KEY=your_key_here npm run llm-search
```

## Static web preview

Build the web bundle:

```sh
EXPO_NO_TELEMETRY=1 CI=1 npx expo export --platform web
```

Serve the exported site:

```sh
python3 -m http.server 8082 --bind 127.0.0.1 --directory dist
```

## Team demo with Gemini

For the easiest public demo with real Gemini responses, deploy to Netlify.

1. Push this project to a GitHub repository.
2. In Netlify, choose **Add new site** -> **Import an existing project**.
3. Pick the GitHub repository.
4. Netlify will read `netlify.toml`:
   - Build command: `EXPO_NO_TELEMETRY=1 CI=1 npx expo export --platform web --clear`
   - Publish directory: `dist`
   - Function directory: `netlify/functions`
5. In Netlify project settings, add this environment variable:

```sh
GEMINI_API_KEY=your_key_here
```

Optional model override:

```sh
GUIDEN_LLM_MODEL=gemini-3.5-flash
```

When hosted, the app calls `/.netlify/functions/search`. Locally, it still calls `http://127.0.0.1:8787/api/search`.

## Product notes

Before a real launch, this app should add:

- Account/auth and anonymous identity controls.
- Server-side story storage.
- Human and automated moderation before publication.
- Strong PII removal and self-harm content handling.
- AI search using embeddings over approved stories only.
- Clear clinical disclaimers and crisis escalation by country.
