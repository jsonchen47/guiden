import { StatusBar } from "expo-status-bar";
import React, { useEffect, useMemo, useState } from "react";
import {
  KeyboardAvoidingView,
  Platform,
  Pressable,
  SafeAreaView,
  ScrollView,
  StyleSheet,
  Text,
  TextInput,
  useWindowDimensions,
  View
} from "react-native";
import { starterStories } from "./src/data/stories";
import { matchStories } from "./src/lib/matching";
import { MatchResult, Story, StoryDraft } from "./src/types";

const appFontFamily = Platform.select({
  web: "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
  default: "System"
});
const displayFontFamily = Platform.select({
  web: "Georgia, 'Times New Roman', ui-serif, serif",
  default: "Georgia"
});

(Text as any).defaultProps = (Text as any).defaultProps || {};
(Text as any).defaultProps.style = [{ fontFamily: appFontFamily }, (Text as any).defaultProps.style];
(TextInput as any).defaultProps = (TextInput as any).defaultProps || {};
(TextInput as any).defaultProps.style = [{ fontFamily: appFontFamily }, (TextInput as any).defaultProps.style];

type IconName =
  | "activity"
  | "arrow-left"
  | "arrow-right"
  | "book-open"
  | "check"
  | "chevron-down"
  | "clock"
  | "compass"
  | "cpu"
  | "edit-3"
  | "heart"
  | "message-circle"
  | "search"
  | "send"
  | "sliders"
  | "thumbs-up";

const iconMap: Record<IconName, string> = {
  activity: "!",
  "arrow-left": "<",
  "arrow-right": ">",
  "book-open": "[]",
  check: "✓",
  "chevron-down": "⌄",
  clock: "○",
  compass: "◆",
  cpu: "*",
  "edit-3": "+",
  heart: "♥",
  "message-circle": "''",
  search: "⌕",
  send: ">",
  sliders: "≡",
  "thumbs-up": "^"
};

function Feather({ name, size, color }: { name: IconName; size: number; color: string }) {
  return (
    <Text style={{ color, fontSize: size, fontWeight: "800", lineHeight: size + 2 }}>
      {iconMap[name]}
    </Text>
  );
}

type Page =
  | { name: "home" }
  | { name: "stories"; category?: string }
  | { name: "search"; query?: string }
  | { name: "share" }
  | { name: "detail"; id: string };

type PatternInsight = {
  label: string;
  description: string;
  count: number;
  percent: number;
  examples: string[];
  evidence?: string;
};

type PatternDefinition = {
  label: string;
  description: string;
  terms: string[];
};

const categories = [
  "Anxiety",
  "Depression",
  "Burnout",
  "OCD",
  "ADHD",
  "PTSD",
  "Bipolar Disorder",
  "Eating Disorders",
  "Grief",
  "Chronic Illness",
  "Cancer",
  "Relationship Issues",
  "Loneliness",
  "Academic Stress",
  "Medical Student Burnout",
  "Other"
];

const prompts = [
  "Struggling with burnout in medical school",
  "Panic attacks before work",
  "Depression that worsens at night",
  "Managing ADHD without medication"
];

const initialDraft: StoryDraft = {
  title: "",
  categories: [],
  background: "",
  struggledWith: "",
  tried: "",
  helped: "",
  timeline: "",
  lessons: "",
  advice: "",
  resources: "",
  tags: "",
  anonymous: true
};

const fallbackPatternAliases: PatternDefinition[] = [
  {
    label: "moved their body regularly",
    description: "People mentioned walks, cardio, yoga, gym routines, or other steady movement as part of feeling less stuck.",
    terms: ["exercise", "walk", "walking", "cardio", "running", "gym", "fitness", "yoga", "workout"]
  },
  {
    label: "worked with a therapist or structured treatment",
    description: "Stories often named CBT, ERP, a therapist, psychiatrist, counseling, or a program that gave them a plan.",
    terms: ["therapy", "therapist", "psychiatrist", "cbt", "erp", "professional", "counseling", "hospitalization"]
  },
  {
    label: "talked with a clinician about medication",
    description: "Some people described starting, reviewing, tapering, or adjusting medication with medical support.",
    terms: ["medication", "meds", "lexapro", "ssri", "antidepressant", "propranolol", "hydroxyzine", "benzo", "xanax"]
  },
  {
    label: "practiced acceptance, grounding, or mindfulness",
    description: "People described letting anxious feelings pass, meditating, breathing, grounding, or observing thoughts without fighting them.",
    terms: ["mindfulness", "meditation", "acceptance", "letting", "observe", "grounding", "breathing"]
  },
  {
    label: "reached out instead of handling it alone",
    description: "Stories mentioned telling friends or family, joining groups, leaning on a partner, or building a support system.",
    terms: ["friend", "friends", "family", "support system", "community", "group", "sister", "partner"]
  },
  {
    label: "made sleep or nighttime more predictable",
    description: "People described bedtime routines, sleep schedules, or changing how they handled difficult nights.",
    terms: ["sleep", "bedtime", "insomnia", "night", "schedule"]
  },
  {
    label: "returned gradually to avoided situations",
    description: "People practiced driving, leaving the house, going into stores, or facing feared situations in small steps.",
    terms: ["exposure", "driving", "store", "leaving", "agoraphobia", "gradual", "practice"]
  },
  {
    label: "wrote things down and challenged thoughts",
    description: "Stories mentioned journaling, tracking wins, naming anxious predictions, or reframing thoughts.",
    terms: ["journal", "journaling", "writing", "thought", "reframe", "prediction", "challenge", "checklist"]
  },
  {
    label: "reduced pressure or set boundaries",
    description: "People described rest, boundaries, breaks, or lowering unsustainable expectations.",
    terms: ["boundary", "boundaries", "rest", "quit", "reduced", "break", "pressure"]
  }
];

const normalizeText = (value: string) => value.toLowerCase().replace(/[^a-z0-9\s/]/g, " ");

const escapeRegExp = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

function getLlmSearchUrl() {
  if (Platform.OS !== "web" || typeof window === "undefined") {
    return "http://127.0.0.1:8787/api/search";
  }

  const hostname = window.location.hostname;
  if (hostname === "localhost" || hostname === "127.0.0.1") {
    return "http://127.0.0.1:8787/api/search";
  }

  return "/.netlify/functions/search";
}

function isHostedDemo() {
  if (Platform.OS !== "web" || typeof window === "undefined") {
    return false;
  }

  const hostname = window.location.hostname;
  return hostname !== "localhost" && hostname !== "127.0.0.1";
}

function includesPatternTerm(text: string, term: string) {
  const normalizedTerm = normalizeText(term).trim();
  if (!normalizedTerm) return false;
  if (normalizedTerm.includes(" ")) {
    return text.includes(normalizedTerm);
  }
  return new RegExp(`\\b${escapeRegExp(normalizedTerm)}\\b`).test(text);
}

function buildPatternInsights(matches: MatchResult[], totalStories: number): PatternInsight[] {
  const source = matches.filter((story) => story.matchScore > 0).slice(0, 24);
  const analysisSet = source.length >= 6 ? source : matches.slice(0, Math.min(24, totalStories));

  return fallbackPatternAliases
    .map((pattern) => {
      const storiesWithPattern = analysisSet.filter((story) => {
        const searchable = normalizeText(
          [story.whatHelped.join(" "), story.tags.join(" "), story.story, story.title].join(" ")
        );
        return pattern.terms.some((term) => includesPatternTerm(searchable, term));
      });

      return {
        label: pattern.label,
        description: pattern.description,
        count: storiesWithPattern.length,
        percent: analysisSet.length ? Math.round((storiesWithPattern.length / analysisSet.length) * 100) : 0,
        examples: storiesWithPattern
          .flatMap((story) => story.whatHelped)
          .filter((item) => {
            const normalizedItem = normalizeText(item);
            return pattern.terms.some((term) => includesPatternTerm(normalizedItem, term));
          })
          .slice(0, 3)
      };
    })
    .filter((insight) => insight.count > 0)
    .sort((a, b) => b.count - a.count || b.percent - a.percent)
    .slice(0, 5);
}

function getInitialPage(): Page {
  if (Platform.OS !== "web" || typeof window === "undefined") {
    return { name: "home" };
  }

  const path = window.location.pathname;
  const params = new URLSearchParams(window.location.search);
  if (path.startsWith("/stories/")) {
    return { name: "detail", id: decodeURIComponent(path.replace("/stories/", "")) };
  }
  if (path === "/stories") {
    return { name: "stories", category: params.get("category") || undefined };
  }
  if (path === "/search") {
    return { name: "search", query: params.get("q") || undefined };
  }
  if (path === "/share") {
    return { name: "share" };
  }
  return { name: "home" };
}

function pageToPath(page: Page) {
  if (page.name === "home") return "/";
  if (page.name === "stories") {
    return page.category ? `/stories?category=${encodeURIComponent(page.category)}` : "/stories";
  }
  if (page.name === "search") {
    return page.query ? `/search?q=${encodeURIComponent(page.query)}` : "/search";
  }
  if (page.name === "share") return "/share";
  return `/stories/${encodeURIComponent(page.id)}`;
}

export default function App() {
  const [page, setPage] = useState<Page>(getInitialPage);
  const [stories, setStories] = useState<Story[]>(starterStories);
  const [draft, setDraft] = useState<StoryDraft>(initialDraft);
  const { width } = useWindowDimensions();
  const compact = width < 760;

  const navigate = (nextPage: Page) => {
    setPage(nextPage);
    if (Platform.OS === "web" && typeof window !== "undefined") {
      window.history.pushState({}, "", pageToPath(nextPage));
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  };

  const submitStory = () => {
    if (!draft.title.trim() || !draft.background.trim() || !draft.helped.trim()) {
      return;
    }

    const story: Story = {
      id: `community-${Date.now()}`,
      sourceUrl: "",
      subreddit: "Guiden",
      title: draft.title.trim(),
      authorName: draft.anonymous ? "Anonymous" : "CommunityMember",
      category: draft.categories[0] || "Other",
      categories: draft.categories.length ? draft.categories : ["Other"],
      excerpt: draft.background.trim().slice(0, 260),
      story: [
        draft.background,
        draft.struggledWith,
        draft.tried,
        draft.helped,
        draft.timeline,
        draft.lessons,
        draft.advice,
        draft.resources
      ]
        .filter((part) => part.trim())
        .join("\n\n"),
      whatHelped: draft.helped
        .split(/\n|,|;/)
        .map((item) => item.trim())
        .filter(Boolean),
      tags: draft.tags
        .split(",")
        .map((item) => item.trim().toLowerCase())
        .filter(Boolean),
      supportType: "Community",
      timeframe: draft.timeline || "Shared recently",
      tone: "community submitted",
      readMinutes: 6,
      helpfulCount: 0,
      moderationStatus: "prototype_submission"
    };

    setStories((current) => [story, ...current]);
    setDraft(initialDraft);
    navigate({ name: "detail", id: story.id });
  };

  const currentStory =
    page.name === "detail" ? stories.find((story) => story.id === page.id) || stories[0] : stories[0];

  return (
    <SafeAreaView style={styles.safeArea}>
      <StatusBar style="dark" />
      <KeyboardAvoidingView behavior={Platform.OS === "ios" ? "padding" : undefined} style={styles.app}>
        <SiteHeader page={page} navigate={navigate} compact={compact} />
        <ScrollView style={styles.scroll} contentContainerStyle={styles.scrollContent}>
          {page.name === "home" && <HomeScreen stories={stories} navigate={navigate} compact={compact} />}
          {page.name === "stories" && (
            <StoriesScreen
              stories={stories}
              activeCategory={page.category}
              navigate={navigate}
              compact={compact}
            />
          )}
          {page.name === "search" && (
            <SearchScreen
              stories={stories}
              initialQuery={page.query}
              navigate={navigate}
              compact={compact}
            />
          )}
          {page.name === "share" && (
            <ShareScreen
              draft={draft}
              setDraft={setDraft}
              submitStory={submitStory}
              compact={compact}
            />
          )}
          {page.name === "detail" && (
            <StoryDetail story={currentStory} navigate={navigate} compact={compact} />
          )}
          <SiteFooter compact={compact} />
        </ScrollView>
      </KeyboardAvoidingView>
    </SafeAreaView>
  );
}

function SiteHeader({
  page,
  navigate,
  compact
}: {
  page: Page;
  navigate: (page: Page) => void;
  compact: boolean;
}) {
  const itemStyle = (active: boolean) => [styles.navButton, active && styles.navButtonActive];

  return (
    <View style={[styles.header, compact && styles.headerCompact]}>
      <Pressable accessibilityRole="link" onPress={() => navigate({ name: "home" })} style={styles.brandRow}>
        <View style={styles.logoMark}>
          <Feather name="compass" size={18} color="#F5F3EE" />
        </View>
        <Text style={styles.brand}>Guiden</Text>
      </Pressable>
      <View style={[styles.nav, compact && styles.navCompact]}>
        <Pressable accessibilityRole="button" onPress={() => navigate({ name: "home" })} style={itemStyle(page.name === "home")}>
          <Feather name="heart" size={18} color={page.name === "home" ? colors.green : colors.ink} />
          <Text style={styles.navText}>Home</Text>
        </Pressable>
        <Pressable accessibilityRole="button" onPress={() => navigate({ name: "stories" })} style={itemStyle(page.name === "stories" || page.name === "detail")}>
          <Feather name="book-open" size={18} color={page.name === "stories" || page.name === "detail" ? colors.green : colors.ink} />
          <Text style={styles.navText}>Stories</Text>
        </Pressable>
        <Pressable accessibilityRole="button" onPress={() => navigate({ name: "search" })} style={itemStyle(page.name === "search")}>
          <Feather name="search" size={18} color={page.name === "search" ? colors.green : colors.ink} />
          <Text style={styles.navText}>AI Search</Text>
        </Pressable>
        <Pressable accessibilityRole="button" onPress={() => navigate({ name: "share" })} style={itemStyle(page.name === "share")}>
          <Feather name="edit-3" size={18} color={page.name === "share" ? colors.green : colors.ink} />
          <Text style={styles.navText}>Share Your Story</Text>
        </Pressable>
      </View>
    </View>
  );
}

function HomeScreen({
  stories,
  navigate,
  compact
}: {
  stories: Story[];
  navigate: (page: Page) => void;
  compact: boolean;
}) {
  const [query, setQuery] = useState("");
  const featured = stories.slice(0, 3);

  const submitSearch = () => {
    if (query.trim()) {
      navigate({ name: "search", query: query.trim() });
    }
  };

  return (
    <View>
      <View style={[styles.hero, compact && styles.heroCompact]}>
        <View style={styles.heroOverlay} />
        <View style={styles.heroInner}>
          <Text style={[styles.heroTitle, compact && styles.heroTitleCompact]}>
            Feeling off or overwhelmed?{"\n"}
            <Text style={styles.heroTitleItalic}>See what helped others.</Text>
          </Text>
          <Text style={styles.heroCopy}>
            Search lived experiences, from everyday low moods to bigger turning points, and spot the patterns that helped people feel better.
          </Text>
          <View style={[styles.searchBox, compact && styles.searchBoxCompact]}>
            <TextInput
              value={query}
              onChangeText={setQuery}
              placeholder="Describe how you're feeling..."
              placeholderTextColor="#667268"
              style={styles.heroInput}
              onSubmitEditing={submitSearch}
            />
            <Pressable
              accessibilityRole="button"
              disabled={!query.trim()}
              onPress={submitSearch}
              style={[styles.primaryButton, styles.heroSearchButton, !query.trim() && styles.buttonDisabled]}
            >
              <Feather name="search" size={17} color="#F5F3EE" />
              <Text style={[styles.primaryButtonText, styles.heroSearchButtonText]}>Search Library</Text>
            </Pressable>
          </View>
          <View style={styles.promptWrap}>
            {prompts.map((prompt) => (
              <Pressable accessibilityRole="button" key={prompt} onPress={() => navigate({ name: "search", query: prompt })} style={styles.promptPill}>
                <Text style={[styles.promptText, styles.promptTextOnDark]}>{prompt}</Text>
              </Pressable>
            ))}
          </View>
        </View>
      </View>

      <SectionHeader
        title="Featured stories"
        subtitle="Real people, real moments, real things that helped."
        action="View all"
        onAction={() => navigate({ name: "stories" })}
      />
      <View style={[styles.cardGrid, compact && styles.singleColumn]}>
        {featured.map((story) => (
          <StoryCard key={story.id} story={story} navigate={navigate} />
        ))}
      </View>

      <SectionHeader title="Browse by topic" />
      <View style={styles.topicGrid}>
        {categories.map((category) => (
          <Pressable
            key={category}
            accessibilityRole="button"
            onPress={() => navigate({ name: "stories", category })}
            style={styles.topicTile}
          >
            <Text style={styles.topicText}>{category}</Text>
          </Pressable>
        ))}
      </View>

      <View style={[styles.calloutBand, compact && styles.calloutBandCompact]}>
        <View style={styles.calloutTextBlock}>
          <Text style={styles.calloutTitle}>Your story might be exactly what someone needs to hear.</Text>
          <Text style={styles.calloutCopy}>
            Sharing your experience, even the messy and non-linear kind, helps others feel less alone and adds to the collective wisdom of this community.
          </Text>
        </View>
        <Pressable accessibilityRole="button" onPress={() => navigate({ name: "share" })} style={styles.primaryButton}>
          <Feather name="edit-3" size={17} color="#F5F3EE" />
          <Text style={styles.primaryButtonText}>Share your story</Text>
        </Pressable>
      </View>
    </View>
  );
}

function StoriesScreen({
  stories,
  activeCategory,
  navigate,
  compact
}: {
  stories: Story[];
  activeCategory?: string;
  navigate: (page: Page) => void;
  compact: boolean;
}) {
  const [sort, setSort] = useState<"helpful" | "newest">("helpful");

  const filtered = useMemo(() => {
    const list = activeCategory
      ? stories.filter((story) => story.categories.includes(activeCategory) || story.category === activeCategory)
      : stories;
    return [...list].sort((a, b) =>
      sort === "helpful" ? b.helpfulCount - a.helpfulCount : a.id.localeCompare(b.id)
    );
  }, [activeCategory, sort, stories]);

  return (
    <View style={styles.pageShell}>
      <Text style={styles.pageTitle}>Stories Library</Text>
      <Text style={styles.pageIntro}>Browse {stories.length}+ authentic stories about feeling better, getting unstuck, and finding small next steps.</Text>
      <View style={[styles.filterBar, compact && styles.filterBarCompact]}>
        <Pressable accessibilityRole="button" onPress={() => navigate({ name: "stories" })} style={styles.selectControl}>
          <Text style={styles.selectText}>{activeCategory || "All Categories"}</Text>
          <Feather name="chevron-down" size={16} color="#1E2820" />
        </Pressable>
        <Pressable accessibilityRole="button" onPress={() => setSort(sort === "helpful" ? "newest" : "helpful")} style={styles.selectControl}>
          <Text style={styles.selectText}>{sort === "helpful" ? "Most Helpful" : "Newest"}</Text>
          <Feather name="sliders" size={16} color="#1E2820" />
        </Pressable>
      </View>
      {activeCategory && (
        <View style={styles.activeFilterRow}>
          <Text style={styles.resultMeta}>Filtered by {activeCategory}</Text>
          <Pressable accessibilityRole="button" onPress={() => navigate({ name: "stories" })}>
            <Text style={styles.textLink}>Clear</Text>
          </Pressable>
        </View>
      )}
      <Text style={styles.resultMeta}>Showing {filtered.length} stories</Text>
      <View style={styles.storyList}>
        {filtered.slice(0, 30).map((story) => (
          <WideStoryCard key={story.id} story={story} navigate={navigate} />
        ))}
      </View>
    </View>
  );
}

function SearchScreen({
  stories,
  initialQuery,
  navigate,
  compact
}: {
  stories: Story[];
  initialQuery?: string;
  navigate: (page: Page) => void;
  compact: boolean;
}) {
  const [query, setQuery] = useState(initialQuery || "");
  const [llmStatus, setLlmStatus] = useState<"idle" | "loading" | "ready" | "error">("idle");
  const [llmInsights, setLlmInsights] = useState<PatternInsight[]>([]);
  const [llmSummary, setLlmSummary] = useState("");
  const [llmError, setLlmError] = useState("");
  const [llmProvider, setLlmProvider] = useState("");
  const [llmModel, setLlmModel] = useState("");
  const results = useMemo(() => matchStories(query, stories), [query, stories]);
  const fallbackInsights = useMemo(() => buildPatternInsights(results, stories.length), [results, stories.length]);
  const hasQuery = query.trim().length > 0;
  const analysisStories = useMemo(() => {
    const matchedSample = results.filter((story) => story.matchScore > 0);
    return (matchedSample.length >= 6 ? matchedSample : results).slice(0, Math.min(24, results.length));
  }, [results]);
  const analysisStoryIds = useMemo(() => analysisStories.map((story) => story.id), [analysisStories]);
  const analysisStoryKey = analysisStoryIds.join("|");
  const sampleSize = analysisStories.length;

  useEffect(() => {
    if (!hasQuery) {
      setLlmStatus("idle");
      setLlmInsights([]);
      setLlmSummary("");
      setLlmError("");
      setLlmProvider("");
      setLlmModel("");
      return;
    }

    const controller = new AbortController();
    const timer = setTimeout(async () => {
      setLlmStatus("loading");
      setLlmError("");

      try {
        const response = await fetch(getLlmSearchUrl(), {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ query: query.trim(), storyIds: analysisStoryIds }),
          signal: controller.signal
        });

        if (!response.ok) {
          const payload = await response.json().catch(() => null);
          if (payload?.provider) setLlmProvider(String(payload.provider));
          if (payload?.model) setLlmModel(String(payload.model));
          throw new Error(payload?.error || `LLM search failed with ${response.status}`);
        }

        const data = await response.json();
        setLlmInsights(Array.isArray(data.patterns) ? data.patterns : []);
        setLlmSummary(typeof data.summary === "string" ? data.summary : "");
        setLlmProvider(typeof data.provider === "string" ? data.provider : "");
        setLlmModel(typeof data.model === "string" ? data.model : "");
        setLlmStatus("ready");
      } catch (error) {
        if (controller.signal.aborted) return;
        setLlmStatus("error");
        setLlmError(error instanceof Error ? error.message : "The LLM search server is not available.");
      }
    }, 650);

    return () => {
      controller.abort();
      clearTimeout(timer);
    };
  }, [analysisStoryKey, hasQuery, query]);

  return (
    <View style={styles.pageShell}>
      <View style={styles.kickerRow}>
        <Feather name="cpu" size={18} color="#C86F4C" />
        <Text style={styles.kicker}>LLM-Powered Pattern Search</Text>
      </View>
      <Text style={styles.pageTitle}>AI Search</Text>
      <Text style={styles.pageIntro}>
        Describe how you feel in natural language. Guiden sends similar stories to an LLM, which turns repeated helpful actions into evidence-backed suggestions.
      </Text>
      <View style={[styles.searchPanelCompact, compact && styles.searchPanelCompactMobile]}>
        <TextInput
          value={query}
          onChangeText={setQuery}
          placeholder="I feel kind of sad at night and don't know what helps..."
          placeholderTextColor="#667268"
          style={styles.searchInputCompact}
          onSubmitEditing={() => setQuery(query.trim())}
        />
        <Pressable disabled={!hasQuery} style={[styles.primaryButton, !hasQuery && styles.buttonDisabled]}>
          <Feather name="search" size={17} color="#F5F3EE" />
          <Text style={styles.primaryButtonText}>Analyze Story Patterns</Text>
        </Pressable>
      </View>
      {!hasQuery && (
        <>
          <Text style={styles.subheading}>Try these examples</Text>
          <View style={styles.promptWrap}>
            {[
              "I have anxiety and panic attacks before exams",
              "I feel a little sad every few nights",
              "My depression gets worse at night",
              "I'm struggling with burnout during medical school",
              "How to manage ADHD without medication"
            ].map((prompt) => (
              <Pressable accessibilityRole="button" key={prompt} onPress={() => setQuery(prompt)} style={styles.promptPillLight}>
                <Text style={styles.promptText}>"{prompt}"</Text>
              </Pressable>
            ))}
          </View>
        </>
      )}
      {hasQuery && (
        <View>
          <PatternInsights
            error={llmError}
            fallbackInsights={fallbackInsights}
            insights={llmInsights}
            sampleSize={sampleSize}
            status={llmStatus}
            summary={llmSummary}
            provider={llmProvider}
            model={llmModel}
            hosted={isHostedDemo()}
          />
          <SectionHeader title="Closest matches" subtitle={`${results.length} stories ranked by shared language and themes.`} />
          <View style={[styles.cardGrid, compact && styles.singleColumn]}>
            {results.slice(0, 9).map((story) => (
              <StoryCard key={story.id} story={story} navigate={navigate} showReasons />
            ))}
          </View>
        </View>
      )}
      <View style={styles.infoPanel}>
        <Text style={styles.infoTitle}>How AI Search Works</Text>
        <Text style={styles.bodyCopy}>
          The app calls a local server at http://127.0.0.1:8787/api/search. That server can use Gemini, Ollama, or OpenAI depending on how you started npm run llm-search, then asks for concrete patterns with counts and evidence.
        </Text>
      </View>
    </View>
  );
}

function PatternInsights({
  error,
  fallbackInsights,
  insights,
  sampleSize,
  status,
  summary,
  provider,
  model,
  hosted
}: {
  error: string;
  fallbackInsights: PatternInsight[];
  insights: PatternInsight[];
  sampleSize: number;
  status: "idle" | "loading" | "ready" | "error";
  summary: string;
  provider: string;
  model: string;
  hosted: boolean;
}) {
  const [loadingDots, setLoadingDots] = useState(".");
  const visibleInsights = status === "ready" && insights.length > 0 ? insights : fallbackInsights;
  const usingFallback = status !== "ready" || insights.length === 0;

  useEffect(() => {
    if (status !== "loading") {
      setLoadingDots(".");
      return;
    }

    const interval = setInterval(() => {
      setLoadingDots((current) => (current.length >= 3 ? "." : `${current}.`));
    }, 420);

    return () => clearInterval(interval);
  }, [status]);

  if (status === "loading") {
    return (
      <View style={styles.aiSummaryPanel}>
        <View style={styles.aiSummaryHeader}>
          <Feather name="cpu" size={18} color="#C86F4C" />
          <Text style={styles.infoTitle}>Asking Gemini{loadingDots}</Text>
        </View>
        <Text style={styles.bodyCopy}>
          Guiden is sending a matched set of similar stories to {hosted ? "the hosted Gemini function" : "the local LLM search server"} and asking it for concrete, evidence-backed suggestions.
        </Text>
        <View style={styles.loadingDotsRow}>
          {[0, 1, 2].map((index) => (
            <View
              key={index}
              style={[
                styles.loadingDot,
                loadingDots.length > index && styles.loadingDotActive
              ]}
            />
          ))}
        </View>
      </View>
    );
  }

  if (visibleInsights.length === 0) {
    return (
      <View style={styles.aiSummaryPanel}>
        <View style={styles.aiSummaryHeader}>
          <Feather name="cpu" size={18} color="#C86F4C" />
          <Text style={styles.infoTitle}>LLM pattern summary</Text>
        </View>
        <Text style={styles.bodyCopy}>
          Guiden did not find a strong repeated pattern yet. Try describing the feeling, situation, and anything you have already tried.
        </Text>
      </View>
    );
  }

  const top = visibleInsights[0];

  return (
    <View style={styles.aiSummaryPanel}>
      <View style={styles.aiSummaryHeader}>
        <Feather name="cpu" size={18} color="#C86F4C" />
        <Text style={styles.infoTitle}>
          {usingFallback ? "Local fallback summary" : `${provider || "LLM"} pattern summary`}
        </Text>
      </View>
      {status === "error" && (
        <View style={styles.warningBox}>
          <Text style={styles.warningText}>
            {provider
              ? `${hosted ? "The hosted Gemini function" : "The local API server"} is running with ${provider}${model ? ` (${model})` : ""}, but the model response could not be used.`
              : hosted
                ? "The hosted Gemini function is not responding yet. Check the Netlify Function logs and make sure GEMINI_API_KEY is set."
                : "The local API server is not connected yet. Start it with npm run llm-search."}{" "}
            {error}
          </Text>
        </View>
      )}
      {!usingFallback && Boolean(summary) && <Text style={styles.aiSummaryLead}>{summary}</Text>}
      <Text style={styles.aiSummaryLead}>
        Based on {sampleSize} similar stories, the strongest evidence-backed suggestion is: {top.label}. {top.count} of {sampleSize} stories mention it.
      </Text>
      <View style={styles.insightGrid}>
        {visibleInsights.map((insight) => (
          <View key={insight.label} style={styles.insightCard}>
            <View style={styles.insightTopline}>
              <Text style={styles.insightCount}>{insight.count}/{sampleSize}</Text>
              <Text style={styles.insightPercent}>{insight.percent}%</Text>
            </View>
            <Text style={styles.insightTitle}>{insight.label}</Text>
            <Text style={styles.insightDescription}>{insight.description}</Text>
            {insight.examples.length > 0 && (
              <Text style={styles.insightExamples}>Examples: {insight.examples.join(", ")}</Text>
            )}
            {Boolean(insight.evidence) && <Text style={styles.insightExamples}>Evidence: {insight.evidence}</Text>}
          </View>
        ))}
      </View>
      <Text style={styles.aiCaveat}>
        These are community patterns, not medical advice. Use them as ideas to discuss, not instructions to follow blindly.
      </Text>
    </View>
  );
}

function ShareScreen({
  draft,
  setDraft,
  submitStory,
  compact
}: {
  draft: StoryDraft;
  setDraft: React.Dispatch<React.SetStateAction<StoryDraft>>;
  submitStory: () => void;
  compact: boolean;
}) {
  const canSubmit = draft.title.trim() && draft.background.trim() && draft.helped.trim();
  const toggleCategory = (category: string) => {
    setDraft((current) => ({
      ...current,
      categories: current.categories.includes(category)
        ? current.categories.filter((item) => item !== category)
        : [...current.categories, category]
    }));
  };

  return (
    <View style={styles.pageShell}>
      <View style={styles.kickerRow}>
        <Feather name="heart" size={18} color="#C86F4C" />
        <Text style={styles.kicker}>Share Hope</Text>
      </View>
      <Text style={styles.pageTitle}>Share Your Recovery Story</Text>
      <Text style={styles.pageIntro}>
        Your experience can give hope to someone who's struggling right now. This prototype stores the submission locally.
      </Text>
      <View style={styles.infoPanel}>
        <Text style={styles.infoTitle}>Your Story Matters</Text>
        <Text style={styles.bodyCopy}>
          You do not need to be fully recovered to share. Stories about learning to manage, finding what helps, or making progress are all valuable.
        </Text>
      </View>
      <View style={styles.formStack}>
        <Field
          label="Story Title *"
          value={draft.title}
          onChangeText={(title) => setDraft((current) => ({ ...current, title }))}
          placeholder="Give your story a descriptive title"
        />
        <Pressable
          accessibilityRole="checkbox"
          onPress={() => setDraft((current) => ({ ...current, anonymous: !current.anonymous }))}
          style={styles.checkboxRow}
        >
          <View style={[styles.checkbox, draft.anonymous && styles.checkboxActive]}>
            {draft.anonymous && <Feather name="check" size={14} color="#F5F3EE" />}
          </View>
          <Text style={styles.bodyCopy}>Post anonymously</Text>
        </Pressable>
        <Text style={styles.fieldLabel}>Categories *</Text>
        <View style={styles.promptWrap}>
          {categories.map((category) => (
            <Pressable
              accessibilityRole="button"
              key={category}
              onPress={() => toggleCategory(category)}
              style={[styles.categoryButton, draft.categories.includes(category) && styles.categoryButtonActive]}
            >
              <Text style={[styles.categoryText, draft.categories.includes(category) && styles.categoryTextActive]}>
                {category}
              </Text>
            </Pressable>
          ))}
        </View>
        <View style={[styles.twoColumnForm, compact && styles.singleColumn]}>
          <Field
            label="Background *"
            value={draft.background}
            onChangeText={(background) => setDraft((current) => ({ ...current, background }))}
            placeholder="What were you experiencing?"
            multiline
          />
          <Field
            label="What You Struggled With"
            value={draft.struggledWith}
            onChangeText={(struggledWith) => setDraft((current) => ({ ...current, struggledWith }))}
            placeholder="Symptoms, situations, or challenges"
            multiline
          />
          <Field
            label="What You Tried"
            value={draft.tried}
            onChangeText={(tried) => setDraft((current) => ({ ...current, tried }))}
            placeholder="What did not help?"
            multiline
          />
          <Field
            label="What Actually Helped *"
            value={draft.helped}
            onChangeText={(helped) => setDraft((current) => ({ ...current, helped }))}
            placeholder="Treatments, habits, support, changes"
            multiline
          />
          <Field
            label="Timeline of Recovery"
            value={draft.timeline}
            onChangeText={(timeline) => setDraft((current) => ({ ...current, timeline }))}
            placeholder="How long did improvement take?"
            multiline
          />
          <Field
            label="Key Lessons Learned"
            value={draft.lessons}
            onChangeText={(lessons) => setDraft((current) => ({ ...current, lessons }))}
            placeholder="What did this teach you?"
            multiline
          />
          <Field
            label="Advice for Others"
            value={draft.advice}
            onChangeText={(advice) => setDraft((current) => ({ ...current, advice }))}
            placeholder="What would you tell someone similar?"
            multiline
          />
          <Field
            label="Helpful Resources"
            value={draft.resources}
            onChangeText={(resources) => setDraft((current) => ({ ...current, resources }))}
            placeholder="Books, apps, websites, groups"
            multiline
          />
        </View>
        <Field
          label="Tags"
          value={draft.tags}
          onChangeText={(tags) => setDraft((current) => ({ ...current, tags }))}
          placeholder="CBT, sleep, boundaries, panic"
        />
        <View style={styles.guidelinesBox}>
          <Text style={styles.infoTitle}>Community Guidelines</Text>
          <Text style={styles.bodyCopy}>Be honest, avoid giving medical advice, and leave out identifying details about yourself or others.</Text>
        </View>
        <View style={styles.formActions}>
          <Pressable accessibilityRole="button" onPress={() => setDraft(initialDraft)} style={styles.secondaryButton}>
            <Text style={styles.secondaryButtonText}>Cancel</Text>
          </Pressable>
          <Pressable accessibilityRole="button" disabled={!canSubmit} onPress={submitStory} style={[styles.primaryButton, !canSubmit && styles.buttonDisabled]}>
            <Feather name="send" size={17} color="#F5F3EE" />
            <Text style={styles.primaryButtonText}>Share Your Story</Text>
          </Pressable>
        </View>
      </View>
    </View>
  );
}

function StoryDetail({
  story,
  navigate,
  compact
}: {
  story: Story;
  navigate: (page: Page) => void;
  compact: boolean;
}) {
  const paragraphs = story.story.split(/\n\n+/).filter(Boolean).slice(0, 7);
  const related = starterStories
    .filter((candidate) => candidate.id !== story.id && candidate.categories.some((category) => story.categories.includes(category)))
    .slice(0, 2);

  return (
    <View style={styles.pageShell}>
      <Pressable accessibilityRole="button" onPress={() => navigate({ name: "stories" })} style={styles.backButton}>
        <Feather name="arrow-left" size={17} color="#1E2820" />
        <Text style={styles.textLink}>Back to Stories</Text>
      </Pressable>
      <View style={styles.detailHeader}>
        <View style={styles.pillRow}>
          {story.categories.slice(0, 3).map((category) => (
            <Text key={category} style={styles.tagPill}>{category}</Text>
          ))}
        </View>
        <Text style={[styles.detailTitle, compact && styles.detailTitleCompact]}>{story.title}</Text>
        <Text style={styles.byline}>
          {story.authorName} - {story.readMinutes} min read - {story.subreddit || "Guiden"}
        </Text>
        <Pressable accessibilityRole="button" style={styles.helpfulButton}>
          <Feather name="thumbs-up" size={17} color="#1E2820" />
          <Text style={styles.helpfulText}>Mark Helpful ({story.helpfulCount.toLocaleString()})</Text>
        </Pressable>
      </View>

      <View style={styles.articleBody}>
        <ArticleSection icon="book-open" title="Background">
          <Text style={styles.articleText}>{paragraphs[0] || story.excerpt}</Text>
        </ArticleSection>
        <ArticleSection icon="activity" title="What Helped">
          <View style={styles.bulletList}>
            {story.whatHelped.slice(0, 6).map((item) => (
              <View key={item} style={styles.bulletRow}>
                <View style={styles.bulletDot} />
                <Text style={styles.articleText}>{item}</Text>
              </View>
            ))}
          </View>
        </ArticleSection>
        <ArticleSection icon="clock" title="Timeline of Recovery">
          <Text style={styles.articleText}>{story.timeframe}</Text>
        </ArticleSection>
        <ArticleSection icon="message-circle" title="The Story">
          {paragraphs.slice(1).map((paragraph, index) => (
            <Text key={`${story.id}-${index}`} style={styles.articleText}>{paragraph}</Text>
          ))}
        </ArticleSection>
      </View>

      <View style={styles.infoPanel}>
        <Text style={styles.infoTitle}>Related Topics</Text>
        <View style={styles.promptWrap}>
          {story.tags.slice(0, 8).map((tag) => (
            <Text key={tag} style={styles.tagPill}>{tag}</Text>
          ))}
        </View>
      </View>

      {related.length > 0 && (
        <View>
          <SectionHeader title="Related stories" />
          <View style={[styles.cardGrid, compact && styles.singleColumn]}>
            {related.map((candidate) => (
              <StoryCard key={candidate.id} story={candidate} navigate={navigate} />
            ))}
          </View>
        </View>
      )}
    </View>
  );
}

function StoryCard({
  story,
  navigate,
  showReasons = false
}: {
  story: Story | MatchResult;
  navigate: (page: Page) => void;
  showReasons?: boolean;
}) {
  const match = story as MatchResult;

  return (
    <Pressable accessibilityRole="button" onPress={() => navigate({ name: "detail", id: story.id })} style={styles.storyCard}>
      <View style={styles.pillRow}>
        {story.categories.slice(0, 2).map((category) => (
          <Text key={category} style={styles.tagPill}>{category}</Text>
        ))}
      </View>
      <Text style={styles.cardTitle}>{story.title}</Text>
      <Text style={styles.cardExcerpt}>{story.excerpt}</Text>
      {showReasons && (
        <View style={styles.reasonBox}>
          {(match.reasons || ["Adjacent experience"]).slice(0, 2).map((reason) => (
            <Text key={reason} style={styles.reasonText}>{reason}</Text>
          ))}
        </View>
      )}
      <View style={styles.cardMeta}>
        <Text style={styles.metaText}>{story.authorName}</Text>
        <Text style={styles.metaText}>-</Text>
        <Text style={styles.metaText}>{story.readMinutes} min read</Text>
        <View style={styles.helpfulMini}>
          <Feather name="thumbs-up" size={14} color="#667268" />
          <Text style={styles.metaText}>{story.helpfulCount.toLocaleString()}</Text>
        </View>
      </View>
    </Pressable>
  );
}

function WideStoryCard({
  story,
  navigate
}: {
  story: Story;
  navigate: (page: Page) => void;
}) {
  return (
    <Pressable accessibilityRole="button" onPress={() => navigate({ name: "detail", id: story.id })} style={styles.wideStoryCard}>
      <View style={styles.pillRow}>
        {story.categories.slice(0, 3).map((category) => (
          <Text key={category} style={styles.tagPill}>{category}</Text>
        ))}
      </View>
      <Text style={styles.wideTitle}>{story.title}</Text>
      <Text style={styles.byline}>{story.authorName} - {story.readMinutes} min</Text>
      <Text style={styles.cardExcerpt}>{story.excerpt}</Text>
      <View style={styles.cardMeta}>
        {story.tags.slice(0, 3).map((tag) => (
          <Text key={tag} style={styles.smallTag}>{tag}</Text>
        ))}
        <View style={styles.helpfulMini}>
          <Feather name="thumbs-up" size={14} color="#667268" />
          <Text style={styles.metaText}>{story.helpfulCount.toLocaleString()} found helpful</Text>
        </View>
      </View>
    </Pressable>
  );
}

function Field({
  label,
  value,
  onChangeText,
  placeholder,
  multiline
}: {
  label: string;
  value: string;
  onChangeText: (value: string) => void;
  placeholder: string;
  multiline?: boolean;
}) {
  return (
    <View style={styles.field}>
      <Text style={styles.fieldLabel}>{label}</Text>
      <TextInput
        value={value}
        onChangeText={onChangeText}
        placeholder={placeholder}
        placeholderTextColor="#7A837B"
        multiline={multiline}
        textAlignVertical={multiline ? "top" : "center"}
        style={[styles.input, multiline && styles.multilineInput]}
      />
    </View>
  );
}

function SectionHeader({
  title,
  subtitle,
  action,
  onAction
}: {
  title: string;
  subtitle?: string;
  action?: string;
  onAction?: () => void;
}) {
  return (
    <View style={styles.sectionHeader}>
      <View>
        <Text style={styles.sectionTitle}>{title}</Text>
        {subtitle && <Text style={styles.sectionSubtitle}>{subtitle}</Text>}
      </View>
      {action && onAction && (
        <Pressable accessibilityRole="button" onPress={onAction} style={styles.linkButton}>
          <Text style={styles.textLink}>{action}</Text>
          <Feather name="arrow-right" size={16} color="#315A49" />
        </Pressable>
      )}
    </View>
  );
}

function ArticleSection({
  icon,
  title,
  children
}: {
  icon: IconName;
  title: string;
  children: React.ReactNode;
}) {
  return (
    <View style={styles.articleSection}>
      <View style={styles.articleHeadingRow}>
        <Feather name={icon} size={21} color="#C86F4C" />
        <Text style={styles.articleHeading}>{title}</Text>
      </View>
      {children}
    </View>
  );
}

function SiteFooter({ compact }: { compact: boolean }) {
  return (
    <View style={[styles.footer, compact && styles.singleColumn]}>
      <View style={styles.footerColumn}>
        <Text style={styles.footerTitle}>Crisis Resources</Text>
        <Text style={styles.footerText}>National Suicide Prevention Lifeline: 988</Text>
        <Text style={styles.footerText}>Crisis Text Line: Text HOME to 741741</Text>
        <Text style={styles.footerText}>International Association for Suicide Prevention: iasp.info/resources</Text>
      </View>
      <View style={styles.footerColumn}>
        <Text style={styles.footerTitle}>About</Text>
        <Text style={styles.footerText}>Guiden is a community-driven platform sharing real experiences of recovery and resilience.</Text>
      </View>
      <View style={styles.footerColumn}>
        <Text style={styles.footerTitle}>Important Disclaimer</Text>
        <Text style={styles.footerText}>Personal experiences are not a substitute for professional medical advice, diagnosis, or treatment.</Text>
      </View>
    </View>
  );
}

const colors = {
  cream: "#F4F1E9",
  paper: "#FFFEFA",
  ink: "#202821",
  muted: "#7B857C",
  green: "#284D38",
  sage: "#8F9C91",
  border: "#DDD8CD",
  clay: "#C97955",
  paleGreen: "#DFE9DF",
  paleClay: "#F2DED2"
};

const shadows = Platform.select({
  web: {
    boxShadow: "0 18px 45px rgba(30, 40, 32, 0.10)"
  },
  default: {
    shadowColor: "#1E2820",
    shadowOpacity: 0.1,
    shadowRadius: 18,
    shadowOffset: { width: 0, height: 12 },
    elevation: 3
  }
});

const styles = StyleSheet.create({
  safeArea: {
    flex: 1,
    backgroundColor: colors.cream
  },
  app: {
    flex: 1,
    backgroundColor: colors.cream
  },
  scroll: {
    flex: 1
  },
  scrollContent: {
    paddingBottom: 36
  },
  header: {
    alignItems: "center",
    backgroundColor: "#F7F5EF",
    borderBottomColor: "rgba(40, 77, 56, 0.12)",
    borderBottomWidth: 1,
    flexDirection: "row",
    justifyContent: "space-between",
    paddingHorizontal: 36,
    paddingVertical: 14,
    zIndex: 10
  },
  headerCompact: {
    alignItems: "flex-start",
    gap: 14,
    paddingHorizontal: 18,
    flexDirection: "column"
  },
  brandRow: {
    alignItems: "center",
    flexDirection: "row",
    gap: 10
  },
  logoMark: {
    alignItems: "center",
    backgroundColor: colors.green,
    borderRadius: 999,
    height: 36,
    justifyContent: "center",
    width: 36
  },
  brand: {
    color: colors.ink,
    fontFamily: displayFontFamily,
    fontSize: 28,
    fontWeight: "700"
  },
  nav: {
    alignItems: "center",
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8
  },
  navCompact: {
    alignItems: "stretch",
    width: "100%"
  },
  navButton: {
    alignItems: "center",
    borderRadius: 8,
    flexDirection: "row",
    gap: 7,
    paddingHorizontal: 13,
    paddingVertical: 9
  },
  navButtonActive: {
    backgroundColor: colors.paleGreen
  },
  navText: {
    color: colors.ink,
    fontSize: 15,
    fontWeight: "700"
  },
  hero: {
    backgroundColor: colors.green,
    minHeight: 560,
    justifyContent: "center",
    overflow: "hidden",
    paddingHorizontal: 36,
    paddingVertical: 70
  },
  heroCompact: {
    minHeight: 0,
    paddingHorizontal: 18,
    paddingVertical: 48
  },
  heroOverlay: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: colors.green
  },
  heroInner: {
    alignItems: "center",
    alignSelf: "center",
    maxWidth: 980,
    width: "100%"
  },
  heroTitle: {
    color: colors.cream,
    fontFamily: displayFontFamily,
    fontSize: 70,
    fontWeight: "400",
    lineHeight: 78,
    maxWidth: 860,
    textAlign: "center"
  },
  heroTitleItalic: {
    fontFamily: displayFontFamily,
    fontStyle: "italic",
    fontWeight: "400"
  },
  heroTitleAccent: {
    color: colors.clay,
    fontFamily: displayFontFamily,
    fontStyle: "italic",
    fontWeight: "400"
  },
  heroTitleCompact: {
    fontSize: 40,
    lineHeight: 47,
    textAlign: "center"
  },
  heroCopy: {
    alignSelf: "center",
    color: "rgba(244, 241, 233, 0.72)",
    fontSize: 18,
    lineHeight: 29,
    marginTop: 18,
    maxWidth: 720,
    textAlign: "center"
  },
  searchBox: {
    alignItems: "center",
    alignSelf: "center",
    backgroundColor: "#F8F6F0",
    borderRadius: 12,
    flexDirection: "row",
    gap: 12,
    marginTop: 36,
    maxWidth: 860,
    padding: 8,
    width: "100%",
    ...shadows
  },
  searchBoxCompact: {
    alignItems: "stretch",
    flexDirection: "column"
  },
  heroInput: {
    color: colors.ink,
    flex: 1,
    fontSize: 16,
    fontWeight: "600",
    minHeight: 50,
    paddingHorizontal: 14
  },
  primaryButton: {
    alignItems: "center",
    backgroundColor: colors.green,
    borderRadius: 7,
    flexDirection: "row",
    gap: 8,
    justifyContent: "center",
    minHeight: 48,
    paddingHorizontal: 18
  },
  primaryButtonText: {
    color: colors.cream,
    fontSize: 15,
    fontWeight: "700"
  },
  heroSearchButton: {
    backgroundColor: "#3F725C",
    borderRadius: 9,
    gap: 8,
    minHeight: 50,
    paddingHorizontal: 18
  },
  heroSearchButtonText: {
    fontSize: 15,
    fontWeight: "800"
  },
  buttonDisabled: {
    opacity: 0.45
  },
  promptWrap: {
    alignSelf: "center",
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 10,
    marginTop: 16,
    maxWidth: 860
  },
  promptPill: {
    backgroundColor: "rgba(244, 241, 233, 0.04)",
    borderColor: "rgba(244, 241, 233, 0.20)",
    borderRadius: 999,
    borderWidth: 1,
    paddingHorizontal: 14,
    paddingVertical: 9
  },
  promptPillLight: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 999,
    borderWidth: 1,
    paddingHorizontal: 14,
    paddingVertical: 9
  },
  promptText: {
    color: colors.ink,
    fontSize: 14,
    fontWeight: "700"
  },
  promptTextOnDark: {
    color: "rgba(244, 241, 233, 0.66)"
  },
  sectionHeader: {
    alignItems: "flex-end",
    alignSelf: "center",
    flexDirection: "row",
    justifyContent: "space-between",
    maxWidth: 1120,
    paddingHorizontal: 24,
    paddingTop: 52,
    width: "100%"
  },
  sectionTitle: {
    color: colors.ink,
    fontSize: 32,
    fontWeight: "800"
  },
  sectionSubtitle: {
    color: colors.muted,
    fontSize: 16,
    marginTop: 6
  },
  linkButton: {
    alignItems: "center",
    flexDirection: "row",
    gap: 6,
    paddingBottom: 5
  },
  textLink: {
    color: colors.green,
    fontSize: 15,
    fontWeight: "700"
  },
  cardGrid: {
    alignSelf: "center",
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 18,
    maxWidth: 1120,
    paddingHorizontal: 24,
    paddingTop: 18,
    width: "100%"
  },
  singleColumn: {
    flexDirection: "column"
  },
  storyCard: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 8,
    borderWidth: 1,
    flexBasis: 340,
    flexGrow: 1,
    gap: 12,
    minWidth: 280,
    padding: 22,
    ...shadows
  },
  pillRow: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8
  },
  tagPill: {
    backgroundColor: colors.paleGreen,
    borderRadius: 999,
    color: colors.green,
    fontSize: 12,
    fontWeight: "700",
    overflow: "hidden",
    paddingHorizontal: 10,
    paddingVertical: 5
  },
  cardTitle: {
    color: colors.ink,
    fontSize: 22,
    fontWeight: "800",
    lineHeight: 28
  },
  cardExcerpt: {
    color: colors.muted,
    fontSize: 15,
    lineHeight: 23
  },
  reasonBox: {
    backgroundColor: colors.paleClay,
    borderRadius: 7,
    gap: 4,
    padding: 10
  },
  reasonText: {
    color: "#7A3418",
    fontSize: 13,
    fontWeight: "700"
  },
  cardMeta: {
    alignItems: "center",
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
    marginTop: "auto"
  },
  metaText: {
    color: colors.muted,
    fontSize: 13,
    fontWeight: "600"
  },
  helpfulMini: {
    alignItems: "center",
    flexDirection: "row",
    gap: 4,
    marginLeft: "auto"
  },
  topicGrid: {
    alignSelf: "center",
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 12,
    maxWidth: 1120,
    paddingHorizontal: 24,
    paddingTop: 18,
    width: "100%"
  },
  topicTile: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 8,
    borderWidth: 1,
    minWidth: 150,
    paddingHorizontal: 18,
    paddingVertical: 16
  },
  topicText: {
    color: colors.ink,
    fontSize: 15,
    fontWeight: "700",
    textAlign: "center"
  },
  calloutBand: {
    alignItems: "center",
    alignSelf: "center",
    backgroundColor: colors.green,
    borderRadius: 8,
    flexDirection: "row",
    gap: 24,
    justifyContent: "space-between",
    marginTop: 56,
    maxWidth: 1120,
    padding: 34,
    width: "100%"
  },
  calloutBandCompact: {
    alignItems: "flex-start",
    flexDirection: "column",
    marginHorizontal: 18,
    width: "auto"
  },
  calloutTextBlock: {
    flex: 1
  },
  calloutTitle: {
    color: colors.cream,
    fontSize: 28,
    fontWeight: "800"
  },
  calloutCopy: {
    color: "#E9E5DC",
    fontSize: 16,
    lineHeight: 24,
    marginTop: 8
  },
  pageShell: {
    alignSelf: "center",
    maxWidth: 1120,
    paddingHorizontal: 24,
    paddingTop: 28,
    width: "100%"
  },
  pageTitle: {
    color: colors.ink,
    fontSize: 42,
    fontWeight: "800",
    lineHeight: 48
  },
  pageIntro: {
    color: colors.muted,
    fontSize: 16,
    lineHeight: 24,
    marginTop: 6,
    maxWidth: 760
  },
  kickerRow: {
    alignItems: "center",
    flexDirection: "row",
    gap: 8,
    marginBottom: 12
  },
  kicker: {
    color: colors.clay,
    fontSize: 14,
    fontWeight: "800",
    textTransform: "uppercase"
  },
  filterBar: {
    flexDirection: "row",
    gap: 12,
    marginTop: 28
  },
  filterBarCompact: {
    flexDirection: "column"
  },
  selectControl: {
    alignItems: "center",
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 7,
    borderWidth: 1,
    flexDirection: "row",
    gap: 10,
    justifyContent: "space-between",
    minHeight: 46,
    minWidth: 180,
    paddingHorizontal: 14
  },
  selectText: {
    color: colors.ink,
    fontSize: 15,
    fontWeight: "700"
  },
  activeFilterRow: {
    alignItems: "center",
    flexDirection: "row",
    gap: 12,
    marginTop: 18
  },
  resultMeta: {
    color: colors.muted,
    fontSize: 15,
    fontWeight: "600",
    marginTop: 18
  },
  storyList: {
    gap: 16,
    marginTop: 18
  },
  wideStoryCard: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 8,
    borderWidth: 1,
    gap: 10,
    padding: 22
  },
  wideTitle: {
    color: colors.ink,
    fontSize: 24,
    fontWeight: "800",
    lineHeight: 30
  },
  byline: {
    color: colors.muted,
    fontSize: 14,
    fontWeight: "700"
  },
  smallTag: {
    color: colors.green,
    fontSize: 13,
    fontWeight: "700"
  },
  searchPanel: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 8,
    borderWidth: 1,
    gap: 14,
    marginTop: 30,
    padding: 22,
    ...shadows
  },
  searchPanelCompact: {
    alignItems: "center",
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 8,
    borderWidth: 1,
    flexDirection: "row",
    gap: 10,
    marginTop: 18,
    padding: 12,
    ...shadows
  },
  searchPanelCompactMobile: {
    alignItems: "stretch",
    flexDirection: "column"
  },
  subheading: {
    color: colors.ink,
    fontSize: 20,
    fontWeight: "800",
    marginTop: 22
  },
  largeInput: {
    backgroundColor: colors.cream,
    borderColor: colors.border,
    borderRadius: 7,
    borderWidth: 1,
    color: colors.ink,
    fontSize: 16,
    minHeight: 130,
    padding: 14
  },
  searchInputCompact: {
    backgroundColor: colors.cream,
    borderColor: colors.border,
    borderRadius: 7,
    borderWidth: 1,
    color: colors.ink,
    flex: 1,
    fontSize: 16,
    minHeight: 48,
    paddingHorizontal: 14
  },
  infoPanel: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 8,
    borderWidth: 1,
    gap: 8,
    marginTop: 34,
    padding: 22
  },
  aiSummaryPanel: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 8,
    borderWidth: 1,
    gap: 16,
    marginTop: 30,
    padding: 22,
    ...shadows
  },
  aiSummaryHeader: {
    alignItems: "center",
    flexDirection: "row",
    gap: 8
  },
  aiSummaryLead: {
    color: colors.ink,
    fontSize: 20,
    fontWeight: "800",
    lineHeight: 29
  },
  loadingDotsRow: {
    flexDirection: "row",
    gap: 8,
    marginTop: 4
  },
  loadingDot: {
    backgroundColor: colors.border,
    borderRadius: 999,
    height: 9,
    width: 9
  },
  loadingDotActive: {
    backgroundColor: colors.clay
  },
  warningBox: {
    backgroundColor: colors.paleClay,
    borderColor: "#E6B59F",
    borderRadius: 7,
    borderWidth: 1,
    padding: 12
  },
  warningText: {
    color: "#7A3418",
    fontSize: 13,
    fontWeight: "700",
    lineHeight: 19
  },
  insightGrid: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 12
  },
  insightCard: {
    backgroundColor: colors.cream,
    borderColor: colors.border,
    borderRadius: 8,
    borderWidth: 1,
    flexBasis: 190,
    flexGrow: 1,
    gap: 8,
    padding: 15
  },
  insightTopline: {
    alignItems: "center",
    flexDirection: "row",
    justifyContent: "space-between"
  },
  insightCount: {
    color: colors.green,
    fontSize: 24,
    fontWeight: "900"
  },
  insightPercent: {
    color: colors.clay,
    fontSize: 14,
    fontWeight: "900"
  },
  insightTitle: {
    color: colors.ink,
    fontSize: 15,
    fontWeight: "800",
    lineHeight: 20
  },
  insightDescription: {
    color: colors.ink,
    fontSize: 13,
    lineHeight: 19
  },
  insightExamples: {
    color: colors.muted,
    fontSize: 13,
    lineHeight: 19
  },
  aiCaveat: {
    color: colors.muted,
    fontSize: 13,
    fontWeight: "700",
    lineHeight: 19
  },
  infoTitle: {
    color: colors.ink,
    fontSize: 18,
    fontWeight: "800"
  },
  bodyCopy: {
    color: colors.muted,
    fontSize: 15,
    lineHeight: 23
  },
  formStack: {
    gap: 18,
    marginTop: 28
  },
  field: {
    flex: 1,
    gap: 8
  },
  fieldLabel: {
    color: colors.ink,
    fontSize: 14,
    fontWeight: "800"
  },
  input: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 7,
    borderWidth: 1,
    color: colors.ink,
    fontSize: 16,
    minHeight: 48,
    paddingHorizontal: 14
  },
  multilineInput: {
    minHeight: 130,
    paddingTop: 12
  },
  checkboxRow: {
    alignItems: "center",
    flexDirection: "row",
    gap: 10
  },
  checkbox: {
    alignItems: "center",
    borderColor: colors.border,
    borderRadius: 5,
    borderWidth: 1,
    height: 22,
    justifyContent: "center",
    width: 22
  },
  checkboxActive: {
    backgroundColor: colors.green,
    borderColor: colors.green
  },
  categoryButton: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 999,
    borderWidth: 1,
    paddingHorizontal: 13,
    paddingVertical: 9
  },
  categoryButtonActive: {
    backgroundColor: colors.green,
    borderColor: colors.green
  },
  categoryText: {
    color: colors.ink,
    fontSize: 14,
    fontWeight: "700"
  },
  categoryTextActive: {
    color: colors.cream
  },
  twoColumnForm: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 18
  },
  guidelinesBox: {
    backgroundColor: colors.paleGreen,
    borderRadius: 8,
    gap: 8,
    padding: 18
  },
  formActions: {
    alignItems: "center",
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 12,
    justifyContent: "flex-end"
  },
  secondaryButton: {
    alignItems: "center",
    borderColor: colors.border,
    borderRadius: 7,
    borderWidth: 1,
    justifyContent: "center",
    minHeight: 48,
    paddingHorizontal: 18
  },
  secondaryButtonText: {
    color: colors.ink,
    fontSize: 15,
    fontWeight: "800"
  },
  backButton: {
    alignItems: "center",
    flexDirection: "row",
    gap: 8,
    marginBottom: 28
  },
  detailHeader: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 8,
    borderWidth: 1,
    gap: 14,
    padding: 28,
    ...shadows
  },
  detailTitle: {
    color: colors.ink,
    fontSize: 38,
    fontWeight: "800",
    lineHeight: 46
  },
  detailTitleCompact: {
    fontSize: 30,
    lineHeight: 36
  },
  helpfulButton: {
    alignItems: "center",
    alignSelf: "flex-start",
    backgroundColor: colors.paleGreen,
    borderRadius: 7,
    flexDirection: "row",
    gap: 8,
    minHeight: 44,
    paddingHorizontal: 14
  },
  helpfulText: {
    color: colors.ink,
    fontSize: 14,
    fontWeight: "800"
  },
  articleBody: {
    backgroundColor: colors.paper,
    borderColor: colors.border,
    borderRadius: 8,
    borderWidth: 1,
    gap: 26,
    marginTop: 22,
    padding: 28
  },
  articleSection: {
    gap: 12
  },
  articleHeadingRow: {
    alignItems: "center",
    flexDirection: "row",
    gap: 10
  },
  articleHeading: {
    color: colors.ink,
    fontSize: 23,
    fontWeight: "800"
  },
  articleText: {
    color: colors.ink,
    fontSize: 17,
    lineHeight: 28
  },
  bulletList: {
    gap: 10
  },
  bulletRow: {
    alignItems: "flex-start",
    flexDirection: "row",
    gap: 10
  },
  bulletDot: {
    backgroundColor: colors.clay,
    borderRadius: 4,
    height: 8,
    marginTop: 10,
    width: 8
  },
  footer: {
    alignSelf: "center",
    borderTopColor: colors.border,
    borderTopWidth: 1,
    flexDirection: "row",
    gap: 24,
    marginTop: 58,
    maxWidth: 1120,
    paddingHorizontal: 24,
    paddingTop: 28,
    width: "100%"
  },
  footerColumn: {
    flex: 1,
    gap: 8
  },
  footerTitle: {
    color: colors.ink,
    fontSize: 16,
    fontWeight: "800"
  },
  footerText: {
    color: colors.muted,
    fontSize: 14,
    lineHeight: 22
  }
});
