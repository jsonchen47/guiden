import { MatchResult, Story } from "../types";

const normalize = (value: string) =>
  value
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, " ")
    .split(/\s+/)
    .filter(Boolean);

const unique = (items: string[]) => Array.from(new Set(items));

export function matchStories(query: string, stories: Story[]): MatchResult[] {
  const terms = unique(normalize(query));

  if (terms.length === 0) {
    return stories.map((story) => ({
      ...story,
      matchScore: 0,
      reasons: ["Popular testimony"]
    }));
  }

  return stories
    .map((story) => {
      const searchable = normalize(
        [
          story.title,
          story.category,
          story.categories.join(" "),
          story.excerpt,
          story.story,
          story.tags.join(" "),
          story.whatHelped.join(" "),
          story.supportType,
          story.tone
        ].join(" ")
      );
      const searchableSet = new Set(searchable);
      const matchedTerms = terms.filter((term) => searchableSet.has(term));
      const tagMatches = story.tags.filter((tag) =>
        terms.some((term) => normalize(tag).includes(term))
      );

      const score = matchedTerms.length * 12 + tagMatches.length * 10;
      const reasons = [
        ...tagMatches.map((tag) => `Shared theme: ${tag}`),
        matchedTerms.length > 0 ? `${matchedTerms.length} phrase match${matchedTerms.length === 1 ? "" : "es"}` : ""
      ].filter(Boolean);

      return {
        ...story,
        matchScore: score,
        reasons: reasons.length > 0 ? reasons.slice(0, 3) : ["Adjacent experience"]
      };
    })
    .sort((a, b) => b.matchScore - a.matchScore);
}
