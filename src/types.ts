export type Story = {
  id: string;
  sourceUrl: string;
  subreddit: string;
  title: string;
  authorName: string;
  category: string;
  categories: string[];
  excerpt: string;
  story: string;
  whatHelped: string[];
  tags: string[];
  supportType: string;
  timeframe: string;
  tone: string;
  readMinutes: number;
  helpfulCount: number;
  moderationStatus: string;
};

export type StoryDraft = {
  title: string;
  categories: string[];
  background: string;
  struggledWith: string;
  tried: string;
  helped: string;
  timeline: string;
  lessons: string;
  advice: string;
  resources: string;
  tags: string;
  anonymous: boolean;
};

export type MatchResult = Story & {
  matchScore: number;
  reasons: string[];
};
