from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


OUTPUT_PATH = Path("outputs/guiden_reddit_mvp/guiden_reddit_story_seed.xlsx")


HEADERS = [
    "story_id",
    "source_platform",
    "source_url",
    "subreddit",
    "source_kind",
    "post_title",
    "curated_summary",
    "story_retelling_for_mvp",
    "struggle_tags",
    "what_helped",
    "support_type",
    "timeframe",
    "tone",
    "risk_flags",
    "pii_removed",
    "moderation_status",
    "prototype_notes",
]


ROWS = [
    {
        "story_id": "reddit_001",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/Anxiety/comments/zxv8wy",
        "subreddit": "r/Anxiety",
        "source_kind": "post",
        "post_title": "Success Story: I beat my anxiety. Mostly.",
        "curated_summary": "The poster describes years of anxiety that escalated after major life stressors into anxiety attacks, food fears, driving fears, and severe weight loss. They report improvement after therapy, psychiatric care, medication, walking, workouts, meditation, and celebrating small wins.",
        "struggle_tags": "anxiety, panic, driving anxiety, appetite, life transition",
        "what_helped": "therapy; psychiatrist; medication; daily walking; gym routine; meditation; small victories",
        "support_type": "professional, lifestyle, practical",
        "timeframe": "about 3-5 months for major turnaround",
        "tone": "hopeful, practical",
        "risk_flags": "medical details, weight loss",
        "pii_removed": "yes",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Strong match for users with panic plus life-transition stress.",
    },
    {
        "story_id": "reddit_002",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/PanicAttack/comments/1rwso2g/how_i_overcame_my_panic_when_i_was_convinced_i/",
        "subreddit": "r/PanicAttack",
        "source_kind": "post",
        "post_title": "How I overcame my panic when I was convinced I would die with this disorder 6 months ago",
        "curated_summary": "The poster says panic disorder worsened into agoraphobia, dropping work and school, and fear of attacks. Their partial hospitalization program helped most, and the story appears useful as a severe-panic recovery reference that emphasizes structured care and persistence.",
        "struggle_tags": "panic disorder, agoraphobia, depression, school, work",
        "what_helped": "partial hospitalization program; professional support; persistence; structured recovery",
        "support_type": "professional, emotional",
        "timeframe": "about 6-9 months",
        "tone": "hopeful, intense",
        "risk_flags": "severe panic, depression, medication discussion",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Use carefully because source includes severe symptoms and treatment details.",
    },
    {
        "story_id": "reddit_003",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/Anxiety/comments/1ds9hh3",
        "subreddit": "r/Anxiety",
        "source_kind": "post",
        "post_title": "It was my all breathing",
        "curated_summary": "The poster describes years of anxiety and panic with air hunger, then reports feeling better after practicing slow, light breathing rather than deep breathing. The main takeaway is that breathing style and CO2 tolerance can matter for some people.",
        "struggle_tags": "anxiety, panic, air hunger, breathing",
        "what_helped": "slow light breathing; breathing practice; hope framing",
        "support_type": "practical, lifestyle",
        "timeframe": "a few days reported",
        "tone": "encouraging, anecdotal",
        "risk_flags": "physiology claim needs disclaimer",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Good for matching breathing-related anxiety, but avoid presenting as medical advice.",
    },
    {
        "story_id": "reddit_004",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/PanicAttack/comments/1m6dhxx",
        "subreddit": "r/PanicAttack",
        "source_kind": "comment",
        "post_title": "Success stories needed",
        "curated_summary": "A commenter reports being off SSRIs and no longer seeing a therapist for months, while still using an anxiety book/audio resource during spikes. They emphasize confidence that they can live life without carrying emergency medication everywhere.",
        "struggle_tags": "panic attacks, anxiety, medication anxiety, confidence",
        "what_helped": "anxiety education resource; repeated listening; confidence building; coping practice",
        "support_type": "practical, emotional",
        "timeframe": "about 8 months",
        "tone": "reassuring",
        "risk_flags": "medication discontinuation mention",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Avoid implying users should stop medication; frame as one person's experience.",
    },
    {
        "story_id": "reddit_005",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/DecidingToBeBetter/comments/n5o9l9",
        "subreddit": "r/DecidingToBeBetter",
        "source_kind": "post_and_comments",
        "post_title": "How to break Depression Habits",
        "curated_summary": "The thread centers on rebuilding self-care after depression and anxiety. Helpful ideas include starting with one tiny habit, rewarding it, stacking new habits after it, using a habit tracker, and treating routine as a baseline to return to during slumps.",
        "struggle_tags": "depression, anxiety, self-care, routine, motivation",
        "what_helped": "one small promise; habit stacking; rewards; habit tracker; routine; therapy",
        "support_type": "practical, professional",
        "timeframe": "gradual",
        "tone": "practical, gentle",
        "risk_flags": "none",
        "pii_removed": "yes",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Strong general-purpose row for users who feel stuck after depression.",
    },
    {
        "story_id": "reddit_006",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/DecidingToBeBetter/comments/1r2fhy6/ive_tried_many_things_for_my_mental_health_and/",
        "subreddit": "r/DecidingToBeBetter",
        "source_kind": "comments",
        "post_title": "I’ve tried many things for my mental health and none of them have worked - what has worked for you?",
        "curated_summary": "Commenters suggest simple repeatable supports: routine, consistent meals, walks twice a day, creative projects, gratitude lists, music, writing, singing, sobriety, therapy, and appropriate medication review.",
        "struggle_tags": "depression, anxiety, routine, low mood, stuck",
        "what_helped": "routine; morning walks; gratitude list; music; writing; art; sobriety; therapy; medication review",
        "support_type": "lifestyle, practical, professional",
        "timeframe": "varied",
        "tone": "idea-rich, mixed",
        "risk_flags": "treatment suggestions need review",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Useful as a broad what-helped row, not a single testimony.",
    },
    {
        "story_id": "reddit_007",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/ptsd/comments/1qabck4/success_stories/",
        "subreddit": "r/ptsd",
        "source_kind": "comment",
        "post_title": "Success stories",
        "curated_summary": "A commenter with severe CPTSD says years of therapy, medication, a loving partner, friends, and sustained hard work helped them build a life they love, including living abroad and preparing for parenthood while accepting that healing is not linear.",
        "struggle_tags": "ptsd, cptsd, depression, relationships, long-term recovery",
        "what_helped": "therapy; medication; partner support; friends; self-compassion; asking for help",
        "support_type": "professional, community, emotional",
        "timeframe": "years",
        "tone": "hopeful, realistic",
        "risk_flags": "trauma content",
        "pii_removed": "yes",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Good row for trauma users who need long-term but hopeful examples.",
    },
    {
        "story_id": "reddit_008",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/ptsd/comments/1qabck4/success_stories/",
        "subreddit": "r/ptsd",
        "source_kind": "comment",
        "post_title": "Success stories",
        "curated_summary": "Another commenter says focusing on rediscovering the self helped, starting with a therapist-guided exercise of writing compassionately to a friend and then to themselves. The self-letter became a grounding tool for difficult moments.",
        "struggle_tags": "ptsd, self-worth, grounding, self-compassion",
        "what_helped": "therapy prompt; writing a letter to self; grounding; self-compassion",
        "support_type": "emotional, professional, practical",
        "timeframe": "not specified",
        "tone": "warm, reflective",
        "risk_flags": "trauma context",
        "pii_removed": "yes",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Useful for matching users who want self-compassion exercises.",
    },
    {
        "story_id": "reddit_009",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/CPTSD/comments/1o5t5cn/anyone_have_a_success_story_where_they_went_from/",
        "subreddit": "r/CPTSD",
        "source_kind": "comment",
        "post_title": "Anyone have a success story where they went from totally alone...",
        "curated_summary": "A commenter describes being cut off from family and friends, barely functioning, and financially unstable, then gradually finding work, building independence, and later recognizing they were the driving force behind recovery even though support mattered.",
        "struggle_tags": "cptsd, isolation, work, independence, financial stress",
        "what_helped": "work stability; self-directed recovery; therapy later; perseverance; support when available",
        "support_type": "practical, emotional, professional",
        "timeframe": "several years",
        "tone": "hard-won, hopeful",
        "risk_flags": "trauma, isolation, financial distress",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Good for users asking whether severe isolation can improve.",
    },
    {
        "story_id": "reddit_010",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/CPTSD_NSCommunity/comments/qsm2hn/cptsd_success_story/",
        "subreddit": "r/CPTSD_NSCommunity",
        "source_kind": "post",
        "post_title": "CPTSD Success Story",
        "curated_summary": "The poster says after months of tapering therapy frequency, their therapist suggested moving to as-needed sessions. They describe recovery as lifelong but feel they are finally living as their authentic self.",
        "struggle_tags": "cptsd, trauma recovery, therapy, identity",
        "what_helped": "long-term therapy; reducing session frequency gradually; recovery reading; authentic self-work",
        "support_type": "professional, emotional",
        "timeframe": "months to years",
        "tone": "reflective, hopeful",
        "risk_flags": "trauma context",
        "pii_removed": "yes",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Good concise therapy-progress row.",
    },
    {
        "story_id": "reddit_011",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/CPTSD/comments/o99bij",
        "subreddit": "r/CPTSD",
        "source_kind": "comment",
        "post_title": "Are there any success stories here?",
        "curated_summary": "A commenter says finding the CPTSD community helped them understand why earlier therapy attempts felt ineffective, leading to a breakthrough and later trauma recovery with a therapist.",
        "struggle_tags": "cptsd, trauma, therapy fit, understanding symptoms",
        "what_helped": "community identification; trauma-informed therapy; reframing past treatment",
        "support_type": "community, professional, emotional",
        "timeframe": "about 3 years before first therapist",
        "tone": "validating",
        "risk_flags": "trauma context",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Useful for users who feel previous therapy did not work.",
    },
    {
        "story_id": "reddit_012",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/socialanxiety/comments/vbz2am",
        "subreddit": "r/socialanxiety",
        "source_kind": "thread",
        "post_title": "Share your success story!",
        "curated_summary": "The thread asks people to share social anxiety wins. It is useful as a source for examples of gradual exposure, small social steps, and community encouragement, though individual rows should be reviewed before public use.",
        "struggle_tags": "social anxiety, isolation, exposure, confidence",
        "what_helped": "small social steps; gradual exposure; community encouragement; therapy if accessible",
        "support_type": "community, practical",
        "timeframe": "varied",
        "tone": "encouraging",
        "risk_flags": "aggregate thread, needs individual review",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Use as candidate pool rather than final testimony row.",
    },
    {
        "story_id": "reddit_013",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/BipolarReddit/comments/1kjevww",
        "subreddit": "r/BipolarReddit",
        "source_kind": "comment",
        "post_title": "Share your success stories PLEASE",
        "curated_summary": "A commenter with a long bipolar/schizoaffective history describes less chaos after medication, stable periods lasting years, relationships, children, graduate degrees, a meaningful career, and the importance of sleep, food, exercise, introspection, therapy, and medication stability.",
        "struggle_tags": "bipolar, schizoaffective, stability, relationships, career",
        "what_helped": "medication stability; therapy; sleep routine; exercise; food routine; introspection",
        "support_type": "professional, lifestyle, practical",
        "timeframe": "years",
        "tone": "realistic, hopeful",
        "risk_flags": "diagnosis and medication content",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Useful but requires careful medical-disclaimer handling.",
    },
    {
        "story_id": "reddit_014",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/BipolarReddit/comments/1637bmk",
        "subreddit": "r/BipolarReddit",
        "source_kind": "comment",
        "post_title": "Success stories?",
        "curated_summary": "A commenter says they have a long marriage, child, pets, home, school, and full-time work after finding a medication plan that keeps them fairly stable. They credit weekly therapy, a strict sleep schedule, and staying attentive to early episode signs.",
        "struggle_tags": "bipolar, stability, sleep, work, family",
        "what_helped": "medication plan; weekly therapy; strict sleep schedule; symptom awareness; routine",
        "support_type": "professional, lifestyle, practical",
        "timeframe": "long-term",
        "tone": "steady, reassuring",
        "risk_flags": "medication content, diagnosis content",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Good match for users seeking stability examples.",
    },
    {
        "story_id": "reddit_015",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/bipolar/comments/tuwgtj",
        "subreddit": "r/bipolar",
        "source_kind": "comment",
        "post_title": "Tell me your success stories!!",
        "curated_summary": "A commenter reports a year of stability after medication, weekly psychotherapy, nutrition, and exercise. They later returned to work and were accepted to graduate school while planning supports to maintain stability.",
        "struggle_tags": "bipolar, depression, school, work, stability",
        "what_helped": "medication; weekly psychotherapy; nutrition; exercise; disability supports; pacing",
        "support_type": "professional, lifestyle, practical",
        "timeframe": "about 6 months to stability, then 1 year stable",
        "tone": "optimistic, practical",
        "risk_flags": "medication content, diagnosis content",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Strong life-rebuilding row, but medical details should be generalized in app.",
    },
    {
        "story_id": "reddit_016",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/bipolar2/comments/1ewdvcr",
        "subreddit": "r/bipolar2",
        "source_kind": "comment",
        "post_title": "tell me a success story",
        "curated_summary": "A commenter says stability is supported by exercising daily, eating the same foods at the same times, and keeping consistent wake and sleep times. The row is useful as a simple routine-based stability example.",
        "struggle_tags": "bipolar2, stability, routine, sleep, exercise",
        "what_helped": "daily exercise; consistent meals; consistent sleep and wake times; routine",
        "support_type": "lifestyle, practical",
        "timeframe": "not specified",
        "tone": "concise, practical",
        "risk_flags": "diagnosis content",
        "pii_removed": "yes",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Clean example for routine matching.",
    },
    {
        "story_id": "reddit_017",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/ADHD/comments/1b8exir/how_has_medication_changed_your_life/",
        "subreddit": "r/ADHD",
        "source_kind": "thread",
        "post_title": "How has medication changed your life?",
        "curated_summary": "A commenter describes ADHD medication as quieting their brain enough to manage daily life, with anxiety and depression improving significantly and a morning routine becoming possible.",
        "struggle_tags": "adhd, anxiety, depression, routine, lateness",
        "what_helped": "ADHD diagnosis; medication; morning routine; capacity to manage tasks",
        "support_type": "professional, practical",
        "timeframe": "about 1.5 months reported",
        "tone": "relieved, practical",
        "risk_flags": "medication content",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Useful for comorbid anxiety/depression where ADHD may be relevant.",
    },
    {
        "story_id": "reddit_018",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/stopdrinking/comments/1bzft9g",
        "subreddit": "r/stopdrinking",
        "source_kind": "post_and_comments",
        "post_title": "People who successfully stopped - how?",
        "curated_summary": "The thread discusses alcohol spirals linked to anxiety, boredom, and depression. One recovery-oriented strategy mentioned is committing to a defined sober period while giving medication or treatment a chance, with community accountability.",
        "struggle_tags": "alcohol, anxiety, depression, boredom, habits",
        "what_helped": "defined sober commitment; medication support; community accountability; changing routines",
        "support_type": "community, professional, practical",
        "timeframe": "8-week commitment mentioned",
        "tone": "supportive, direct",
        "risk_flags": "substance use, treatment content",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Good for substance-related mental health matching; needs careful framing.",
    },
    {
        "story_id": "reddit_019",
        "source_platform": "Reddit",
        "source_url": "https://dd.reddit.com/r/stopdrinking/comments/1i5wogx/the_weight_is_falling_off_since_quitting_drinking/",
        "subreddit": "r/stopdrinking",
        "source_kind": "post",
        "post_title": "The weight is FALLING OFF since quitting drinking",
        "curated_summary": "The poster describes positive physical and emotional changes after stopping nightly wine use, including weight loss, hydration, and improved skin. For Guiden, the useful angle is how removing alcohol can improve wellbeing for some people.",
        "struggle_tags": "alcohol, habit change, physical health, wellbeing",
        "what_helped": "quitting alcohol; hydration; medical support; habit replacement",
        "support_type": "lifestyle, professional",
        "timeframe": "not specified",
        "tone": "energized",
        "risk_flags": "substance use, weight-loss focus",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Consider excluding if app focus should stay mental-health recovery rather than weight.",
    },
    {
        "story_id": "reddit_020",
        "source_platform": "Reddit",
        "source_url": "https://us.reddit.com/r/B12_Deficiency/",
        "subreddit": "r/B12_Deficiency",
        "source_kind": "community_reference",
        "post_title": "B12 Deficiency success story references",
        "curated_summary": "The subreddit includes success-story discussions where people report anxiety/depression-like symptoms improving after identifying a medical contributor such as B12 deficiency. This should be used only as a prompt to consider medical evaluation, not as a diagnosis.",
        "struggle_tags": "depression, anxiety, medical evaluation, fatigue",
        "what_helped": "medical testing; identifying deficiency; clinician-guided supplementation",
        "support_type": "professional",
        "timeframe": "varied",
        "tone": "cautious, informative",
        "risk_flags": "medical causation claim, diagnosis risk",
        "pii_removed": "yes",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Useful as a safety reminder to include professional evaluation when symptoms persist.",
    },
]


README_ROWS = [
    ["Purpose", "Seed workbook for the Guiden MVP using Reddit-sourced, human-curated summaries rather than copied full post text."],
    ["Use in app", "Search over story_retelling_for_mvp, curated_summary, struggle_tags, what_helped, support_type, and tone."],
    ["Do not use", "Do not treat these rows as medical advice, model-training data, or production-ready licensed content."],
    ["Moderation", "Only rows marked approved_for_closed_mvp should be shown in closed prototype tests without another review pass."],
    ["Privacy", "Usernames and direct personal identifiers are intentionally omitted."],
    ["Next step", "Replace or supplement these rows with opt-in Guiden submissions as soon as possible."],
]


def build_story_retelling(row: dict[str, str]) -> str:
    tags = row["struggle_tags"]
    helped = row["what_helped"].replace(";", ",")
    timeframe = row["timeframe"]
    summary = row["curated_summary"]
    title = row["post_title"]

    return (
        f"Retold for MVP testing from the source titled '{title}'. This is not copied from the Reddit post; it is a fuller narrative version built from the curated notes. "
        f"The person is dealing with {tags}. {summary} "
        f"The story reads like someone reaching a point where the problem had stopped feeling like a temporary rough patch and had started affecting ordinary life. "
        f"They were not just trying to feel better in the abstract; they were trying to get back pieces of daily functioning, trust their own body or mind again, and rebuild a sense that the future was not permanently closed off. "
        f"What makes the story useful for Guiden is the middle section: the person does not describe one magic solution. Instead, the recovery path is a combination of experiments, repeated supports, and small proof points. "
        f"The practical supports recorded for this row are: {helped}. "
        f"Those details let a user see both the emotional arc and the concrete actions: what changed, what was repeated, what kind of help entered the picture, and what someone might try discussing with a professional or trusted support person. "
        f"The timeframe is {timeframe}, so the app should present the story as gradual and personal rather than as a guaranteed quick fix. "
        f"The strongest takeaway is that the person found a way to move from feeling controlled by the issue toward having more choices again. "
        f"For the prototype, this story should be displayed with the source URL, the peer-support disclaimer, and the separate 'what helped' list so testers can react to both the narrative and the actionable parts."
    )


def autosize(sheet, widths: dict[str, int]) -> None:
    for idx, header in enumerate(HEADERS, start=1):
        width = widths.get(header, 18)
        sheet.column_dimensions[get_column_letter(idx)].width = width


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Stories"

    ws.append(HEADERS)
    for row in ROWS:
        ws.append([row.get(header, "") for header in HEADERS])

    header_fill = PatternFill("solid", fgColor="1F3A37")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 34

    autosize(
        ws,
        {
            "story_id": 14,
            "source_platform": 16,
            "source_url": 48,
            "subreddit": 22,
            "source_kind": 20,
            "post_title": 42,
            "curated_summary": 58,
            "story_retelling_for_mvp": 86,
            "struggle_tags": 34,
            "what_helped": 48,
            "support_type": 28,
            "timeframe": 24,
            "tone": 20,
            "risk_flags": 34,
            "pii_removed": 14,
            "moderation_status": 24,
            "prototype_notes": 48,
        },
    )

    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 168

    for row_idx, row in enumerate(ROWS, start=2):
        story_col_idx = HEADERS.index("story_retelling_for_mvp") + 1
        ws.cell(row=row_idx, column=story_col_idx).value = build_story_retelling(row)

    table = Table(displayName="GuidenStorySeed", ref=f"A1:{get_column_letter(len(HEADERS))}{ws.max_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)

    tag_ws = wb.create_sheet("Tag Summary")
    tag_ws.append(["tag", "story_count"])
    tag_counts: dict[str, int] = {}
    for row in ROWS:
        for tag in row["struggle_tags"].split(","):
            cleaned = tag.strip()
            tag_counts[cleaned] = tag_counts.get(cleaned, 0) + 1
    for tag, count in sorted(tag_counts.items(), key=lambda item: (-item[1], item[0])):
        tag_ws.append([tag, count])

    for cell in tag_ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    tag_ws.freeze_panes = "A2"
    tag_ws.column_dimensions["A"].width = 30
    tag_ws.column_dimensions["B"].width = 14
    tag_ws.sheet_view.showGridLines = False
    tag_table = Table(displayName="GuidenTagSummary", ref=f"A1:B{tag_ws.max_row}")
    tag_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True)
    tag_ws.add_table(tag_table)

    readme_ws = wb.create_sheet("README")
    readme_ws.append(["Field", "Guidance"])
    for row in README_ROWS:
        readme_ws.append(row)
    for cell in readme_ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    readme_ws.column_dimensions["A"].width = 18
    readme_ws.column_dimensions["B"].width = 110
    for row in readme_ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    readme_ws.sheet_view.showGridLines = False

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(OUTPUT_PATH.resolve())


if __name__ == "__main__":
    main()
