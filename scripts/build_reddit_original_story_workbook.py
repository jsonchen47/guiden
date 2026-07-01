from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


OUTPUT_PATH = Path("outputs/guiden_reddit_mvp/guiden_reddit_original_stories.xlsx")


HEADERS = [
    "story_id",
    "source_platform",
    "source_url",
    "subreddit",
    "source_kind",
    "post_title",
    "original_story",
    "struggle_tags",
    "what_helped",
    "support_type",
    "timeframe",
    "tone",
    "risk_flags",
    "retrieval_status",
    "moderation_status",
    "prototype_notes",
]


ROWS = [
    {
        "story_id": "reddit_original_001",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/anxietysuccess/comments/7xp7ut/i_overcame_anxiety_and_so_can_you_success_story/",
        "subreddit": "r/anxietysuccess",
        "source_kind": "post",
        "post_title": "I overcame Anxiety and so can you (success story)",
        "original_story": """Hello guys and gals, ive been wanting to write my story and thoughts for quite some time now, in hopes that it could be useful for someone. Its something ive been wanting to do but never quite felt it was the right time as I had been so focused on my own progress, that is to say...I always thought I needed to somehow be in a perfect place in life and feeling perfect before I could write something like this.

The obvious truth however is that I dont believe it is possible to ever feel perfect or have what is considered a perfect life, none the less, I feel I can safely say I have for the most part, overcome anxiety and I would like to share my story with all of you while also going more in depth for how you can also so if you have the time to read through what I have to say, I truly believe you will get something out of it and that these methods can also work for you so get a tea or something nice to drink and lets do this together.

My Story:
I will try to be brief mostly, and more detailed with what I think is relevant but I think its important that I paint close to the entire picture of how my past has been and up to this point because I think many of you could potentially relate. So as of writing this, im turning 34 today, yeah today is my birthday and im writing this because what a better way to celebrate than digging up the past haha. Alright sorry...so I was a product of the 90s meaning...at 8 years old doctors put me on Ritalin.

I scored highly in intelligence tests but couldnt be bothered to focus in schools. It mostly worked and things went on etc.

Fast forward to like 15 or so, I was a rebellious teen (who wasnt?) I had some issues at home and my mom had trouble dealing with it so eventually a Psychiatrist put me on Prozac for depression.

If anyone can recall, doctors were very eager to prescribe any type of drug like this even to people under the age of 18...Regardless, I never considered myself a depressed kid or teen, I had some anger issues but that was mostly due to no one listening to my wants or needs and forcing me to try and be someone I wasnt. Anyhow, I trusted the doctors and my mom and they told me I needed to be on these drugs for the rest of my life. How does this all relate to anxiety you ask? We will get to that but first..

Fast forward to 25...id been swallowing pills for 10 years now without a 2nd thought of it until somehow one night I realized, why am I taking these? I never feel depressed, I have tons of friends and a great career as an Artist etc.

After this realisation I decided I did not want to spend my life taking pharmaceutical drugs so I tapered off of Effexor which I had been on for like 6 years.

What happened after this was the start of the hell for me. I suffered very extreme Withdrawel symptoms including intense panic attacks. The first one hit me and I shook it off, but the 2nd one sent me to the hospital.

It happened while at work due to my nervous system being in such a state of havoc after getting off drugs that eventually I found another doctor who again gave me the info that I needed the drugs so I was put back on another antidepressant which helped but this time I now had an anxiety disorder that would not go away.

So how did this start to effect my life? I had trouble leaving the house but the pills helped, i still would get panic attacks but it was manageable. I will go more into anxiety later but first I will continue with my story.

I knew I needed to be off these drugs, I also know that I needed to deal with the anxiety so after a few years I found a great Psychiatrist who I will refer from this point forward as "Coach". After seeing him he evaluated me and agreed that I was not someone that needed to be medicated although he agreed I had now developed an anxiety disorder so with his help I began a very slow tapering of my drugs while seeing him every week. I tapered off for a year before getting off.

During our time together we also focused on Meditation and Mindfulness, things like acceptance and also trying to sort out all the other neurosis.

The real hell began for me after I finally got off the drugs. There are no words in the English language that can describe the extreme suffering I went through and in all honesty, I consider myself a strong person and I am still amazed I survived with my life from what happened next.

Withdrawels began, the initial few weeks were a ramp up to the worst of it but I had an entire host of issues including a hospital visit for Trachardia (Extreme heart beat) non serious, extreme fatigue, twitching muscles, insomnia, and the most debilitating panic attacks ive ever experienced.

These panic attacks were a result of my nervous system completely fried from these drugs and being cut off from them, they would be so intense that at the same time of having them, I would have intense suicidal feelings come over me. They were so intense that after one would end, I would sleep for 14 hours. I was sick...so sick I looked white and dead, so sick my family was terrified. I wanted to be checked into the mental hospital but thankfully Coach advised against in.

He assured me that I was not losing my mind, which it felt like I was going crazy, and he assured me I was healthy and that we could do this...he gave me a prescription for Xanax and I used that to ride out the worst of the storm.

For 6 months I essentally was disabled. I spent my days in such pain and agony from drug withdrawls and anxiety that all I could do was lay in bed and pray and take Xanax and hope for sleep. It was not possible to do anything other than this. During this time as well, my entire childhood hit me as if it had been supressed all the years by drugs so I had all this emotional stuff hit me as well.

I wanted so bad to kill myself every single day but something in me kept fighting, I wanted to know how life could be, I made it a point to allow this to transform me into more than I was.

Around 6 months I was able to walk around my block, and accept visits from my best friend. It sounds pretty pathetic but thats where I was in life. Luckily during all of this I was able to live at my moms and work from home which I managed to keep up thanks to the Xanax. It was also around this time I was basically not taking anymore xanax unless very extreme. Anxiety was constant though and I was now in a place where I could deal with the anxiety once and for all.

I kept seeing coach every week, he kept giving me the motivation to continue and I kept working on myself, my discipline was high as I wanted this more than anything. Still though it was hard to function, I couldnt go out in public or leave my house much, my emotions were distorted, I couldnt handle action movies or violence anymore, or anything sexual, I could cry easily at seeing a sad film for example..>I had a lot of intense emotions.

Fast forward to a year and I was starting to force myself to go out once a week to see my friends...it was very hard but I forced it and it helped.

I wont go into a ton of details but I suffered for years, although I kept making progress. I started that journey around 29 or so and I am now writing this, 34 from Sweden. (I had spent my entire life in the US on the West Coast)

So am I perfect now? Not by a long shot, I still have some odd issues from that ordeal but I survived and how can I claim that I overcame anxiety? Well here is the proof:

From then till now, I couldnt drive a car or leave my house, I was afraid to travel, I had horrible anxiety, panic, physical symptoms all the time, derealization, constantly worried about my health..I had no life and also the old me was a bit self centered, a good guy but selfish a bit and un aware of other peoples suffering.

Here is my life since then:
I have minimal symptoms of anxiety, I backpacked Europe by myself, I met my current gf while traveling through Vienna Austria and ended up living there.. I got in great shape, we have both moved to Sweden together and I work in an office again. I can go anywhere and do anything without anxiety, i rarely feel it and if it comes up at times again, it hardly phases me, and perhaps most of all, through my suffering i became kinder and more loving.

I have complete compassion now for others and can easily sense when someone is not feeling well.""",
        "struggle_tags": "anxiety, panic attacks, derealization, withdrawal, depression",
        "what_helped": "psychiatrist support; slow tapering; meditation; mindfulness; acceptance; exposure; discipline; walking; friends and family",
        "support_type": "professional, practical, emotional, lifestyle",
        "timeframe": "years",
        "tone": "long-form, intense, hopeful",
        "risk_flags": "suicidal thoughts, medication withdrawal, benzodiazepine mention",
        "retrieval_status": "visible_via_search_and_direct_open",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Strong long-form anxiety recovery narrative. Includes medical/withdrawal details that need app disclaimers.",
    },
    {
        "story_id": "reddit_original_002",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/PanicAttack/comments/1rwso2g/how_i_overcame_my_panic_when_i_was_convinced_i/",
        "subreddit": "r/PanicAttack",
        "source_kind": "post",
        "post_title": "How I overcame my panic when I was convinced I would die with this disorder 6 months ago",
        "original_story": """Hi everyone! Just wanted to share what I consider to be a success story! Never thought I would have the opportunity to share one of my own.

I (20f) have had panic disorder since I was 17 years old. While I was able to manage it initially, about 9 months ago I had the worst flare up of my life. I couldn’t drive, I developed agoraphobia and couldn’t leave my house. I quit my job, dropped out of my graduate program. I cancelled a dream vacation I’d been planning. I was so depressed, and felt like with every panic attack I had, I was falling deeper and deeper into a hole I couldn’t crawl out of.

I tried everything. I was on daily benzos, propranolol, hydroxyzine, you name it. I took mood stabilizers, antidepressants, antipsychotics. It felt like NOTHING worked. I even committed myself to a partial hospitalization program (which did actually help the most, and I do recommend, but was not a “fix”). It got to the point that it honestly didn’t matter whether or not I had a panic attack, because I was spending my entire day worrying about it either way.

That last sentence is quite literally the thing that gave me my life back (for the most part—a wonderful support system, a few medication tweaks, and therapy were all NECESSITIES for myself). I had spent thousands of dollars and years and medications just looking for anything to stop the panic so that I could keep living my life. It wasn’t until my friends were planning a trip to Florida and I was sobbing in my bed about not being able to go that it hit me.

The panic attacks weren’t killing me or stopping me. The fear around them was.

I know everybody says acceptance is necessary, and it is by no means the easiest. In fact, more than any other treatment I tried, it was by far the hardest. I kicked and screamed and pulled my hair in anger and fury about my situation. But nothing I did was going to take away panic attacks, and no matter how badly I screamed about how unfair it was, it wasn’t going to change anything.

Progress was slow at first, but once it picks up it absolutely flies by. I stopped trying to make my panic less intense, and stopped worrying about “ruining the moment” for others. I let myself have a panic attack in the store, on the plane, in the car. I let myself have a panic attack anywhere I wanted. I treated it like a sneeze—another bodily function. And you know how I got my freedom back? It wasn’t by eliminating panic attacks. It was by not being scared to have one wherever I went.

That I COULD go to Florida and have a panic attack. I COULD drive and have a panic attack.

And honestly, I no longer give a FUCK what anyone says about me or if I get weird looks or if I have to pull over.

Granted, once you start being unafraid to have them, the panic attacks also do stop too. But that isn’t the point. The point is that I no longer fear them, I view them the same way I view a sneeze. And with that, I have bought my entire life back.

If anyone out there is struggling or has been in a place like I was, don’t be afraid to reach out. Know that it is NOT like this forever, you WILL reach a point of acceptance one day, and you CAN do it. I will most definitely have panic attacks at some point in my life again, but I know the next time one comes I won’t be terrified anymore. And that’s the kind of peace I never thought I’d ever see in my life.

(Apologies for the long post, just wanted to share my story).
All questions and comments are welcome!""",
        "struggle_tags": "panic disorder, agoraphobia, depression, school, work, fear of panic",
        "what_helped": "partial hospitalization; medication tweaks; therapy; support system; acceptance; exposure",
        "support_type": "professional, emotional, practical",
        "timeframe": "about 9 months",
        "tone": "direct, hopeful, practical",
        "risk_flags": "medication list, severe panic, depression",
        "retrieval_status": "visible_via_direct_open",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Excellent Guiden-style testimony about fear of panic versus panic itself.",
    },
    {
        "story_id": "reddit_original_003",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/PanicAttack/comments/1pi6cc1/i_am_100_recovered_from_panic_disorder_after_2/",
        "subreddit": "r/PanicAttack",
        "source_kind": "post_and_comments",
        "post_title": "I am 100% recovered from Panic Disorder after 2 years",
        "original_story": """I was convinced I was going crazy or slipping out of reality. I was trapped in a constant state of high-alert hyper-awareness. I felt like a prisoner in a body I couldn't even feel because I was just "floating" around most of the time due to severe derealization. It got to the point where I couldn't leave the house, and was even terrified to go to the bathroom or take a shower. My symptoms were constant and included basically all the symptoms including these:

-Severe Depersonalization & Derealization (DP/DR) -Terrifying existential thoughts -Dizziness / Vertigo / "Boat rocking" sensation -Shakiness, numbness, and restlessness -Racing heart and dry mouth

I lost 2 years of my life and missed a full year of school. But I got out. I didn't recover through distraction, breathing exercises, or fighting the feelings. I recovered through Acceptance.

Ask me anything about how acceptance actually works, my experience with meds vs. natural recovery, or specific symptoms. I want to prove to you that you aren't stuck like this forever.

Selected answer from the poster:
There is no “doing” in acceptance; there is just letting it be and doing nothing to fix the feeling. By trying to intensely focus on it, you are likely not being okay with it—that is not acceptance, that is resistance. Anxiety pulls you in two directions that are polar opposites, meaning acceptance and resistance cannot exist in the same space. True acceptance is quieting the ego, which is the force driving your fear by desperately trying to control everything for survival.

To stop the frantic struggle, you must let go of that control. You can say to yourself: "I don't care if I have a panic attack. Really, I actually don't care. Anxiety, I accept you" And act like it.

Selected answer from the poster:
Acceptance means actively choosing to do nothing to "fix" the panic, allowing the wave of fear to pass through you instead of running from it. It requires you to stop "feeding the barking dog” by dropping all safety behaviors—such as checking your pulse or calling a friend—because these actions only convince your brain that the danger is real. Much like falling into quicksand, the more you struggle and fight for control, the faster you sink; to float, you must physically relax into the tension.

You shift your mindset from a victim to a scientist, observing the symptoms as harmless data rather than catastrophic threats in a completely non judgemental way. Ultimately, acceptance is quieting your ego's desperate need for control and genuinely telling the anxiety, "I don't care if I have a panic attack, I accept you" and you acting like it.""",
        "struggle_tags": "panic disorder, derealization, depersonalization, agoraphobia, school avoidance",
        "what_helped": "acceptance; stopping safety behaviors; observing symptoms; letting panic pass",
        "support_type": "practical, emotional",
        "timeframe": "about 2 years",
        "tone": "educational, hopeful",
        "risk_flags": "medical/therapy opinion, severe symptoms",
        "retrieval_status": "visible_via_direct_open",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Includes original post plus poster Q&A answers because the post is an AMA.",
    },
    {
        "story_id": "reddit_original_004",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/CPTSD_NSCommunity/comments/qsm2hn/cptsd_success_story/",
        "subreddit": "r/CPTSD_NSCommunity",
        "source_kind": "post",
        "post_title": "CPTSD Success Story",
        "original_story": """This is a CPTSD success story.

Yesterday--after a few months of tapering session frequency--my therapist suggested I was ready to switch to "as needed" therapy. I know that recovery is a lifelong journey, but it really feels like I'm finally "living" as my authentic self.

Here's how I got here:

1. Misdiagnosed with body dysmorphia and treated with CBT for 5 years. The CBT tools are great, but negative self concept is just one symptom of CPTSD. I was still struggling.
2. A friend suggested I might be borderline (hey thanks, friend ), so I looked into that. Therapist disagreed with that diagnosis, but in learning about BPD I discovered overlapping symptoms with CPTSD.
3. Read Pete Walker's Complex PTSD: From Surviving to Thriving. WOW, what an eye opener. I started using these tools right away to escape emotional flashbacks and shrink my inner critic. This was absolutely transformative and I recommend it often.
4. Read The Body Keeps the Score. Also great, but I think Pete's book was more relevant to my specific trauma.
5. I kept notes on my emotional flashbacks to better understand my triggers and--hopefully--learn what my childhood trauma was. But I was hitting a wall with what my current therapist could help me with. Per the recommendation in Pete Walker's book, I looked for a therapist that specialized in EMDR.
6. Found EMDR therapists in my area, interviewed two, started with one. We had a few talk-therapy sessions before going into EMDR. We did the binaural buzzer thing. I was skeptical, but in a few minutes I was revisiting long-forgotten memories and reprocessing them from the perspective of an adult. This is really incredible stuff.
7. As the year draws to an end, I'm running out of stuff to work on with my therapist and we're switching to "as needed."

The therapies suggested by Pete Walker really are life changing. I am finally getting to experience myself. I'm learning what I like and dislike. I'm making genuine friendships without any weird attachment issues. I have to remember to be patient with myself, but there is so much living to catch up on!

Learning about trauma has been really eye opening. Once you learn the tools and work on yourself a bit, you start to see trauma expressed by other people. Recovery has given me a profound empathy for other survivors, and a unique ability to speak to their pain.

I hope my success can inspire those who have not started their journey yet. I know it seems daunting, but it's so rewarding.""",
        "struggle_tags": "cptsd, trauma, emotional flashbacks, inner critic, therapy fit",
        "what_helped": "CPTSD education; Pete Walker book; notes on triggers; EMDR; trauma-informed therapy; tapering sessions",
        "support_type": "professional, educational, emotional",
        "timeframe": "years",
        "tone": "clear, reflective, hopeful",
        "risk_flags": "trauma context, therapy modality claims",
        "retrieval_status": "visible_via_direct_open",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "High-quality structured story with a concrete recovery path.",
    },
    {
        "story_id": "reddit_original_005",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/B12_Deficiency/comments/1ldcd2c/my_success_story/",
        "subreddit": "r/B12_Deficiency",
        "source_kind": "post",
        "post_title": "My success story",
        "original_story": """Hi, all. I’ve previously shared a bit of my story in this group before but thought I’d give an update and hopefully give some of y’all hope that it does get better.

When I think back to how I felt for the past few years, especially a year ago when my B12 reached 164 pg/mL vs how I feel now, it’s a night and day difference. Even before then, when my B12 wasn’t quite as bad (about 340 in 2022) I still felt quite bad compared to now.
I used to be so drained of energy every single day. I would come home from work and immediately get into bed, too exhausted both mentally and physically to do much else. I had daily headaches that had at least a moderate intensity but on some days were quite intense. I was always dizzy and felt weak.
The neurological symptoms became apparent a few months before I discovered my deficiency, and manifested as pins and needles mainly in my hands, ringing in my ears, snow in my visual field, and feeling very off balance. I also had severe depression, anxiety, and worsening ADHD. At its worst, I felt delirious at times, like I was starting to lose contact with reality.
My ability to function declined over time but reached a point of being unable to function shortly before a suicide attempt in October of 2024. This was preceded by poor performance and attendance at work, made even worse by severe sleep deprivation and a very low appetite. It was at a psychiatric hospital that my B12 was tested for the first time and that began my path to recovery from all of this.
I got weekly B12 injections for about two months, then I switched to taking a 5000 ug B12 supplement daily. My symptoms improved precipitously, especially the neuropsychiatric symptoms. But I was disappointed a bit that I didn’t get a complete resolution of my symptoms. I saw improvements in energy and fatigue, but there was still a major problem with these symptoms despite the B12 therapy.
I noticed that my hair continued to fall out in high amounts as it had before, and asked to get an iron panel and discovered the other source of my symptoms was likely iron deficiency. My ferritin was 6 ng/mL.
Fortunately, I was referred to hematology and gastroenterology. The hematologist quickly got me scheduled to receive two iron infusions of faraheme. After two weeks, the difference was so subtle that it really discouraged me and made me question if I’d ever fully recover. But after a month, especially after the two month mark, my symptoms improved to such a great extent that I’m still blown away by it.
I can finally say that I feel alive. I feel the best that I have felt in years. I used to be a very on and off runner, trying to run but never being able to run more than once a week and I could barely do a mile or two on a treadmill. Now, I run on a trail about every other day, run about 3 miles and much of it is uphill. I don’t even feel nearly as exhausted as I did after exercise before.
And I feel so strong and powerful during my runs, like my body is finally able to produce energy and be fully oxygenated. I’m doing great at work—my boss says I’ve made impressive progress over the past 6 months (coinciding with the start of my B12 therapy). I make far fewer mistakes and can get so much more done with so much less effort. I can think clearly and my brain isn’t so foggy anymore.
My PCP made the comment that there has been a stark difference in my presentation a year ago vs now, as a year ago I was depressed, apathetic, had a more flat affect, but now, I was smiling and laughing just in regular conversation.
On a run I got back from recently, I cried happy tears. I’m still in a state of disbelief that it’s even possible to feel this good. I forgot what it felt like to have energy to do the things I enjoy and to feel great while I was doing them. I didn’t know I could just live without random spells of depression and anxiety consuming me. I feel so hopeful for the future and have gained my confidence in myself back.
I got my life back, and I couldn’t be happier that life gave me another chance and that I have access to the healthcare that enable me to get here.
If you’re feeling hopeless, don’t give up. I know how frustrating it can be when you end up with more questions than answers. I’m still kind of in that boat even now with the discovery I have a stomach ulcer and antral erosive gastritis that has no clear cause. But you can't give up on yourself. Advocate for yourself and do whatever it takes to save yourself, you will be so grateful that you did.""",
        "struggle_tags": "depression, anxiety, ADHD, fatigue, medical evaluation, B12 deficiency",
        "what_helped": "psychiatric hospitalization testing; B12 injections; supplements; iron panel; hematology; gastroenterology; iron infusions; running",
        "support_type": "professional, medical, lifestyle",
        "timeframe": "months to years",
        "tone": "intense, hopeful, medical",
        "risk_flags": "suicide attempt mention, medical treatment details, supplementation details",
        "retrieval_status": "visible_via_direct_open",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Useful as a medical-evaluation story, but should never be shown as diagnostic advice.",
    },
    {
        "story_id": "reddit_original_006",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/depression/comments/1g694m7",
        "subreddit": "r/depression",
        "source_kind": "post",
        "post_title": "it does get better (success story)",
        "original_story": """i guess i’m a success story. as of right now, at least. i was diagnosed earlier this year and i’ve decided to take the medicated route along with therapy. along with depression i also have generalized anxiety, still working on figuring out the triggers for that one.

it took me a hot minute to get to where i am now. i had written multiple goodbye letters, plotted out years of my death if I wasn’t successful by then, it wasn’t until i was spiraling in front of my grandma that it really hit me that i was sick. i was about to take a 3 hour drive home instead of staying in my hometown for the night to get a mental health assessment the next morning when she used her Big Girl Grandma Voice on me for the first time in almost 15 years and told me I was staying and that was that.

she told me my actions were worrying everyone around me and the fact that i had anxiety just from getting the test proves that i needed to take it. so I took it. I got diagnosed after further questioning and my life has changed since then.

there was about a 3 month adjustment period but i’m a brand new person it seems! i’m doing my dream job by being a cosmetology school student and i’m taking in literal strangers and working on their hair. my confidence in myself has gone through the roof. i even wear makeup now and i like experimenting with my hair. it may sound vain to some but i hated myself for YEARS. you can’t take loving myself away from me any more.

please get help. live another day. if not for anyone else then for yourself. you deserve to live.""",
        "struggle_tags": "depression, generalized anxiety, self-worth, career, diagnosis",
        "what_helped": "family intervention; mental health assessment; diagnosis; medication; therapy; cosmetology school; self-expression",
        "support_type": "professional, family, emotional",
        "timeframe": "about 3 months after treatment adjustment",
        "tone": "urgent, hopeful",
        "risk_flags": "suicidal ideation, goodbye letters",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Strong depression recovery row but contains suicide references.",
    },
    {
        "story_id": "reddit_original_007",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/mentalhealth/comments/15gf5pl",
        "subreddit": "r/mentalhealth",
        "source_kind": "post",
        "post_title": "Positive success stories regarding work",
        "original_story": """M25. Have been struggling with depression for 6 years that has been getting worse with time, but kind of stabilized for the past period. Diagnosed with multiple disorders (major depression, quiet borderline, bipolar, an unspecified personality disorder, which i find most descriptive). Have tried many medications and a ton of therapists, and nothing seems to make a huge difference.

My biggest practical struggle is work. I am doing a PhD, I know I am smart and have potential, but with consistent extreme lack of desire and energy, everyday is a huge ordeal. I know that people here will understand without me explaining.

I am not writing to vent. I am writing to request from the community here to share some success/positive stories or experiences. I always see comments that empathize with OP, which is really valuable. But here, I want people older than me or that spent more time with the struggle to tell me there is hope by actually telling me some of their stories.

I have a huge personal barrier against therapy, so my alternative is self-reflection and discussing with friends, which is extremey helpful. Here’s my own mini success story about this from discussions and reflection:

I realized that I have to allow myself to accept that I genuinely have it more difficult than the average human. I have extreme expectations (like working 30 hours a week) that are physically impossible for me, and that always fail, hence inducing extreme guilt and disappointment that exacerbates my depression. I always try different work strategies to see what will work. My latest thing is working in “sessions”, and tracking the number of sessions (or total hours of work) per day. After a year, i discovered that even in my best state, I work an average of 20 hours a week. That’s my limit and that’s what i should aim for. For 6 months now, i have been struggling to get over 13-14 hours a week, so my current aim is to design low-expectation schedules and routines that slowly take me from 13 hours to 20 hours the end of 2023.

These routines take into account “fail days” that I take off due to being actually mentally impaired. But since they are planned, I feel no guilt, and hence take them happily and wait for the next day where I have more energy for a new attempt.

TL;DR: What’s your success story about being able to survive in practical life with depression? What helped u most? Therapy, meds, family, friends of self-reflection?""",
        "struggle_tags": "depression, work, PhD, expectations, low energy, guilt",
        "what_helped": "self-reflection; discussing with friends; tracking work sessions; realistic limits; low-expectation schedules; planned fail days",
        "support_type": "practical, emotional",
        "timeframe": "about 6-12 months of experimenting",
        "tone": "practical, reflective",
        "risk_flags": "diagnosis labels, work impairment",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Not a classic recovery story, but very useful for practical-functioning matching.",
    },
    {
        "story_id": "reddit_original_008",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/socialanxiety/comments/14v0q7r",
        "subreddit": "r/socialanxiety",
        "source_kind": "post",
        "post_title": "Finally tackling social anxiety at 25 years, after tens of failed attempts",
        "original_story": """Well, I had always been a regular here when I was at my lowest, so I thought it was only fair to give back to this amazing community, this is not intended as advice, but rather to tell a "success" story of sorts, feel free to ask anything in the comments or via private chat.

I rushed to find the earliest therapist, got lucky and started therapy only 1 week later, fully disclosing everything and being 100% ready for the hard path ahead, after that session I tried to recover contact with 3 friendships I had let die in the past, 1 rejected me, 1 was unavailable, and the last one was very happy about it, so now I have a friend I can talk to freely and hopefully go out and have fun.

That afternoon I posted here on Reddit in search of friends, and the process was draining, but I can confidently say now, that after 3 weeks since that post and around 20 conversations, I have made two very good friends from it all, I remain in contact with everyone that showed interest, and I am loving every and all interactions, I am actually having so much fun socializing.""",
        "struggle_tags": "social anxiety, isolation, friendship, therapy, exposure",
        "what_helped": "starting therapy quickly; full disclosure; reconnecting with old friends; online friendship attempts; repeated conversations",
        "support_type": "professional, community, practical",
        "timeframe": "about 3 weeks for early progress",
        "tone": "encouraging, early-progress",
        "risk_flags": "private chat mention",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Good early-win social anxiety row.",
    },
    {
        "story_id": "reddit_original_009",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/bipolar/comments/1kyl0g0",
        "subreddit": "r/bipolar",
        "source_kind": "comments",
        "post_title": "What is your “success story”?",
        "original_story": """Was diagnosed at 18, currently 26. I have been in a stable relationship for 8 years and married for 2. I experience mild auditory hallucinations that don’t cause issues for me. I haven’t had a manic or depressed episode in a year. I’m a happy person living with more joy and stability than a lot of neurotypicals. Medication, therapy, and lifestyle changes (sleep habits, stress levels) can work. You can be mentally ill and happy with your life simultaneously.

Another commenter:
Got diagnosed type 2 fall of my freshman year of college. I went through a life changing manic episode and moved back in with my parents. I didn’t think I’d be capable of finishing college, but I transferred somewhere new after my break and am now a proud college graduate and getting my masters! GTD (Getting Things Done) was created by Allen post hospitalization in a psych ward, and while most people think of it as a productivity system, I think it is also a mood regulation system and could be a big help in keeping from feeling overwhelmed during college.

Another commenter:
I raised an amazing son on my own. I got my bachelor of education and taught for 21 years. I got sober and have maintained sobriety for 7plus years. I have had a lovely and successful relationship for the last five years. I have maintained great relationship with my familyand friends despite the ups and downs of this illness. I am happy and looking forward to the second half of my life.

Another commenter:
My success story is that after years of following someone else's path and their vision for me, I've finally started living in friendship with my authentic self. This required well-chosen medications, various therapies, being surrounded by love and care, and a lot of inner work. I feel fortunate because life hasn't been kind to me in the past.""",
        "struggle_tags": "bipolar, stability, relationships, school, work, sobriety, authenticity",
        "what_helped": "medication; therapy; sleep habits; stress management; productivity system; sobriety; relationships; love and care",
        "support_type": "professional, lifestyle, community, practical",
        "timeframe": "years",
        "tone": "hopeful, multi-perspective",
        "risk_flags": "diagnosis content, medication content, psych ward mention",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Thread row combining multiple short success comments.",
    },
    {
        "story_id": "reddit_original_010",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/ADHD/comments/1b8exir/how_has_medication_changed_your_life/",
        "subreddit": "r/ADHD",
        "source_kind": "comments",
        "post_title": "How has medication changed your life?",
        "original_story": """Selected comment:
I have the non-stimulant meds and the initial impact was nausea and a clear mind. Within the first couple months, I felt so much peace because I had never experienced a decluttered head. After 3 months, I experienced happiness on a daily basis because I was less worried about small stuff, I could think more clearly, I had energy to hangout with friends or do hobbies that I normally would pass up in favor of sleep, and I was more longer experiencing lethargy.
Now that I’m 16 months medicated, I’m at peace and feel in control on my thoughts/feelings. The annual periods when I’d be most depressed no longer exist and I’m consistently optimistic about my wellbeing.

Selected comment:
I mean…everything was better, immediately. Life wasn’t so exhausting. I could do basic things without having to motivate myself to do it. Without even having to think about it, really. So much of my time was freed up. I was able to get two advanced degrees with high GPAs.

Not sure I could articulate myself better, though. I fixed this by splitting my stim dosages throughout the day so it didn’t just bomb my system all at once.

Some people are able to do fine without medication, but I was never one of them. I could make myself LOOK fine, but it just wasn’t the case. Medication made the looks match the reality.

Selected comment:
Social benefits to medication for me: I don’t interrupt people constantly anymore, I can pause and think about what I’m going to say first instead of blurting things out - the improvements with impulsivity have been of great social benefit. Also, I am not as late to social events anymore; I still have “time blindness”, but I’m much more on top of my tasks. Like, I can get ready to go out with less distraction / stay more on task and thus make it out the door faster.
Also, I FEEL like being social more and have more time to be social because I am not busy either 1) paralyzed, overwhelmed by all the things I’m procrastinating or 2) cramming in the stuff I procrastinated.""",
        "struggle_tags": "adhd, depression, anxiety, social functioning, executive function, lethargy",
        "what_helped": "ADHD medication under doctor care; clearer thinking; energy; social engagement; impulse control; task follow-through",
        "support_type": "professional, practical",
        "timeframe": "months to 16 months",
        "tone": "practical, relieved",
        "risk_flags": "medication anecdotes, subreddit medication warning",
        "retrieval_status": "visible_via_direct_open",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Comment compilation, not one story. Useful for ADHD-related search testing.",
    },
    {
        "story_id": "reddit_original_011",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/PanicAttack/comments/15gf5dp/i_dont_feel_the_same_after_my_panic_attack/",
        "subreddit": "r/PanicAttack",
        "source_kind": "comments",
        "post_title": "I don't feel the same after my Panic Attack",
        "original_story": """Selected comment:
I’ve been there and I do feel better now. It’s been about five months since I had my first panic attack, and although I had many more in the next two months, I haven’t had any for the last three months - just overwhelming anxiety and that feeling that something is off, that I’m not long for this world.

But that feeling has been getting better, and you’re not going to feel as awful as you do now. Stay confident that they are symptoms of anxiety and panic. It always gets better.

Follow-up from same commenter:
Thank you for asking! I’m doing pretty good but it took a while. I really leaned into the strategy of feeling shitty because of dread and anxiety, but basically just saying “dread and anxiety can’t hurt me” and accepting that I feel awful but not trying to fix it. The key to reducing anxiety ended up being not trying. Took a long time, a lot of really torturous nights where I broke down and took half an Ativan. But eventually it subsided.
It’s never fully gone but I feel basically back to where I was before all this.

Additional follow-up:
I’m a big fan of the acceptance technique, where you essentially say “this sensations sucks and I don’t like it but it is just a sensation and I’m not going to actively try to stop it”. If you want to know more about that, try the DARE method or listen to the Disordered podcast. Other than that, really time is what helped more than anything.

Additional update:
That aspect of it is almost gone, actually! Still more anxious than I used to be, and it’s a long road but I don’t feel any dissociation or general dread. Hope you’re doing okay, sorry for the delayed reply.""",
        "struggle_tags": "panic attack aftermath, dread, health anxiety, dissociation, acceptance",
        "what_helped": "time; acceptance technique; DARE method; therapy; medication as-needed mention",
        "support_type": "practical, emotional, professional",
        "timeframe": "months",
        "tone": "reassuring, realistic",
        "risk_flags": "benzodiazepine mention",
        "retrieval_status": "visible_via_direct_open",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Good row for users worried they will never feel normal after a panic attack.",
    },
    {
        "story_id": "reddit_original_012",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/mentalhealth/comments/tjps9b",
        "subreddit": "r/mentalhealth",
        "source_kind": "comment",
        "post_title": "people who have been through depression, or terrible life tell your success story or how you have changed",
        "original_story": """Used to be unemployed and addicted to pot. Was a recluse ignoring friends and family to stay high.

What worked best is to list all the problems. Then pick the most serious one and start by fixing that. I quit pot and it took months to get over it. Wait until you feel comfortable without that problem before moving onto the next. And never blame others for your problems, even if it’s their fault. Change yourself, you can’t change others and a victim mindset is the worst thing to have.

Now I have a good tech job, I’m physically fit, and have a good amount of friends I spend time with.""",
        "struggle_tags": "substance use, isolation, unemployment, depression-adjacent",
        "what_helped": "listing problems; choosing one problem first; quitting pot; waiting before tackling next issue; fitness; rebuilding social life",
        "support_type": "practical, lifestyle",
        "timeframe": "months",
        "tone": "direct, practical",
        "risk_flags": "substance use, moralizing language",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Useful but tone may need softening before public product use.",
    },
]

ADDITIONAL_ROWS = [
    {
        "story_id": "reddit_original_013",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/lexapro/comments/130l8ae/10_weeks_into_lexapro_success_story/",
        "subreddit": "r/lexapro",
        "source_kind": "post",
        "post_title": "10 Weeks Into Lexapro (success story)",
        "original_story": """So I’m officially 10 weeks in from taking my first SSRI. I suffered from hormonal anxiety, panic attacks and depression for the last three years. I tried therapy, supplementation, etc. I am extremely fit and workout basically as a living. Nothing was working and I finally hit an all time low in Feb.

I tried taking birth control for about six months, it seemed to help initially, but then I started feeling bad again. After a month off the birth control, my brain went haywire. Bad thoughts, constant state of panic, etc.

I told myself I would never do medication, but at that point, there was nothing else I could do if I wanted to live.

The first five weeks were extremely difficult. I literally sat on Reddit and read through stories all day long. I would get online and look up Lexapro reviews to try to keep myself from giving up…. All. Day. Long.

At week five, I woke up like a brand new person. Literally the first time I didn’t feel anxiety in years.""",
        "struggle_tags": "hormonal anxiety, panic attacks, depression, medication fear",
        "what_helped": "SSRI under clinician care; reading success stories; persistence through adjustment period",
        "support_type": "professional, emotional",
        "timeframe": "about 5-10 weeks",
        "tone": "hopeful, relieved",
        "risk_flags": "medication content, bad thoughts",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Good row about reading stories during early medication anxiety.",
    },
    {
        "story_id": "reddit_original_014",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/lexapro/comments/1hym06v",
        "subreddit": "r/lexapro",
        "source_kind": "post",
        "post_title": "Success Story",
        "original_story": """Hello, I wanted to share my success story because I've been lurking here the last 3-4 months. I ended up starting lexapro after checking myself into a behavioral health hospital during a severe depressive episode. The depression was caused by crippling health anxiety and panic attacks that woke me up out of my sleep every night to vomit. I lost all interest in my hobbies and was unable to function at work because of my anxiety.

I had SI because of it. I experienced it for about 3 weeks before I checked myself into a hospital. I've always had anxiety triggered by sickness and germs but never as bad as I experienced it during that time. I stayed at the hospital for a week and started lexapro along with lorazepam. It took 3 weeks for the side effects to go away and 4 weeks for the panic attacks and vomiting to stop. I stopped taking the lorazepam after 4 weeks. It's now been 3 months on lexapro and I feel like myself again. I'm not experiencing any libido side effects. I'm not foggy I feel like myself. I do get headaches a few times a week but I can manage them with one dose of Tylenol.

My initial side effects were nausea, loss of appetite, fatigue that made me nap for hours during the day and have insomnia at night. I'm not going to lie those first couple of weeks were rough! I was given hydroxyzine to sleep at night and used that for about a month. I no longer have any issues sleeping and stopped taking it.

I say all that to say if you've just started lexapro don't be scared away by all of the doom and gloom here. Everyone is different this medication has a pretty good success rate. Give it 4-6 weeks with therapy and a psychiatrist if you can. I feel like a weight has been lifted. I'm social again and enjoying all of the things I loved. Everyone in my life has noticed a change.

Some advice to get through the rough weeks. Move your body, I'm not saying get a gym membership, but go on a walk. I sometimes walked outside 3x a day maybe, 30 mins sometimes an hour but doing something physically helped take my mind away from my side effects.""",
        "struggle_tags": "health anxiety, panic attacks, depression, insomnia, work impairment",
        "what_helped": "behavioral health hospital; psychiatrist; medication; therapy; walking; time",
        "support_type": "professional, lifestyle, emotional",
        "timeframe": "about 3 months",
        "tone": "reassuring, detailed",
        "risk_flags": "suicidal ideation, medication names, hospitalization",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Strong but clinically sensitive story.",
    },
    {
        "story_id": "reddit_original_015",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/lexapro/comments/1n71zg7",
        "subreddit": "r/lexapro",
        "source_kind": "post",
        "post_title": "My Lexapro success story",
        "original_story": """Somewhat long-winded post ahead, but I feel like these subs tend to lean negative, so I figured I'd throw my own positive experience in here in case anyone is looking for some reassurance.

In 2012, my first year in undergrad, I started experiencing somatic anxiety symptoms, namely chest pain and a shortness of breath. At the time I thought I was having heart problems, but several visits to the hospital revealed nothing. This made, of course, me more anxious (health anxiety, IYKYK), and I started having panic attacks. After several weeks of doctor visits that yielded no diagnoses besides pleurisy (hilariously wrong!) I started feeling depressed, having random bouts of crying and a fear of sleeping, because that's when the symptoms would be at their first. This went on for many weeks; I ended up losing a ton of weight and I just sort of stopped enjoying life.

Finally, I met with a new internal medicine doctor who gave me Xanax for the panic attacks and Zoloft for the depression. The Zoloft didn't work well; I remember telling the doctor that "I didn't feel like myself" which prompted a switch to Lexapro.

Lexapro, as it turned out, was a much better fit. I found I needed the Xanax less and less frequently, my chest pains gradually went away, and I was able to enjoy life again without the ruminations and dread I thought I'd be dealing with forever. It was a long road and it didn't work immediately, and therapy was key in helping me navigate that transitional period. But I ended up staying on Lexapro for over 12 years; it held down my anxiety for the vast majority of those days and it only recently started pooping out earlier in 2025.

In hindsight, I couldn't have asked for a better result. Knowing what I know now about SSRIs, I consider myself lucky to not have had any side effects, and I haven't had a single panic attack since I started Lexapro all those years ago.""",
        "struggle_tags": "health anxiety, panic attacks, depression, chest pain, sleep fear",
        "what_helped": "doctor evaluation; medication switch; therapy; time",
        "support_type": "professional, emotional",
        "timeframe": "weeks to years",
        "tone": "reassuring, long-term",
        "risk_flags": "medication content, health anxiety",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Good long-term stability story.",
    },
    {
        "story_id": "reddit_original_016",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/lexapro/comments/1alxflr",
        "subreddit": "r/lexapro",
        "source_kind": "post",
        "post_title": "Success story!",
        "original_story": """Just wanted to share a success story to give people some hope if they’re going through a rough time! Since 2021, my anxiety started creeping higher and higher to the point that I was having multiple panic attacks per day and became almost unable to leave the house. I started a low dose of Lexparo in October, and by Christmas I was already feeling way better. I was able to leave the house, had more energy, and had an appetite for the first time in years. Now, in early February, I honestly feel like I don’t even have anxiety anymore. I feel normal. I can’t even remember the last time I had a panic attack, and I go days without thinking about my anxiety.

The main side effects at the beginning were hot flashes, dry mouth, jaw clenching, and tiredness. I take it at night because taking it in the morning made me SOOO sleepy - I would fall asleep if I sat down. I had a little bit of insomnia taking it at night, though, which seems to have resolved and I sleep just fine through the night now.""",
        "struggle_tags": "anxiety, panic attacks, agoraphobia, appetite, low energy",
        "what_helped": "SSRI under clinician care; time; sleep adjustment",
        "support_type": "professional",
        "timeframe": "October to February",
        "tone": "encouraging, relieved",
        "risk_flags": "medication side effects",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Concise medication-related anxiety recovery row.",
    },
    {
        "story_id": "reddit_original_017",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/lexapro/comments/1ox0clz/if_you_are_worried_about_taking_lexapro_heres_a/",
        "subreddit": "r/lexapro",
        "source_kind": "post",
        "post_title": "If you are worried about taking Lexapro... here's a success story",
        "original_story": """I have been meaning to return here, as I made a few posts months ago while I was dealing with a particularly bad episode of panic / anxiety.

I have an issue with health anxiety that's pretty severe. It seeps into my thoughts about taking medication (interactions, side effects, etc), which makes me anxious or panic about any medication I'm taking or have been prescribed.

After a particularly bad panic attack, I took the leap of faith and started on lexapro (5mg) with much hesitation. I heard all the horror stories about side effects, ineffectiveness, and dependence. I wasn't sure what it was going to feel like, and I was petrified about the first few weeks and how they would make me feel.

The first few weeks had anxiety scattered throughout, but as I was already going through a rough spot with anxiety, it was pretty much just a continuation of the level I was already at. I had a bit of nausea after taking the pill for about two weeks. It wasn't debilitating, just a bit uncomfortable.. however I timed my dose right after lunch so I would gain my appetite back for dinner. It made me a tad bit sleepy, but again nothing crazy. These eventually subsided and now I feel none of the negative side effects I once started with.

I am about 3 months in and I think I can finally say my worries were incorrect (as they often are). I haven't had a panic attack since the one that pushed me to take the Lexapro. I still have all my emotions, I feel normal and happy. I still get anxious about normal situations. I have been more social in situations where I would've been freaking out. My wife had commented on how much more like myself I've been acting in these scenarios. I have felt panic attacks kindof creeping up and then fizzle out. I'm able to move on from these thoughts, which is definitely a new thing.""",
        "struggle_tags": "health anxiety, panic attacks, medication fear, social anxiety",
        "what_helped": "starting prescribed medication; waiting through adjustment; public meals; social events; spouse support",
        "support_type": "professional, emotional, practical",
        "timeframe": "about 3 months",
        "tone": "reassuring, grounded",
        "risk_flags": "medication content",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Useful for health-anxiety users afraid of treatment.",
    },
    {
        "story_id": "reddit_original_018",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/lexapro/comments/1hp2v05",
        "subreddit": "r/lexapro",
        "source_kind": "post",
        "post_title": "Success story, 8 months in on SSRI",
        "original_story": """I’ve been on SSRI’s for almost 8 months now, having started with Lexapro and moving to Celexa 3 months in. I’ve been waiting to share my success story to make sure the stable place I am in now didn’t fade, and it hasn’t.

I am a 40 year old woman in general good health, and have dealt with anxiety on and off, starting in my 20s. It was always situational, usually triggered in crowded elevators or with public speaking. It would dissipate quickly and didn’t cause me much stress outside these moments. I noticed it getting a little worse in my late 30s. The last few years, I would randomly get overwhelming anxiety while have dinner with groups of people, sometimes having to excuse myself to walk outside or go to the bathroom to splash water on my face and practice breathing.

One night I split a delta 8 edible with my husband and it hit me way too hard. I started panicking and shaking. The next few weeks were hell. I lost my appetite completely. My resting heart rate on my worst day was 160. I was sweating and shaking and my husband took me to the psychiatric ER. I fell into the deepest, darkest depression I had ever known. I wanted to just cease existing. I lost all interest in art, movies, music- anything that I used to find joy in.

The first 10 days on Lexapro were the hardest. Eventually my sleep got better, thankfully. I noticed a drop in my anxiety as well. I was still waking up anxious but it would taper enough in the day so I could at least function. Toward the end to the 3 month mark, it seemed the sadness was progressing. My twin sister is on Celexa and she urged me to switch. Since switching to Celexa, everyday has been better. I’d say it took two months of being on it for my to actually feel like myself again.

Now, I am almost 5 months in and 8 months on SSRI’s total. No more waking up with anxiety, no more sadness. I’m back to making art, watching movies and socializing. I cuddle with my dogs again. I feared the life I had known was gone forever, but Celexa gave me my life back. I am so grateful.""",
        "struggle_tags": "anxiety, panic attacks, depression, appetite loss, sleep, substance-triggered panic",
        "what_helped": "psychiatric ER; psychiatrist; medication switch; sleep support; spouse and sister support; returning to art and socializing",
        "support_type": "professional, family, emotional",
        "timeframe": "about 8 months",
        "tone": "detailed, hopeful",
        "risk_flags": "substance mention, medication names, psychiatric ER, passive death wish",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Good row about needing a different medication fit and time.",
    },
    {
        "story_id": "reddit_original_019",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/lexapro/comments/1b4zogz",
        "subreddit": "r/lexapro",
        "source_kind": "post",
        "post_title": "Succes story",
        "original_story": """I decided to write a success story to help people who are too scared to start or who are already too deep in to it and still feel rock bottom.

Last summer, I was alone on a 5h flight back from my solo trip. Last 2H of the flight my hands started to shake, my neck became incredibly tight and I was not able to take a sip of my water, let alone put food in my mouth. This happened out of nowhere while watching a movie. This led to a full blown panic attack and to probably worst day of my life. I was looking for a way out and even hoping the plane just to crash. I felt incredibly claustrophobic and alone. I didn't know how to act. I told the flight attendant what was happening and she took me to the front section of the plane behind a closed curtain and try to help and figure out what was happening. She gave me snacks, drinks but what helped was simply talking to me. After some time this calmed me down. I remained with her until the landing. Once we touched ground and the door opened, the panic dissapeard.

Months later I woke up middle of night with a full blown anxiety/panic attack. It started around 1am till 5am. I fell asleep till 6am and woke up again to another panic attack. I put my cloth in went to visit my GP first thing in the morning. She prescribed me Xanax for 2 weeks. This worked just fine. After the 2 weeks, the panic/anxiety attack came back immediately. Non stop through the whole day. There was no stop to it. I went back to my GP and she prescribed me Lexapro.

The first 2-3 days were great. The 4th day, the hell started. Anxiety went through the roof. There were moments my body felt on fire or frozen for hours. Shaking, tremors, disoriented. This made me catastrophe and very anxious. I was scared to do simple tasks like doing groceries.""",
        "struggle_tags": "panic attacks, flight anxiety, claustrophobia, insomnia, daily anxiety",
        "what_helped": "flight attendant support; GP visit; short-term medication; SSRI trial; talking during panic",
        "support_type": "professional, emotional, practical",
        "timeframe": "months",
        "tone": "intense, in-progress",
        "risk_flags": "plane crash thought, medication names, severe panic",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Partial story from visible source snippet; useful for panic-on-plane matching.",
    },
    {
        "story_id": "reddit_original_020",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/dryalcoholics/comments/1glx4ku",
        "subreddit": "r/dryalcoholics",
        "source_kind": "comment",
        "post_title": "Life is better sober",
        "original_story": """It took me around 500 days of sobriety to get that feeling that life was better.

I was convinced all those people who preached it were bullshitting, but alas - with enough patience and time, I'm starting to recognise some happiness again and learn that things have actually improved. Just took a while for it to happen and a while for me to see it all.

The mental side of things is a battle for sure, but honestly, even if it doesn't feel that way, you are getting stronger with it as each day passes.

My first year was pretty terrible to be honest. Year 2 was when I started to really work towards things and started being happier and bettering myself, and from then on for the most part it’s been progressively getting better. I still struggle a lot, and have my bad days, but man I feel 100 percent it’s the best decision I ever made.

I went from being bed ridden most of the time with severe depression and anxiety, with the exception of going to work or going out to the bars and making a fool of myself, to getting into fitness and being in the best shape of my life, starting a side hustle that I made pretty decent money at and gave me a natural high that I loved, dating for the first time of my life, and currently having a girl friend I really like, fixing my attendance problems at work and winning awards and bonuses, better relationships with friends and family, and now I planning on going back to college in the spring, and so many more things. None of those would be possible if I was still drinking, and I pray I never go back and screw all these good things up.""",
        "struggle_tags": "alcohol, depression, anxiety, isolation, work attendance, sobriety",
        "what_helped": "long-term sobriety; patience; fitness; side hustle; dating; work attendance; college plans",
        "support_type": "lifestyle, practical, emotional",
        "timeframe": "about 500 days to year 2",
        "tone": "realistic, hard-won",
        "risk_flags": "alcohol use disorder, depression",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Strong sobriety story because it avoids instant-transformation framing.",
    },
    {
        "story_id": "reddit_original_021",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/stopdrinking/comments/18e75ul",
        "subreddit": "r/stopdrinking",
        "source_kind": "comment",
        "post_title": "How is your mental health now that you no longer drink?",
        "original_story": """I've always been a person with big emotions who struggled with anxiety, and I started drinking as a teenager to "help" with that. I drank for close to 20 years and by the time I started trying to get sober last year, I was so anxious I rarely left the house, and I was an angry nightmare with a seriously messy mental space.

It took a little time, but within some weeks of getting sober I noticed my anxiety declining and my sleep improving. I also noticed my emotional reactions becoming less strong. Over the 10.5 months I was sober last year, I felt the best I have EVER felt in my head for my entire life. And then right before Christmas I decided to try moderation, drank for the next six months, and one of the first things that happened was my anxiety sky-rocketed again.

I'm back to six months sober and the anxiety and mental distress have absolutely eased up. I just, in general, feel better. I don't feel as sad as I used to for no reason, I don't feel anxious for no reason, and if there IS reason to feel distress, it's much less strong than it used to be and I'm much better equipped to handle it rationally.""",
        "struggle_tags": "alcohol, anxiety, depression, emotional regulation, sleep",
        "what_helped": "sobriety; noticing relapse/moderation effects; sleep improvement; emotional regulation",
        "support_type": "lifestyle, practical",
        "timeframe": "weeks to 10.5 months, then six months sober again",
        "tone": "reflective, practical",
        "risk_flags": "alcohol use",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Good before/after sobriety row for anxiety and mood.",
    },
    {
        "story_id": "reddit_original_022",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/leaves/comments/gzz82o",
        "subreddit": "r/leaves",
        "source_kind": "post",
        "post_title": "Today marks two years of being sober!",
        "original_story": """Two years ago today was the last day I smoked pot after almost 4 years of constant smoking. Before quitting, I couldn’t go a day without being high so I didn’t have to deal with shame, depression, and high social anxiety. My relationship with my family was a mess, I lost tons of weight because I stopped eating due to depression (I was 6’1” and 120 lbs on a good day), and the idea of going clean for even a single week was inconceivable. I had a moment of clarity that things needed to change. Life couldn’t go on like this.

I know some of you are dealing with this right now.. that it’s too difficult to quit and impossible to imagine a life away from pot, but know this: you can conquer addiction. You are strong, you are a success story in the making, and the struggle is so worth it!! I’m praying for you everyday. Since quitting, I have found my passions in life, I’m going back to school, my relationship with family is soooo much better and I’ve gained a healthy 40 pounds!!""",
        "struggle_tags": "cannabis, depression, social anxiety, shame, family relationships, appetite",
        "what_helped": "quitting cannabis; moment of clarity; school; family repair; finding passions",
        "support_type": "lifestyle, emotional, family",
        "timeframe": "two years",
        "tone": "encouraging, celebratory",
        "risk_flags": "substance use, weight/appetite details",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Strong cannabis recovery row.",
    },
    {
        "story_id": "reddit_original_023",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/leaves/comments/r41ylx",
        "subreddit": "r/leaves",
        "source_kind": "post_and_comment",
        "post_title": "Has your life really gotten that much better?",
        "original_story": """Original post:
Please tell me success stories. I was an every night smoker for about four years. I have bad seasonal depression and have no self worth or motivation. Thought it best to take some time away. Weed wasn’t fun anymore. It was a routine. I’ve completed almost a full month off of it and I can’t even remember what made me want to stop. All I’ve gained is a little confidence in telling myself no and trouble sleeping. I miss it though I know I don’t really need it. I’m eager to smoke again once the month is up but hope / will try to not let it become an every day thing again.

Reply:
im too lazy to type out a full success story but ya quitting weed made my life a lot better. i quit for a year and now i smoke every now and then. A month isnt very long, you wont reap the same benefits of better motivation and drive that you would otherwise get if you quit for say, even half a year. A year would be ideal however. Easier said then done, but i did it, somehow.""",
        "struggle_tags": "cannabis, seasonal depression, low motivation, self-worth, sleep",
        "what_helped": "taking time away; building confidence saying no; longer abstinence period; patience",
        "support_type": "community, practical",
        "timeframe": "one month to one year",
        "tone": "ambivalent, realistic",
        "risk_flags": "substance use",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Useful because it captures early ambivalence, not just after-the-fact success.",
    },
    {
        "story_id": "reddit_original_024",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/stopdrinking/comments/16kruju",
        "subreddit": "r/stopdrinking",
        "source_kind": "comments",
        "post_title": "Did anyones depression get better after stopping drinking?",
        "original_story": """Selected comment:
I feel like I can outwardly cope with the ups and downs of life better, I’m there more for the people who rely on me, and I’m overall a more functional/more likeable human. I’m living my life with integrity now. I’ve been 3 years alcohol free as of this weekend.

While my life is better overall I assumed for years that the root of all my problems was the alcohol. It turns out I likely have a lot deeper seated mental health problems and that’s been harder to cope with. 3 years sober and sometimes it feels like this is as good as it can ever be for me.

But it’s still so much better than when I was drinking.

Selected comment:
I may have had a different experience. Has not drinking helped my depression/anxiety- absolutely! BUT… I have always been the “sweep it under the rug” type of person when things have happened. I’ve been in therapy but never truly done the work to overcome my traumas. I have been sober 13 months now (yay!!!) but around month 3 or 4, I REALLY STRUGGLED emotionally. For the first time in my life, I was feeling everything and having to face it head on without alcohol to push it back under the rug. I almost felt like not drinking was crushing me mentally. I went back to therapy and faced my demons. I am in a much better place now. I just had to keep reminding myself that it was just an illusion, a trick. When it passed I began to experience real, substantial happiness from life again. While drinking I felt like life was happening around me but I felt no connection to it. I only derived an approximation of joy from drinking and that was all. I still had anxiety and depression issues even in sobriety, but working with a counselor I was able to actually target and address those issues. The wounds could heal now.""",
        "struggle_tags": "alcohol, depression, anxiety, trauma, sobriety, therapy",
        "what_helped": "sobriety; therapy; facing trauma; integrity; emotional processing",
        "support_type": "professional, lifestyle, emotional",
        "timeframe": "13 months to 3 years",
        "tone": "realistic, nuanced",
        "risk_flags": "alcohol use, trauma mention",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Excellent nuanced row: sobriety helped but did not erase all mental-health work.",
    },
    {
        "story_id": "reddit_original_025",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/stopdrinking/comments/176set9",
        "subreddit": "r/stopdrinking",
        "source_kind": "comments",
        "post_title": "Not drinking made things worse",
        "original_story": """Selected comment:
Journaling sounds corny but can be so helpful in early sobriety.
I'm 7 months in now, went to therapy 5 months in for anxiety/ptsd issues. Decided to go on an SSNI at 6 months, it was the right choice for me.
It's very hard to handle everything when you're no longer just numbing it out.
I go to bed early often and try to focus on doing anything I can to regulate my nervous system.
The first 3 months were brutal, I was exhausted and felt like shit almost every day.
I'm very glad to be 7 months sober now, it's been a lot of work but I've done a lot of healing.
Growth isn't comfortable but you can do it!

Selected comment:
Stick with it. Start an exercise routine. Help your brain and body produce the good mood chemicals naturally. I got up 30 minutes early every weekday morning and walked 15-20 mins at a fast pace. Starting small made getting into a new routine easier. Over a year and a half later I’m running 3 miles and weight training and have lost 50 pounds from my all time high. My confidence, self esteem, and overall mood, attitude and motivation have improved drastically. My anxiety, stomach issues, bloating, and overall sense of doom and gloom have almost all disappeared.

You have to learn how to be a normal person again. And you have to fight your brain and body along the way. You have to learn how to regulate and handle your emotions, which are more intense without alcohol dulling them. Nothing in life is free or easy, including sobriety.""",
        "struggle_tags": "alcohol, anxiety, PTSD, emotional regulation, exercise, sobriety",
        "what_helped": "journaling; therapy; medication under care; early bedtime; nervous system regulation; walking; running; weight training",
        "support_type": "professional, lifestyle, practical",
        "timeframe": "7 months to 1.5 years",
        "tone": "practical, honest",
        "risk_flags": "alcohol use, medication mention, PTSD",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Good row for users whose symptoms worsen temporarily after quitting substances.",
    },
    {
        "story_id": "reddit_original_026",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/nosurf/comments/ixerld",
        "subreddit": "r/nosurf",
        "source_kind": "post",
        "post_title": "Nosurf helped me manage my OCD, now I am going to college next fall",
        "original_story": """I'm 24, since I was a young age I've had a severe anxiety disorder (OCD) but it was misdiagnosed until last year when I was a severe bout of it and had a mental breakdown.

I has also been on my computer that day for stupid reasons (browsing twitter and reddit)... and I was just disgusted knowing my life was being wasted.

My mental health was poor but I was functioning (barely), I had very few real life friends, I had steady work but I was struggling with working and relied a lot on my partner, I didn't want a college degree and felt like I would never have a reliable job, I was a smart person but a constant lazy failure because I chronically couldn't stop procrastinating.""",
        "struggle_tags": "OCD, anxiety, internet overuse, procrastination, college, work",
        "what_helped": "reducing internet use; recognizing wasted time; changing daily behavior; college goal",
        "support_type": "lifestyle, practical",
        "timeframe": "not specified",
        "tone": "reflective, early-progress",
        "risk_flags": "OCD diagnosis, mental breakdown mention",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Partial but useful digital-behavior mental-health row.",
    },
    {
        "story_id": "reddit_original_027",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/nosurf/comments/1tr5pcn/what_i_realized_about_social_media/",
        "subreddit": "r/nosurf",
        "source_kind": "post",
        "post_title": "What I realized about Social Media",
        "original_story": """It’s just too anxiety inducing for me. I can’t be on it for too long or else I’ll start to get anxious, irritable, and irritated. So this is what worked for me: I treat it as a journal. Once a year when I have something to post, I’ll post, then deactivate and go full hibernation for a year or until I have something to post again. I see it as a reflective journal to see how far I’ve gone looking back at my life, but for myself, not for that bs clout that “influencers” these days post about.

I have zero interest in people’s lives because once I start to view their stories and posts, that’s when my mental health suffers and I start to spiral down again. I begin to think about what others think of me, if they’ve viewed my story, if they liked my story, or why they didn’t like my story etc. So you see what I mean? It causes too much unnecessary anxiety-inducing stress that shouldn’t be there to begin with. It’s all fake man. Don’t waste your time on something that’s not real to begin with.""",
        "struggle_tags": "social media anxiety, comparison, irritability, rumination",
        "what_helped": "deactivating social media; using it as a yearly journal; avoiding story/post checking",
        "support_type": "lifestyle, practical",
        "timeframe": "yearly cycle",
        "tone": "direct, practical",
        "risk_flags": "none",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Good row for anxiety tied to social media checking.",
    },
    {
        "story_id": "reddit_original_028",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/Anxiety/comments/1kuix8i",
        "subreddit": "r/Anxiety",
        "source_kind": "comment",
        "post_title": "Can anybody with a success story please share it?",
        "original_story": """Any success story, big or small.

Selected response:
My psychiatrist told me anyone can recover. I recovered through combination of taking medication, practing exposure therapy and radical acceptance techniques.""",
        "struggle_tags": "anxiety, recovery hope, exposure therapy, radical acceptance",
        "what_helped": "psychiatrist encouragement; medication; exposure therapy; radical acceptance",
        "support_type": "professional, practical",
        "timeframe": "not specified",
        "tone": "brief, hopeful",
        "risk_flags": "medication mention",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "approved_for_closed_mvp",
        "prototype_notes": "Short but clean row for matching to exposure/acceptance.",
    },
    {
        "story_id": "reddit_original_029",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/depression/comments/1h43ina",
        "subreddit": "r/depression",
        "source_kind": "comment",
        "post_title": "Success stories",
        "original_story": """Change of lifestyle, scenery, and staying as busy as I can works better than any meds or therapy for me.

Hope to hear your success story someday!

I’ve been suffering from depression this time around for almost two years.""",
        "struggle_tags": "depression, lifestyle, scenery change, busyness",
        "what_helped": "change of lifestyle; change of scenery; staying busy",
        "support_type": "lifestyle, practical",
        "timeframe": "almost two years",
        "tone": "brief, hopeful",
        "risk_flags": "therapy/med comparison",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Short row; avoid presenting as anti-therapy/anti-medication advice.",
    },
    {
        "story_id": "reddit_original_030",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/AskReddit/comments/16sy8ea",
        "subreddit": "r/AskReddit",
        "source_kind": "comment",
        "post_title": "People who were depressed or had social anxiety, what’s your success story?",
        "original_story": """Depression: Some success with medication and therapy, limited success with lifestyle changes like diet or exercise. Sorry, but I don't really have a success story.""",
        "struggle_tags": "depression, social anxiety, medication, therapy, lifestyle",
        "what_helped": "medication; therapy; limited lifestyle changes",
        "support_type": "professional, lifestyle",
        "timeframe": "not specified",
        "tone": "candid, mixed",
        "risk_flags": "not a full success story",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Included as a realistic non-linear/mixed-outcome row.",
    },
    {
        "story_id": "reddit_original_031",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/depression_partners/comments/1hy5p4r",
        "subreddit": "r/depression_partners",
        "source_kind": "comment",
        "post_title": "Success stories",
        "original_story": """Happy to say that my husband and I are a success story.

I had come to terms with the fact that if he didn’t get treatment, actively work on dealing with the depression and learn to curb his emotionally abusive behavior, I’d have to consider divorcing him. Instead, I went to therapy to learn how to communicate more assertively and without fear of triggering an episode, as I had become scared of speaking my mind because my husband had previously accused me of triggering episodes, which my therapist helped me recognize was misplaced blame and manipulation.""",
        "struggle_tags": "partner depression, communication, assertiveness, relationship boundaries",
        "what_helped": "therapy for partner; assertive communication; recognizing misplaced blame; treatment expectations",
        "support_type": "professional, relationship, emotional",
        "timeframe": "not specified",
        "tone": "serious, boundary-focused",
        "risk_flags": "emotional abuse, partner mental health",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Relationship-oriented row; useful but sensitive.",
    },
    {
        "story_id": "reddit_original_032",
        "source_platform": "Reddit",
        "source_url": "https://www.reddit.com/r/MentalHealthPH/comments/1fpqvja",
        "subreddit": "r/MentalHealthPH",
        "source_kind": "comment",
        "post_title": "Can you share your success story?",
        "original_story": """What I did is go to a therapist and change of environment. Super tiwala ako sa mental health professionals na makakatulog sila sa akin so open ako sa mental health topics. Sabi nga ng therapist, 10% ng effort goes to them, 90% sa pasyente/client. Napagod na rin ako that time and gusto ko na maging maayos. Mas nakatulong kesa sa gamot. And then habang nagrereview, sinabayan ko ng therapy and exercise. After overcoming my fear sa board exam(although I failed), my achievement is kumuha at hinarap ko yung board exam.

Another poster:
Diagnosed with depression last year and still ongoing with meds and therapy. I wouldn't consider it a success story but some people might. Before I was diagnosed, I was really good at compartmentalizing my trauma. I repressed it and focused on work and fun things. But depression doesn't go away, it waits. It waits til your done with whatever kept you busy. So finally, it came to a point it would literally leak out of me.""",
        "struggle_tags": "depression, trauma, therapy, exercise, exam fear, environment change",
        "what_helped": "therapy; change of environment; exercise; facing board exam; medication and therapy",
        "support_type": "professional, lifestyle, practical",
        "timeframe": "not specified",
        "tone": "reflective, culturally specific",
        "risk_flags": "mixed language, trauma mention",
        "retrieval_status": "visible_via_search_result",
        "moderation_status": "needs_human_review",
        "prototype_notes": "Good international/culturally varied row; includes Tagalog/English text.",
    },
]

ROWS.extend(ADDITIONAL_ROWS)


def extension_row(
    story_id: str,
    source_url: str,
    subreddit: str,
    source_kind: str,
    post_title: str,
    original_story: str,
    struggle_tags: str,
    what_helped: str,
    support_type: str,
    timeframe: str,
    tone: str,
    risk_flags: str,
    prototype_notes: str,
    retrieval_status: str = "visible_via_search_result",
    moderation_status: str = "needs_human_review",
) -> dict[str, str]:
    return {
        "story_id": story_id,
        "source_platform": "Reddit",
        "source_url": source_url,
        "subreddit": subreddit,
        "source_kind": source_kind,
        "post_title": post_title,
        "original_story": original_story,
        "struggle_tags": struggle_tags,
        "what_helped": what_helped,
        "support_type": support_type,
        "timeframe": timeframe,
        "tone": tone,
        "risk_flags": risk_flags,
        "retrieval_status": retrieval_status,
        "moderation_status": moderation_status,
        "prototype_notes": prototype_notes,
    }


MORE_ROWS = [
    extension_row(
        "reddit_original_033",
        "https://www.reddit.com/r/PanicAttack/comments/1tj4iff/success_stories_overcoming_anxietypanic/",
        "r/PanicAttack",
        "comment",
        "Success Stories Overcoming Anxiety/Panic?",
        """Yes I have panic disorder. I used exposure therapy, especially the flooding technique, to not have a panic attack for years. Now when I do have strong anxiety and depersonalization-derealization, I practice exposure therapy by asking it to get worse and wanting it to get worse, and it goes away. It is terrifying, but only scary for a short while.""",
        "panic disorder, anxiety, depersonalization, derealization",
        "exposure therapy; flooding; asking symptoms to intensify; reducing fear of symptoms",
        "practical",
        "years",
        "brief, direct, hopeful",
        "panic symptoms",
        "Short comment-style recovery account from a success-story request thread.",
    ),
    extension_row(
        "reddit_original_034",
        "https://www.reddit.com/r/PanicAttack/comments/1tj4iff/success_stories_overcoming_anxietypanic/",
        "r/PanicAttack",
        "comment",
        "Success Stories Overcoming Anxiety/Panic?",
        """I have come such a long way with my panic. Exposure therapy really helped me, because I was thrown into situations where I had to face my triggers and cope with them for months. It was also extremely helpful to have years of different types of therapy and a medication routine that was very tailored and took a while to figure out. My life and my relation to panic has totally changed. I have coping mechanisms and game plans for when an attack starts: get to my safe space, pull the curtains, four-square breathing, take my meds, use my favorite blanket to ground myself. I still get a random panic attack every few months, but I cannot remember the last one. A big part is acceptance that the feeling will come, be unbearable for a while, then end. Learning to sit with it was the hardest part.""",
        "panic attacks, anxiety, triggers",
        "exposure therapy; tailored medication routine; therapy; grounding blanket; breathing; acceptance",
        "professional, practical, emotional",
        "years",
        "practical, reflective",
        "medication mention, panic symptoms",
        "Useful because it describes a concrete panic plan and acceptance framing.",
    ),
    extension_row(
        "reddit_original_035",
        "https://www.reddit.com/r/PanicAttack/comments/1tj4iff/success_stories_overcoming_anxietypanic/",
        "r/PanicAttack",
        "comment",
        "Success Stories Overcoming Anxiety/Panic?",
        """A mixture of cardio-focused exercise two to three times a week, exposure therapy, and inner child work helped me minimize my panic symptoms. I used to be the person who could not leave the house without getting a panic attack. I had to force myself through the discomfort to get to a better place. My panic attacks used to last hours and now only last minutes, and the last time I had one was around six months ago. It was not instant; it took repeated exposures before I noticed a difference, at least six months for me. I went out even when my mind told me to go back home: movies alone, concerts, social events. A panic attack might happen after, but every time I knew I would be okay and just needed to live through the feeling.""",
        "panic attacks, agoraphobia, social anxiety",
        "cardio exercise; exposure therapy; inner child work; repeated solo outings; meditation; breathing",
        "lifestyle, practical, emotional",
        "at least 6 months",
        "encouraging, specific",
        "panic symptoms",
        "Strong mini-story with timeline and examples of exposure steps.",
    ),
    extension_row(
        "reddit_original_036",
        "https://www.reddit.com/r/PanicAttack/comments/1tj4iff/success_stories_overcoming_anxietypanic/",
        "r/PanicAttack",
        "comment",
        "Success Stories Overcoming Anxiety/Panic?",
        """Every time I went outside I felt tightness in my chest and tunnel vision that screamed, run, you are in danger. I started taking Xanax, Dideral, and Prozac. Medication gave relief for a month, but withdrawal from Xanax was worse than having anxiety. I forced myself to go outside with the symptoms and tried to hold still even though I was having an attack. I said, okay, kill me then, I am not doing anything. That was the last attack I lived through, around six months ago. Now I can even go across cities alone. Understanding panic and exposure was the key for me. Your body is not in danger; it is trying to protect you.""",
        "panic attacks, agoraphobia, chest tightness, tunnel vision",
        "exposure; staying still during panic; understanding fight-or-flight; medical testing reassurance",
        "practical, educational",
        "about 6 months",
        "intense, blunt, hopeful",
        "benzodiazepine mention, medication withdrawal, panic symptoms",
        "Needs review because it contains strong language about accepting feared death sensations.",
    ),
    extension_row(
        "reddit_original_037",
        "https://www.reddit.com/r/PanicAttack/comments/1szssll/how_i_overcame_my_severe_panic_attacks_and_social/",
        "r/PanicAttack",
        "post",
        "How I overcame my severe panic attacks and social anxiety condition - Personal Story",
        """It all started with a panic attack in a social situation. I had some social anxiety before that, but it never really bothered me because I did not know it was a condition. That first intense panic attack, where I froze and could not speak in front of someone, traumatized me. From that moment social anxiety completely took over my life. I slowly stopped going to the gym, stopped hanging out with friends, and barely went outside. This went on for about two years. I tried supplements, phenibut, pregabalin, valium, mindfulness, meditation, and breathing techniques, but nothing really worked. Eventually I understood anxiety as an emotion triggered by fight or flight and a perceived threat. In my case, the threat became fear of feeling anxiety in social situations. What changed things was applying acceptance. Instead of avoiding situations, I started doing them anyway. I let anxiety be there and let panic attacks happen. I stopped trying to calm them down or look for temporary solutions. Eventually it stopped controlling me. Now anxiety can show up, but it does not bother me anymore. It is just there, like any other emotion.""",
        "social anxiety, panic attacks, avoidance",
        "acceptance; doing social situations anyway; dropping resistance; understanding fight-or-flight",
        "practical, educational",
        "about 2 years",
        "long-form, explanatory, hopeful",
        "medication/substance mentions, panic symptoms",
        "Direct-open post excerpt; useful for acceptance-based matching.",
        "visible_via_direct_open",
    ),
    extension_row(
        "reddit_original_038",
        "https://www.reddit.com/r/socialanxiety/comments/1k64a5r",
        "r/socialanxiety",
        "post",
        "My story of how I cured from social-anxiety (and keep going every day!)",
        """I will tell you my experience as a person that had very deep social anxiety. Before the success story, I want to describe how my life looked before I overcame my fears. The post frames recovery as something that continues every day, with the poster explaining that they had to keep going, face situations, and build a life while carrying fear rather than waiting to feel completely ready first.""",
        "social anxiety, avoidance, confidence",
        "facing feared situations; daily practice; persistence",
        "practical, emotional",
        "ongoing",
        "reflective, motivational",
        "social anxiety",
        "Search-visible excerpt only; not verified as full original post.",
    ),
    extension_row(
        "reddit_original_039",
        "https://www.reddit.com/r/socialanxiety/comments/1k1ti3i",
        "r/socialanxiety",
        "post",
        "I spent 10 years doing exposure therapy and recorded most wins/losses",
        """Every time I overcame a fear, I wrote down what I did and why it mattered. My best success story was after six years of knowing someone, I finally shared one of the most pivotal stories of my life. It was a long story too, so I had a lot of anxiety about sitting through it without rushing or quitting halfway. The post describes ten years of exposure therapy, recording wins and losses, and gradually building confidence through repeated deliberate action.""",
        "social anxiety, confidence, fear of vulnerability",
        "exposure therapy; recording wins and losses; sharing vulnerable stories; long-term practice",
        "practical, emotional",
        "10 years",
        "reflective, patient",
        "anxiety symptoms",
        "Good example of slow, logged exposure rather than a quick turnaround.",
    ),
    extension_row(
        "reddit_original_040",
        "https://www.reddit.com/r/Anxiety/comments/1febqo6",
        "r/Anxiety",
        "comment",
        "Tell me your success stories",
        """My success in managing anxiety and depression came through persistence. My biggest success story has not been one thing, but developing better self-awareness around my triggers and symptoms. Slowly I improved. Having close friends, my family, and coworkers I liked at my retail job helped me, and I was able to combat my social anxiety a bit.""",
        "anxiety, depression, social anxiety",
        "persistence; self-awareness; support from friends and family; supportive coworkers",
        "emotional, practical",
        "gradual",
        "brief, grounded",
        "depression mention",
        "Comment-style success story focused on support network and trigger awareness.",
    ),
    extension_row(
        "reddit_original_041",
        "https://www.reddit.com/r/Anxiety/comments/1g3dng3",
        "r/Anxiety",
        "comment",
        "Any success stories in managing anxiety?",
        """I overcame years of chronic stress and anxiety. The post is part of a thread asking for more stories about managing anxiety, and the visible text frames recovery as possible through steady management rather than a sudden cure.""",
        "chronic stress, anxiety",
        "anxiety management; steady recovery; persistence",
        "practical",
        "years",
        "brief, hopeful",
        "anxiety symptoms",
        "Thin search-visible excerpt; included for breadth but lower detail.",
    ),
    extension_row(
        "reddit_original_042",
        "https://www.reddit.com/r/HealthAnxiety/comments/17c3uy1",
        "r/HealthAnxiety",
        "comment",
        "How did you overcome it, Its getting hard.",
        """I tried everything I could think of to overcome my health anxiety, but only noticed a significant improvement after getting an app for it. I was skeptical at first but it helped me so much. My anxiety is actually manageable for the first time in six years. I still have low points, but I have a much better toolset for dealing with it now. I have not fully overcome it and still have triggers, but I do not obsess over them anymore.""",
        "health anxiety, obsessive checking, rumination",
        "health anxiety app; better tools; reduced obsession; trigger management",
        "digital tool, practical",
        "6 years",
        "balanced, realistic",
        "health anxiety",
        "Useful for testing search results around health-anxiety tools.",
    ),
    extension_row(
        "reddit_original_043",
        "https://www.reddit.com/r/Agoraphobia/comments/1i5u2cb",
        "r/Agoraphobia",
        "post",
        "Agoraphobia success story! (Long and detailed)",
        """This time last year I was confined to my bed and thought my life was over. Today I have accomplished things I would never have dreamed of twelve months ago. I had anxiety from age eleven and later relapsed badly in June 2023 after a panic episode. I lost an internship, stayed home, ordered groceries online, and did not leave the house for months. In January 2024 I started online counselling. In April my counsellor prompted me to walk around my driveway for a few minutes while on Zoom. She challenged me to go two houses down, then three, then four, until I reached the end of the neighbourhood. Later I walked to the local shop, then the post office, then farther places. Now I have gone to college, socialised, started driving, and taxi rides that used to cripple me barely phase me. I still have a long way to go, but anxiety no longer has complete hold of my life. Exposure therapy helped tremendously with counselling.""",
        "agoraphobia, panic attacks, anxiety, school, work loss",
        "online counselling; exposure therapy; driveway walks; gradual distance expansion; college return",
        "professional, practical, emotional",
        "about 12 months",
        "long-form, detailed, hopeful",
        "panic symptoms, medication mention",
        "Direct-open/source-visible long post; excellent prototype match.",
    ),
    extension_row(
        "reddit_original_044",
        "https://www.reddit.com/r/Agoraphobia/comments/1rszfzu/mini_success_story/",
        "r/Agoraphobia",
        "post",
        "Mini Success Story",
        """My agoraphobia is because of anxiety and OCD. Over the last eight to nine months I have very slowly managed to get comfortable within a four to five mile radius of my house, when a year ago I could barely make it to the bathroom or even make a phone call. I went to a random park four miles away and tanned for over two hours with a friend. I picked up tennis and bike riding around my neighborhood and a park a mile down the road. Recovery is not consistent success; it has ups and downs. Do not beat yourself up over bad days or weeks. Distractions helped in the moment, even bouncing a tennis ball in the driveway at first during exposure therapy. Take small steps whenever comfortable and do not push yourself on the worst days.""",
        "agoraphobia, anxiety, OCD, avoidance",
        "tiny exposure steps; tennis ball distraction; park visits; biking; self-compassion",
        "practical, lifestyle, emotional",
        "8-9 months",
        "gentle, practical",
        "anxiety symptoms",
        "Great example of small-step exposure.",
    ),
    extension_row(
        "reddit_original_045",
        "https://www.reddit.com/r/Agoraphobia/comments/107rl2y",
        "r/Agoraphobia",
        "post",
        "Success Story",
        """Like many of you, I have struggled with agoraphobia for many years, most of my life in fact. I once could travel alone to different countries anxiety-free, but during COVID I relapsed heavily into old ways. Yesterday I drove by myself thirty minutes away on the motorway to visit a historical monument. The location did not matter; the distance, going alone, and driving on the motorway were the hard parts. I experienced rapid heart rate, hypervigilance, sweating, the works. But despite all that, I arrived safe and sound and actually enjoyed myself. I pushed my boundaries and comfort zone farther out. I am proud of myself for making the effort. I have a long way to go, but wanted to share because others helped me before.""",
        "agoraphobia, relapse, driving anxiety, panic symptoms",
        "solo driving exposure; pushing comfort zone; community support",
        "practical, emotional",
        "many years",
        "celebratory, realistic",
        "panic symptoms",
        "Concrete exposure victory after COVID relapse.",
    ),
    extension_row(
        "reddit_original_046",
        "https://www.reddit.com/r/Agoraphobia/comments/1dofl6a",
        "r/Agoraphobia",
        "post",
        "From wanting to die to being able to fly. Success story.",
        """I started developing agoraphobia nine years ago. I had four housebound years where I lost everything, called crisis lines, and wrote goodbye letters to family. I essentially lost my thirties. Eventually I learned to care for myself, stop beating myself up over failures, and only do things I had potential to enjoy. I leaned into the love of my long-term partner and slowly expanded the bubble. I learned the DARE Response. I started doing things I thought I would never do again: shopping, camping, hiking, kayaking, concerts, weddings. I got off meds and lost thirty pounds. It culminated with a trip from Canada to Scotland, my first time traveling internationally.""",
        "agoraphobia, depression, self-harm history, travel fear",
        "self-compassion; partner support; DARE response; gradual exposure; enjoyable activities",
        "emotional, practical, lifestyle",
        "9 years",
        "intense, triumphant",
        "suicidal thoughts, self-harm mention, medication mention",
        "High-impact agoraphobia story; needs careful safety moderation.",
    ),
    extension_row(
        "reddit_original_047",
        "https://www.reddit.com/r/Agoraphobia/comments/1km05uh",
        "r/Agoraphobia",
        "comment",
        "Has anyone recovered?",
        """I still struggle with it, but there was a time I was housebound. Now I own a car and drive back and forth three hours from my city to my hometown. I take the ferry to a remote island to see my boyfriend's parents. I go hiking. I still do some of this with considerable anxiety, but I do it. The next step is going on a plane.""",
        "agoraphobia, driving anxiety, travel anxiety",
        "driving practice; ferry travel; hiking; doing activities with anxiety present",
        "practical, lifestyle",
        "gradual",
        "realistic, hopeful",
        "anxiety symptoms",
        "Good example of functioning with remaining anxiety.",
    ),
    extension_row(
        "reddit_original_048",
        "https://www.reddit.com/r/Agoraphobia/comments/1km05uh",
        "r/Agoraphobia",
        "comment",
        "Has anyone recovered?",
        """I have lived with agoraphobia since I was a pre-teen and am now turning twenty-nine. I have been limited in a lot of ways: housebound for six months, uncomfortable getting a license, and afraid of leaving safe areas. Five years ago I met someone online who changed my life. Two years later I moved 2000 km from my hometown and learned a whole new area of my country. Am I recovered? Absolutely not. But I found people who see me, not my disabilities or fears. I have friends and a partner who help make socializing accessible. I even went on a cruise on the ocean. I still have days I cannot leave, but it can get better, slightly and slowly.""",
        "agoraphobia, isolation, safe zones",
        "supportive partner; accessible socializing; relocation; cruise exposure; accepting gradual progress",
        "emotional, practical",
        "years",
        "warm, realistic",
        "anxiety symptoms",
        "Good relational support story for agoraphobia.",
    ),
    extension_row(
        "reddit_original_049",
        "https://www.reddit.com/r/Agoraphobia/comments/1km05uh",
        "r/Agoraphobia",
        "comment",
        "Has anyone recovered?",
        """Yes, one hundred percent recovered. I developed agoraphobia when I was a teen and had it until my late twenties. I am in my early thirties now. I got better after becoming physically unwell for two and a half years. By the end I was so sick and weak I could not walk. When I recovered from the illness, I started working on strength, and my desperation and desire to walk outweighed my fear of the outside world.""",
        "agoraphobia, physical illness, fear of outside",
        "physical recovery; strength rebuilding; motivation to walk; exposure through necessity",
        "lifestyle, practical",
        "teen years to early thirties",
        "brief, powerful",
        "physical illness mention",
        "Interesting recovery route where physical rehab shifted fear priorities.",
    ),
    extension_row(
        "reddit_original_050",
        "https://www.reddit.com/r/Agoraphobia/comments/1i3eq7d",
        "r/Agoraphobia",
        "comment",
        "Please share success story and help me to gain hope.",
        """I was scared to start medication, but it was a huge turning point for me. I went from completely housebound and having severe physical anxiety symptoms even at home to now going around my city, eating at restaurants, going to events, bars, friends' houses, and reading poetry at open mic nights. I have minimal to no anxiety doing these things. Anytime I do have anxiety I can handle it, it passes quickly, and I carry on with what I am doing. I know meds feel scary and everyone is different, but for me they have been life changing.""",
        "agoraphobia, physical anxiety, housebound",
        "medication; city outings; restaurants; events; open mic exposure",
        "professional, practical, lifestyle",
        "not specified",
        "encouraging, practical",
        "medication mention, anxiety symptoms",
        "Medication-positive agoraphobia account; needs app disclaimer.",
    ),
    extension_row(
        "reddit_original_051",
        "https://www.reddit.com/r/Agoraphobia/comments/yfsl1d",
        "r/Agoraphobia",
        "post",
        "My Success Story",
        """I became agoraphobic last spring and it got to the point I could not even walk around my property. Now I can drive up to an hour away. I have been on dates, socializing, and involved with my university. I am still working on distances and some days are better than others. I started therapy this year, and that was and is key to my recovery. I learned various coping skills and how to be independent. What pushed me was a mandatory dinner I had to attend at my university. I took the plunge, and ever since, I am living like how I used to. I have also widened my social circle, which has been key to recovery. Recovery is possible.""",
        "agoraphobia, social isolation, driving anxiety",
        "therapy; coping skills; mandatory dinner exposure; social circle; driving practice",
        "professional, practical, emotional",
        "about 1 year",
        "clear, hopeful",
        "anxiety symptoms",
        "Concise full success arc with concrete before/after.",
    ),
    extension_row(
        "reddit_original_052",
        "https://www.reddit.com/r/OCDRecovery/comments/1gj6y1i/id_like_to_hear_some_success_stories/",
        "r/OCDRecovery",
        "comment",
        "I'd like to hear some success stories!",
        """I have had OCD since I was five years old and was diagnosed at thirty-four. I have sexually intrusive thoughts, ROCD, and religious themed OCD. I went to exposure and CBT-based therapy and was taught to let the thought come and go on with my day as I would without it in my head. Just let it pass to the background. That helped. Realizing OCD intrusive thoughts are just thoughts helped too. It took me a year to get some real control. I have been about eighty percent in control and twenty percent struggle with only the first intrusive thought. OCD can flare up, but you can get a grip and not spiral as much.""",
        "OCD, intrusive thoughts, ROCD, religious OCD",
        "exposure therapy; CBT; letting thoughts pass; OCD education; therapist support",
        "professional, educational, practical",
        "about 1 year",
        "specific, realistic",
        "sexual/religious intrusive thoughts",
        "Needs sensitive tagging because of OCD theme details.",
    ),
    extension_row(
        "reddit_original_053",
        "https://www.reddit.com/r/OCDRecovery/comments/1gj6y1i/id_like_to_hear_some_success_stories/",
        "r/OCDRecovery",
        "comment",
        "I'd like to hear some success stories!",
        """I have had OCD since I was five. Debilitating OCD ruined just about every experience in my life that should have been happy. A year ago my OCD was so terrible I almost went to a mental hospital. I was losing friendships and relationships. I thought I would be stuck with the obsession forever or it would kill me. Medication changes with my psychiatrist helped after weeks, but I still needed therapy. The obsessions were still there, but much more manageable. Later I switched medication again and have never felt so good. My obsessions are gone and I feel normal. I am sure I will face a new theme one day, but now I know no matter how hard it gets, it will be okay.""",
        "OCD, obsessions, relationships, hopelessness",
        "psychiatrist support; medication adjustment; therapy; patience through medication changes",
        "professional, emotional",
        "about 1 year",
        "hopeful, medication-focused",
        "medication mention, suicidal ideation implication",
        "Medication-focused OCD recovery story; requires disclaimer in product.",
    ),
    extension_row(
        "reddit_original_054",
        "https://www.reddit.com/r/OCD/comments/17idy3c",
        "r/OCD",
        "post",
        "Success Story",
        """Thanks to ERP, I am really living my life again and I am free. I do almost no compulsions now. In the last two years all my compulsions were mental and it was the worst OCD ever. I cannot say I am free of intrusive thoughts or anxiety pain, mostly headaches, but it is better than life before with compulsions. Most importantly, after half a year of ERP I feel it getting better. I get fewer intrusive thoughts and headaches; before it was twenty-four seven. I have struggled with OCD since childhood, with rituals, good and bad numbers, pure OCD, just-right OCD, health anxiety, fear of going blind, and more. Therapy was unavailable during the pandemic, and benzo withdrawal spiked OCD, but ERP was the only thing that finally helped.""",
        "OCD, pure OCD, compulsions, health anxiety",
        "ERP; reducing compulsions; persistence despite withdrawal; OCD education",
        "professional, practical",
        "half a year of ERP",
        "detailed, hopeful",
        "benzodiazepine withdrawal, intrusive thoughts",
        "Strong ERP example with symptom progression.",
    ),
    extension_row(
        "reddit_original_055",
        "https://www.reddit.com/r/OCD/comments/k7p6en",
        "r/OCD",
        "post",
        "SUCCESS STORY YAY I AM HAPPY",
        """I have been battling Pure OCD, specifically HOCD and ROCD, for seven to eight months and have started getting glimmers of hope every day. We have all been there: suicidal ideations, thoughts of it lasting forever, and whatnot. As a survivor, I know I will be prepared for relapse if one comes, thanks to self-therapy including CBT, ACT, ERP, meditation, exercise, and reading books about purpose and hope. I never took medication because I was afraid of it altering my brain. I wanted to carve out a shining story to motivate people amidst the sadness and despair.""",
        "pure OCD, HOCD, ROCD, suicidal ideation",
        "CBT; ACT; ERP; meditation; exercise; reading; relapse preparation",
        "educational, lifestyle, practical",
        "7-8 months",
        "energetic, hopeful",
        "suicidal ideation, intrusive thoughts",
        "Contains sensitive OCD subtype language and self-harm reference.",
    ),
    extension_row(
        "reddit_original_056",
        "https://www.reddit.com/r/OCD/comments/1ahi82h",
        "r/OCD",
        "post",
        "My OCD (success) story",
        """I used to visit the OCD subreddit compulsively when I was at my lowest during my OCD crisis. It gave me hope to read others' success stories. I never thought I would be posting my own. I had genetic predisposition and was always obsessive, but in 2019 after a breakup and rumors in my social circle I developed HOCD. A therapist identified the obsessions as textbook OCD. Later the compulsions worsened into checking, cleaning, and organizing, and the sight of my girlfriend could trigger near panic. During the chaos I began seeing an OCD specialist who introduced ERP. It allowed me to stay sane. Years later, after keeping up the techniques, my OCD is different and much improved.""",
        "OCD, HOCD, compulsions, panic, relationship anxiety",
        "OCD specialist; ERP; continuing techniques; reading success stories",
        "professional, practical, emotional",
        "years",
        "reflective, hopeful",
        "intrusive thoughts, sexual trauma nightmare detail in source",
        "Sensitive source; excerpt avoids explicit triggering details.",
    ),
    extension_row(
        "reddit_original_057",
        "https://www.reddit.com/r/OCDRecovery/comments/1riattu/how_i_recovered_from_ocd/",
        "r/OCDRecovery",
        "post",
        "How i recovered from OCD",
        """I was diagnosed with OCD in 2016, Pure O. A psychiatrist diagnosed me and prescribed Zoloft, but it did not work for me. I did not go to therapy because I could not find a therapist and, as a busy mom, it got pushed aside. Slowly I recovered. At one point I got completely fed up with self-pity and playing victim. I was tired of being pushed around by OCD and having my head and life revolve around OCD twenty-four seven. I have a family and a whole life ahead of me, and I decided I was done living that way and was going to heal. That mindset shift mattered for me.""",
        "OCD, pure O, rumination",
        "mindset shift; refusing to organize life around OCD; family motivation; gradual recovery",
        "emotional, practical",
        "years",
        "self-directed, motivational",
        "medication mention",
        "Personal self-directed recovery account; not medical advice.",
    ),
    extension_row(
        "reddit_original_058",
        "https://www.reddit.com/r/OCDRecovery/comments/1rn60op/i_have_completely_recovered_from_what_id_consider/",
        "r/OCDRecovery",
        "post",
        "I have completely recovered from what I'd consider extreme OCD",
        """I have completely recovered from what I would consider extreme OCD. My main themes were harm OCD and health OCD, but I had many other fears: magical thinking, fear of being delusional, and seeing possible dangers everywhere. I thought someone might want to hear a success story. I recovered through medication, practicing ERP, and radical acceptance.""",
        "OCD, harm OCD, health OCD, magical thinking",
        "medication; ERP; radical acceptance",
        "professional, practical, emotional",
        "not specified",
        "brief, confident",
        "harm OCD, medication mention",
        "Brief but useful for matching harm/health OCD recovery searches.",
    ),
    extension_row(
        "reddit_original_059",
        "https://www.reddit.com/r/OCD/comments/gx0cy7/id_like_to_hear_some_success_story_from_people/",
        "r/OCD",
        "comment",
        "I'd like to hear some success story from people who overcame their OCD!",
        """Here is what I will say about recovery: recovery is a choice you have to make every day. I overcame the worst of my OCD about six years ago through therapy and medication. The visible comment frames improvement as an ongoing daily practice rather than a one-time cure.""",
        "OCD, recovery maintenance",
        "therapy; medication; daily recovery choice",
        "professional, practical",
        "about 6 years",
        "brief, grounded",
        "medication mention",
        "Short source-visible comment on maintenance and treatment.",
    ),
    extension_row(
        "reddit_original_060",
        "https://www.reddit.com/r/BPDrecovery/comments/yvi78t/any_bpd_recovery_success_stories/",
        "r/BPDrecovery",
        "comment",
        "any bpd recovery success stories?",
        """I tried going through BPD without medication, therapy, or any aid for a long time. I did okay through research and reading, but I was still not in a healthy place and did not have the right tools. My first year in therapy and DBT was very successful and I accomplished more than I had in the past few years alone. I sought help because I truly wanted to change and be better to myself and those around me. Routine, good health, sleep, therapy, exercise, and going outside actually work. DBT saved my life. I did DBT in group therapy twice and it got me over the worst period of my life when I was barely surviving.""",
        "BPD, emotional dysregulation, relationships, trauma",
        "DBT; group therapy; routine; sleep; exercise; going outside; wanting help",
        "professional, lifestyle, emotional",
        "first year of therapy",
        "direct, hopeful",
        "BPD, medication mention",
        "Strong DBT story; useful for skills-based matching.",
    ),
    extension_row(
        "reddit_original_061",
        "https://www.reddit.com/r/BPDrecovery/comments/yvi78t/any_bpd_recovery_success_stories/",
        "r/BPDrecovery",
        "comment",
        "any bpd recovery success stories?",
        """What has helped me is meditation through Headspace, journaling with shadow work prompts, exercising, and medication. I was diagnosed with ADHD, and when I started taking meds for that, it also helped some overlapping BPD symptoms. My therapist and I talk about self-compassion a lot, which is hard for people with BPD to practice. My self-esteem was very low when I was diagnosed and I was mean to myself. I would not forgive myself for anything, while tolerating bad behavior from others. Learning self-compassion helped me heal and believe I had a life worth living. I am a success story and feel proud and privileged to remember it every day.""",
        "BPD, ADHD, low self-esteem, self-compassion",
        "meditation app; journaling; exercise; medication; therapy; self-compassion",
        "professional, lifestyle, emotional",
        "not specified",
        "warm, self-reflective",
        "medication mention, BPD",
        "Good for self-compassion and overlapping ADHD/BPD searches.",
    ),
    extension_row(
        "reddit_original_062",
        "https://www.reddit.com/r/BPDrecovery/comments/yvi78t/any_bpd_recovery_success_stories/",
        "r/BPDrecovery",
        "comment",
        "any bpd recovery success stories?",
        """I had BPD, CPTSD, anxiety disorder, and dissociative disorder. I started treatment at fifteen because I was suicidal and deeply depressed. I got weekly supportive counseling until I started CBT psychotherapy at twenty. At twenty-six I started DBT and went through it two times. After that I got trauma therapy, but because I had already learned strong distress and mindfulness skills in DBT, we ended trauma therapy early. At thirty I was diagnosed to be in recovery. I am thirty-seven now and have not suffered from BPD symptoms in years. My life is healthy and in balance. I know my limits and triggers and avoid pushing them. DBT therapy saved my life.""",
        "BPD, CPTSD, anxiety, dissociation, suicidality",
        "supportive counseling; CBT; DBT twice; trauma therapy; distress tolerance; mindfulness",
        "professional, emotional, practical",
        "15 to 37",
        "long-form, hopeful",
        "suicidal thoughts, BPD, dissociation",
        "High-value long recovery timeline.",
    ),
    extension_row(
        "reddit_original_063",
        "https://www.reddit.com/r/BPDrecovery/comments/yvi78t/any_bpd_recovery_success_stories/",
        "r/BPDrecovery",
        "comment",
        "any bpd recovery success stories?",
        """I am now the most content I have ever been. In a few weeks I am marrying my love, who has stood by my side for six years. I dated him for months before I told him about my BPD and I was terrified how he would react, but he was unfazed and accepted me wholly. A few years ago I found myself in a full place of recovery. The suicidal thoughts and urges are completely gone. My mood is stable. I am more mindful in how I interact with others, what I share, and how I listen. I have compassion for my past self. Recovery is not impossible. You have to want it and work for it. It takes years, but you can get there if you put your all into it.""",
        "BPD, suicidal thoughts, relationships, mood instability",
        "supportive partner; mindfulness; years of work; self-compassion",
        "emotional, practical",
        "years",
        "tender, hopeful",
        "suicidal thoughts, BPD",
        "Relationship-centered recovery story.",
    ),
    extension_row(
        "reddit_original_064",
        "https://www.reddit.com/r/BPD/comments/pyvazq",
        "r/BPD",
        "post",
        "Our success story",
        """It is rare to see BPD success stories, so I want to share my husband's and my story. My husband has BPD, ADHD, and PTSD, and I have anxiety and PTSD. We have been together seven and a half years and are still deeply in love. Before diagnosis things were difficult, but ever since he was diagnosed I have seen a change. He has been using DBT, staying in touch with mental health providers about medication, staying employed despite burnout, trying to eat regularly, drinking water, getting sun for fifteen minutes a day, and taking walks. He told trusted friends about his disorder and they accepted him. He is open and honest with me about symptoms day to day and respects my boundaries most of the time.""",
        "BPD, ADHD, PTSD, anxiety, relationships",
        "diagnosis; DBT; medication management; employment; hydration; sunlight; walks; trusted friends; boundaries",
        "professional, lifestyle, emotional",
        "7.5 years",
        "relational, hopeful",
        "BPD, PTSD, medication mention",
        "Partner-observed progress story; good for caregiver/relationship search.",
    ),
    extension_row(
        "reddit_original_065",
        "https://www.reddit.com/r/leaves/comments/p3mv8e",
        "r/leaves",
        "post",
        "One year!",
        """After several attempts to quit weed for good, I finally hit the one-year mark. I estimate I am about eighty to ninety percent recovered. The clearest difference is my ability to handle stress. In the beginning, if I slept poorly, drank too much, or got into an argument, I would feel horribly dissociated and anxious. These days almost nothing brings me anxiety. I feel much more grounded and creative. My rough timeline: months one to three were constant anhedonia, depression, anxiety, and fragile nervous system. Months three to six peaked at the three-month mark and the worst was over by six months. Months six to nine brought slow return of creativity, sociability, and joy. Months nine to twelve brought increased normalcy.""",
        "cannabis withdrawal, anxiety, depression, anhedonia, dissociation",
        "abstinence; patience; stress tolerance; tracking recovery timeline",
        "lifestyle, practical",
        "1 year",
        "timeline, hopeful",
        "substance recovery, depression",
        "Clear cannabis recovery timeline.",
    ),
    extension_row(
        "reddit_original_066",
        "https://www.reddit.com/r/leaves/comments/1tpa3ap/anxiety/",
        "r/leaves",
        "comment",
        "Anxiety?!?!",
        """One year sober. What even is anxiety? After a few months sober, I started to notice changes. Things stopped affecting me like they used to. It has been a huge relief not getting anxious over everything like I once did, and my confidence is coming back slowly but surely. Do not let other people's experiences scare you. Maybe yours gets worse, maybe it gets better. You will not know until you get there, and if it gets worse, seek professional help. Weed certainly will not fix it, but talking with a professional can.""",
        "cannabis use, anxiety, confidence",
        "sobriety; time; professional help if needed; confidence rebuilding",
        "lifestyle, professional",
        "1 year",
        "brief, reassuring",
        "substance recovery",
        "Comment from a quitting-weed anxiety thread.",
    ),
    extension_row(
        "reddit_original_067",
        "https://www.reddit.com/r/leaves/comments/1tpa3ap/anxiety/",
        "r/leaves",
        "comment",
        "Anxiety?!?!",
        """I take medication for anxiety and depression. I think weed was interfering with it because it felt like it was not working. I tried several doses and types. My doctor brother told me I needed to quit, so I quit cold turkey. The first two to three weeks anxiety was worse than before, like panic-attacks-crawl-into-a-ball bad. I knew from everything I read that it was common. Two months in, my anxiety is still there but it is so much better than it has been the past year. I am so glad I took his advice.""",
        "cannabis use, anxiety, depression, panic attacks",
        "quitting cannabis; medical advice from doctor; waiting through withdrawal; medication support",
        "professional, lifestyle",
        "2 months",
        "relieved, realistic",
        "medication mention, substance recovery, panic symptoms",
        "Needs review because it includes cold-turkey quitting and medication.",
    ),
    extension_row(
        "reddit_original_068",
        "https://www.reddit.com/r/leaves/comments/1tpa3ap/anxiety/",
        "r/leaves",
        "comment",
        "Anxiety?!?!",
        """I am nearly a month in and my anxiety has settled down significantly, though it was really bad for the first couple weeks. I have an anxiety disorder already and was going through stressful times, but quitting made it all hit at once: pounding heart, sleep troubles, and feeling like I was going crazy. Things that helped: cutting caffeine, using a quit-weed app to track symptoms and remind myself this sucks but is normal and will be over eventually, breathing exercises, and catching myself when I realized I was spiraling. It is hard but gets easier.""",
        "cannabis withdrawal, anxiety disorder, insomnia, panic symptoms",
        "cut caffeine; quit-weed tracking app; breathing exercises; catching spirals",
        "digital tool, lifestyle, practical",
        "about 1 month",
        "practical, encouraging",
        "substance recovery, anxiety symptoms",
        "Useful for early withdrawal coping search.",
    ),
    extension_row(
        "reddit_original_069",
        "https://www.reddit.com/r/leaves/comments/1tpa3ap/anxiety/",
        "r/leaves",
        "comment",
        "Anxiety?!?!",
        """A little more than a year out from quitting, I would say quitting made my anxiety better, with the nuance of a much lower baseline anxiety but intense periods of anxiety right after I quit. It is hard to say if the anxiety got more intense or if taking away my major and only coping skill made it harder to manage. Overall, my baseline is much lower.""",
        "cannabis use, anxiety, coping skills",
        "long-term abstinence; noticing baseline anxiety; tolerating intense early waves",
        "lifestyle, reflective",
        "a little over 1 year",
        "nuanced, realistic",
        "substance recovery",
        "Good balanced account of short-term worse, long-term better.",
    ),
    extension_row(
        "reddit_original_070",
        "https://www.reddit.com/r/leaves/comments/15dogke",
        "r/leaves",
        "post",
        "1 Year Weed & Cigarette Free Withdrawal Timeline",
        """On July 28, 2022 I quit weed, caffeine, and cigarettes simultaneously. In 2019 I lost my father and started smoking weed to numb the hurt. I experienced every major withdrawal symptom and was convinced I was dying or had brain damage. I had anxiety, panic attacks, depersonalization, severe brain fog, tinnitus, heart palpitations, dizziness, depression, and more. I went to the ER three times thinking I was having a heart attack, but the monitor showed nothing wrong. The first three months were hell. After three months symptoms slowly tapered. At five months they lessened significantly and I returned to the gym. As of one year, symptoms are basically gone, with mild waves that I recognize and can manage by reminding myself how far I have come.""",
        "cannabis withdrawal, nicotine, caffeine, panic attacks, depersonalization, depression, grief",
        "abstinence; gym; time; symptom tracking; remembering progress",
        "lifestyle, practical",
        "1 year",
        "timeline, intense, hopeful",
        "substance recovery, ER visits, panic symptoms, depression",
        "Strong withdrawal timeline; medical disclaimer needed.",
    ),
    extension_row(
        "reddit_original_071",
        "https://www.reddit.com/r/leaves/comments/x6yvgj",
        "r/leaves",
        "post",
        "I quit weed and it has been the best decision of my life",
        """Ten months ago I was 125 pounds, suicidally depressed, anxious as hell, unable to do my passions, spiritually sick with no social life, stomach issues, and smoking weed every day. Ten months after quitting weed I am 195 pounds, the happiest I have ever been, with no severe depression or severe anxiety. I am passionate about my passions, spiritually thriving, have a job, have a social life, am clear headed, make better decisions, learn new things about myself, am honest, help others, and go on dates. I asked for help, threw everything weed-related away, go to AA meetings, work the twelve steps, lived in sober living, took suggestions from people I admire, stayed honest and social, did things I did not want to do but had to do, stayed consistent, never gave up, and realized I am addicted to weed.""",
        "cannabis addiction, depression, anxiety, isolation",
        "asking for help; discarding paraphernalia; AA; 12 steps; sober living; honesty; consistency",
        "community, lifestyle, emotional",
        "10 months",
        "before-after, hopeful",
        "suicidal depression, substance recovery",
        "Strong before/after sobriety story.",
    ),
    extension_row(
        "reddit_original_072",
        "https://www.reddit.com/r/leaves/comments/1b0qbx9",
        "r/leaves",
        "comment",
        "Has anyone in this sub actually quit weed successfully?",
        """I quit weed for over a year now, one year and six months. I did not view it as hurting me and thought it was a great coping skill, the only thing getting me through the day. I blamed depression and anxiety. Then I saw this Reddit group and followed it, which sparked research about how weed affected mental health and my body. I was scared and anxious to quit, but I was broke, my mental health and memory were worse than ever, and I was losing friends. Reading people's experiences gave me hope and courage. I did it with an amazing support system of friends and my boyfriend. After quitting, I slowly noticed motivation, confidence, memory, diet, and savings improved. I was not anxious about where and when I would smoke next. I set boundaries with friends who smoked, with a clear mind. My skin cleared and I smelled better.""",
        "cannabis use, anxiety, depression, memory, motivation",
        "support group; research; friends; partner support; boundaries with smoking friends",
        "community, emotional, practical",
        "1.5 years",
        "reflective, hopeful",
        "substance recovery",
        "Good social-support quitting story.",
    ),
    extension_row(
        "reddit_original_073",
        "https://www.reddit.com/r/leaves/comments/z63r0y",
        "r/leaves",
        "post",
        "One year weed free",
        """I smoked weed for twenty years, every day for the last ten. I quit a year ago. Since then I gained clarity and confidence to leave an abusive relationship, quit a toxic job, make more money as a freelancer because I can set schedules and goals and stick to them, live without a roommate for the first time, save for retirement, and take a vacation overseas. It is easier to be social. I can hold conversations and connect with people. The paranoia is gone. My apartment is clean and lovely. I have plants now and can keep them alive. It has not all been roses; I have had low moments and had to confront why I was losing myself in a haze every afternoon or evening. Smoking again will not help depression; it will only make it worse.""",
        "cannabis use, depression, anxiety, paranoia, abusive relationship",
        "quitting cannabis; freelancing goals; social reconnection; clean environment; confronting avoidance",
        "lifestyle, practical, emotional",
        "1 year",
        "before-after, reflective",
        "substance recovery, abusive relationship mention",
        "Strong life-rebuild after cannabis story.",
    ),
    extension_row(
        "reddit_original_074",
        "https://www.reddit.com/r/leaves/comments/1ron5g7/1_year_without_weed_my_mind_finally_came_back/",
        "r/leaves",
        "post",
        "1 Year Without Weed My Mind Finally Came Back",
        """My anxiety had gotten a lot better, but depression and boredom were still rough. Slowly something changed. My sleep became deeper. The constant tightness in my chest disappeared. The terrifying racing heart at night stopped showing up. The fear of dying randomly that used to haunt me began fading. It felt like my nervous system was finally breathing again. Today it has been one full year without weed: no panic attacks, no constant anxiety cloud hanging over my head. Sometimes I still get anxious like any human does, but it passes and no longer controls my life. The life I thought weed enhanced was actually being stolen quietly and slowly. Your brain is not broken; it just needs time to heal.""",
        "cannabis use, anxiety, panic attacks, depression, insomnia",
        "one year abstinence; time; sleep recovery; nervous system calming",
        "lifestyle, emotional",
        "1 year",
        "soothing, hopeful",
        "substance recovery, panic symptoms",
        "Good emotional 'brain healing' narrative.",
    ),
    extension_row(
        "reddit_original_075",
        "https://www.reddit.com/r/leaves/comments/h9plc4",
        "r/leaves",
        "post",
        "One. Year. Free.",
        """June 13, 2020 marked one year weed and other substance free. I stopped because of a panic attack while high. The first six months were the most difficult of my life and the first three were the most horrendous, but everything was downhill after that. I experienced extreme anxiety, panic, fear, weight gain, and many unpleasant things. I went to psychotherapy because of anxiety, started working out because of weight gain, got out of my comfort zone because of fear, and started to enjoy life rather than smoking every night. I do not regret it. It made me better. I did not know who I was or what I was capable of before I stopped. I was always called lazy; turns out I can get the job done and wake up early. After I stopped, my life was back on track.""",
        "cannabis use, panic attacks, anxiety, fear, identity",
        "psychotherapy; working out; comfort-zone exposure; abstinence",
        "professional, lifestyle, practical",
        "1 year",
        "reflective, proud",
        "substance recovery, panic symptoms",
        "Good identity and self-efficacy recovery story.",
    ),
    extension_row(
        "reddit_original_076",
        "https://www.reddit.com/r/stopdrinking/comments/wzk5fs",
        "r/stopdrinking",
        "post",
        "1 year sober",
        """I have had problems with alcohol since I was a teenager. I did not drink all the time, but when I did I would binge. Twice before I gave up drinking for a year, then slowly returned: only on occasions, then only weekends, then binge drinking again. I have severe depression and when I drink I feel great, but it burns through all my happiness and I get unbelievably low for weeks after. The last time I drank I had one of the best nights of my life. The day after and for weeks following I was miserable and wanted to die. I got close to making it happen. So I stopped cold turkey. Not drinking has cost me socially and at work, and people question me, but a year sober matters.""",
        "alcohol, binge drinking, depression, suicidal thoughts",
        "sobriety; recognizing depression crash after drinking; cold turkey; persistence despite social pressure",
        "lifestyle, emotional",
        "1 year",
        "raw, sober, reflective",
        "suicidal thoughts, alcohol recovery",
        "Needs safety moderation; strong alcohol/depression story.",
    ),
    extension_row(
        "reddit_original_077",
        "https://www.reddit.com/r/stopdrinking/comments/1r1q67r/1_year_sober_missing_my_old_life/",
        "r/stopdrinking",
        "post",
        "1 year sober, missing my old life",
        """I am a terribly anxious alcoholic and drug addict, though alcohol was my primary poison. I do not miss years of alcohol dependence: sitting in my room alone and drunk twenty-four hours a day, crying when I walked to the bottle shop, wanting to stop so badly but feeling too mentally messed up to ask for help, seizures and detoxes, each time feeling fleeting hope before falling back because I felt like I did not belong in the world. I do miss parties, festivals, social gatherings, dancing, and alcohol pushing my anxiety away so I could express myself without crippling fear of rejection. I know I cannot get those times back, but I am at a year and trying to keep going.""",
        "alcohol, drug use, anxiety, social fear, sobriety ambivalence",
        "sobriety; honest grief for old life; community support; remembering consequences",
        "community, emotional",
        "1 year",
        "ambivalent, honest",
        "alcohol recovery, drug use, seizures/detox mention",
        "Useful because it captures ambivalence instead of simple triumph.",
    ),
    extension_row(
        "reddit_original_078",
        "https://www.reddit.com/r/depression_partners/comments/1hy5p4r",
        "r/depression_partners",
        "comment",
        "Success stories",
        """My husband and I are a success story. He developed major depressive disorder about five years ago after twenty happy years of marriage. For nearly a year he lashed out, blamed me for things that were not my fault, found fault in everything, and refused help. I retreated to avoid emotional abuse and became his caregiver while handling our shared lives. During recovery we healed our relationship by doing favorite sports together, spending time with friends, finding things to laugh about, and rebalancing responsibilities so he could take on more shared and emotional work. It took time for me to trust him again and set boundaries, and time for him to curb behaviors and find happiness. Today he is about ninety-five percent back, with the remaining five percent more empathetic, introspective, and open-minded.""",
        "major depression, relationship strain, caregiving, boundaries",
        "coping tools; sports together; friends; laughter; rebalanced responsibilities; boundaries",
        "emotional, practical, relational",
        "about 5 years",
        "relational, hopeful",
        "emotional abuse mention, depression",
        "Partner perspective; useful for relationship-support searches.",
    ),
    extension_row(
        "reddit_original_079",
        "https://www.reddit.com/r/depression_help/comments/103m0uf",
        "r/depression_help",
        "comment",
        "Success stories",
        """I guess I would be a success story. Once you start feeling better and getting older, you realize everyone has demons they are tackling and everyone is trying to function. I went from not being able to leave the house because of anxiety and depression to being functional and successful in some areas, while working on improving others like anyone else. My triggers can still set me off sometimes, but I use coping skills and push through.""",
        "depression, anxiety, leaving the house, triggers",
        "coping skills; aging perspective; pushing through triggers; functional goals",
        "practical, emotional",
        "not specified",
        "brief, realistic",
        "depression, anxiety symptoms",
        "Good compact depression/anxiety improvement story.",
    ),
    extension_row(
        "reddit_original_080",
        "https://www.reddit.com/r/depression_help/comments/mi50zk",
        "r/depression_help",
        "comment",
        "Looking for success stories from adults who have been depressed most of their lives",
        """I have been in depression for seven years and also suffer from anxiety, with panic attacks in the past. It might not be the success story you are looking for, but I am getting better every year. Right now I am able to apply for jobs, while a few years ago I could not do so without panic attacks. I let myself imagine for one minute that I really enjoy my work and that it is possible to enjoy working. I felt what that felt like and amplified the feeling. The idea was to consciously imprint a different belief instead of only imprinting negative experiences.""",
        "depression, anxiety, panic attacks, work avoidance",
        "gradual improvement; applying for jobs; positive visualization; belief work",
        "practical, emotional",
        "7 years",
        "reflective, self-directed",
        "depression, panic symptoms",
        "Nonclinical self-directed coping story; label carefully.",
    ),
    extension_row(
        "reddit_original_081",
        "https://www.reddit.com/r/Psychosis/comments/1qrsrsy/success_stories/",
        "r/Psychosis",
        "comment",
        "Success stories",
        """I spent pretty much a full year in bed after my psychosis. It took me a long time to recover fully, but now it seems like a million years ago. I smile, I laugh. It will come. Give yourself time. Try not to avoid people too much and share your feelings with those willing to hear.""",
        "psychosis recovery, post-psychosis depression, isolation",
        "time; social connection; sharing feelings; not isolating",
        "emotional, practical",
        "about 1 year",
        "gentle, hopeful",
        "psychosis mention",
        "Short post-psychosis encouragement.",
    ),
    extension_row(
        "reddit_original_082",
        "https://www.reddit.com/r/Psychosis/comments/1qrsrsy/success_stories/",
        "r/Psychosis",
        "comment",
        "Success stories",
        """Yes, full recovery is possible but it may take years to feel like yourself again. My last break was caused by PTSD, trauma, and benzo withdrawals, and it took a full three years to get back to baseline. All negative symptoms are gone, including the anhedonia. I would say reduce triggers if you can, reduce stress, do not isolate, create a routine, find a job, and keep as busy as possible. I had many breaks and there were none I have not recovered from.""",
        "psychosis recovery, PTSD, trauma, anhedonia, isolation",
        "reduce triggers; reduce stress; routine; job; staying busy; avoiding isolation",
        "professional, lifestyle, practical",
        "about 3 years",
        "direct, hopeful",
        "psychosis, trauma, benzodiazepine withdrawal",
        "Useful for long-tail recovery and routine matching; needs clinical caution.",
    ),
]

ROWS.extend(MORE_ROWS)


BULK_SOURCE_THREADS = [
    {
        "source_url": "https://www.reddit.com/r/Agoraphobia/comments/txa59a",
        "subreddit": "r/Agoraphobia",
        "post_title": "Recovery?!",
        "struggle_tags": "agoraphobia, panic disorder, anxiety, alcohol recovery",
        "what_helped": "500 days sober; weekly therapy; medication consistency; errands; exercise; gradual obligations",
        "support_type": "professional, lifestyle, practical",
        "timeframe": "500 days",
        "risk_flags": "alcohol recovery, panic symptoms",
        "fragments": [
            "After hitting rock bottom with drinking and self-medicating agoraphobia and panic, the poster reached 500 days sober and described a cleaner determination to rebuild life.",
            "The poster said weekly therapy, taking medication correctly, eating regularly, and pushing forward after setbacks helped them become functional again.",
            "A major sign of progress was being able to drive an hour to work daily after previously feeling trapped by panic and agoraphobic avoidance.",
            "They described saving money, having hobbies, running errands, taking fewer naps, exercising, and stepping up weekly to something new.",
            "Their encouragement focused on tiny actions: cut the grass, walk to the mail, cook a meal, stretch in the yard, or sit on the porch.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Agoraphobia/comments/gslhov",
        "subreddit": "r/Agoraphobia",
        "post_title": "success..... kind of?",
        "struggle_tags": "agoraphobia, panic attacks, school, self-doubt",
        "what_helped": "graduation exposure; staying through panic; therapist reframing baby steps as success",
        "support_type": "professional, practical, emotional",
        "timeframe": "single milestone",
        "risk_flags": "panic symptoms",
        "fragments": [
            "The poster graduated high school despite agoraphobia, driving to school, walking across the stage, and staying even while panicking.",
            "They felt chest tightness and exhaustion afterward, but the important part was that they did not leave before completing the goal.",
            "They struggled to count the event as success because anxiety was present, even though the goal was accomplished.",
            "Their therapist framed baby steps and imperfect successes as the road to recovery.",
            "This story is useful because it shows recovery can include panic and still be real progress.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Agoraphobia/comments/wom5il",
        "subreddit": "r/Agoraphobia",
        "post_title": "Honest question about these success stories",
        "struggle_tags": "agoraphobia, hopelessness, remission, treatment planning",
        "what_helped": "psychoeducation; behavioral change plan; exposure; third-wave CBT framing",
        "support_type": "educational, practical",
        "timeframe": "long-term",
        "risk_flags": "hopelessness",
        "fragments": [
            "A commenter described agoraphobia recovery as less about luck and more about strong psychoeducation plus a consistent behavioral change plan.",
            "The thread emphasized that avoidance keeps agoraphobia alive and that recovery asks people not to avoid the places that spark anxiety.",
            "One person identified as being in full remission and pushed back on the belief that only lucky people recover.",
            "The discussion pointed toward behavior-led CBT approaches as especially relevant for agoraphobia.",
            "This row is a research-style recovery testimony: the helpful ingredient was believing a structured plan could change outcomes.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Agoraphobia/comments/1n4f4yh/im_a_completely_recovered_agoraphobic_does_anyone/",
        "subreddit": "r/Agoraphobia",
        "post_title": "I'm a completely recovered agoraphobic does anyone need advice?",
        "struggle_tags": "agoraphobia, exposure therapy, work, driving",
        "what_helped": "exposure therapy; commuting; work routine; refusing avoidance",
        "support_type": "practical, emotional",
        "timeframe": "3 years plus 1 year recovered",
        "risk_flags": "panic symptoms",
        "fragments": [
            "The poster was agoraphobic for three years, did not leave the house, and described recovery as the hardest thing they had done.",
            "After recovering for a little over a year, they commute forty-five minutes to work every day with practically no anxiety.",
            "They had no license, job, or consistent help at the worst point, so the recovery story centers on rebuilding practical independence.",
            "Their core advice was not to let fear dictate what a person can and cannot do.",
            "In the comments, another recovering person explained that exposure helps the brain remap threat levels after doing the scary thing and surviving.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Agoraphobia/comments/1lhwskp",
        "subreddit": "r/Agoraphobia",
        "post_title": "Agoraphobia and panic disorder. Any success stories out there?",
        "struggle_tags": "agoraphobia, panic disorder, travel anxiety, emetophobia",
        "what_helped": "medication; therapy; slow exposures; allowing panic; train travel; breathing and distraction",
        "support_type": "professional, practical",
        "timeframe": "ongoing",
        "risk_flags": "medication mention, panic symptoms",
        "fragments": [
            "One commenter said they were recovering through medication, therapy, and very slow exposures.",
            "Another practiced letting panic be there during a wedding week, naming body sensations and reminding themselves anxiety would pass.",
            "A person who panicked around water practiced walking by a canal, looking at the scene, smiling, and staying instead of avoiding.",
            "One recovery update described taking a six-hour train to New York City alone, going to Times Square, shops, bars, and restaurants.",
            "The train traveler credited propranolol for reducing racing-heart symptoms enough to rebuild confidence, while also using calming techniques.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Agoraphobia/comments/1ovjrac",
        "subreddit": "r/Agoraphobia",
        "post_title": "I'm doing so well - My partial success story (SO FAR)",
        "struggle_tags": "agoraphobia, alcohol withdrawal, panic disorder, generalized anxiety",
        "what_helped": "alcohol recovery; therapy; exposure therapy research; bite-sized goals; gratitude; mindset work",
        "support_type": "professional, lifestyle, emotional",
        "timeframe": "months",
        "risk_flags": "alcohol withdrawal, panic symptoms",
        "fragments": [
            "The poster had agoraphobia for over five years and became much worse after stopping a three-year drinking problem and going through withdrawal.",
            "At the worst point they were bedbound for a week, then room-bound for months because standing caused panic attacks.",
            "They researched agoraphobia, alcohol recovery, panic disorder, generalized anxiety, and exposure therapy while already in therapy.",
            "Their practical advice was to break goals into bite-sized pieces and complete one at a time until progress starts to feel possible.",
            "They invited others to choose a small goal, such as walking around the block, getting the mail, or reaching the corner, and keep practicing until it becomes easier.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/PanicAttack/comments/1tj4iff/success_stories_overcoming_anxietypanic/",
        "subreddit": "r/PanicAttack",
        "post_title": "Success Stories Overcoming Anxiety/Panic?",
        "struggle_tags": "panic disorder, anxiety, social anxiety, depersonalization",
        "what_helped": "exposure therapy; flooding; acceptance; exercise; inner child work; support",
        "support_type": "practical, emotional, professional",
        "timeframe": "months to years",
        "risk_flags": "panic symptoms",
        "fragments": [
            "One commenter with panic disorder used exposure therapy and flooding, asking the anxiety to get worse until it lost power.",
            "Another said exposure helped because life forced them to face triggers and cope with them for months.",
            "A commenter combined years of therapy, a tailored medication routine, breathing, grounding, and acceptance that panic will rise and end.",
            "Another person used cardio, exposure therapy, and inner child work to reduce panic attacks from hours to minutes.",
            "A socially anxious commenter learned they could speak up at work by focusing on what they could control rather than how others perceived them.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/PanicAttack/comments/1szssll/how_i_overcame_my_severe_panic_attacks_and_social/",
        "subreddit": "r/PanicAttack",
        "post_title": "How I overcame my severe panic attacks and social anxiety condition - Personal Story",
        "struggle_tags": "social anxiety, panic attacks, avoidance, fear of symptoms",
        "what_helped": "acceptance; understanding fight-or-flight; social exposures; dropping resistance",
        "support_type": "educational, practical",
        "timeframe": "about 2 years",
        "risk_flags": "medication/substance mentions, panic symptoms",
        "fragments": [
            "The poster's first intense social panic attack made anxiety feel dangerous and embarrassing, and social anxiety took over their life.",
            "They stopped going to the gym, stopped seeing friends, and barely went outside for about two years.",
            "They tried many tactics but found the turning point was understanding anxiety as fight-or-flight triggered by a perceived threat.",
            "They realized the threat had become the fear of feeling anxiety itself.",
            "Recovery came from doing social situations anyway, letting anxiety and panic be present, and no longer treating symptoms as something to urgently fix.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/PanicAttack/comments/dhsety",
        "subreddit": "r/PanicAttack",
        "post_title": "I have finally overcome my Panic Disorder.",
        "struggle_tags": "panic disorder, work disruption, recovery milestone",
        "what_helped": "time; coping practice; confidence from repeated survival",
        "support_type": "practical, emotional",
        "timeframe": "months",
        "risk_flags": "panic symptoms",
        "fragments": [
            "The poster framed the post as a victory after finally overcoming panic disorder.",
            "Their first major panic attack had been severe enough to keep them out of work for an entire week.",
            "The story is useful as a milestone marker: panic can feel life-halting at first and later become something a person sees as overcome.",
            "The recovery signal was not just symptom reduction, but being able to celebrate a return of confidence.",
            "This concise account can help the prototype match people looking for proof that panic disorder can improve over time.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/PanicAttack/comments/f6vi09",
        "subreddit": "r/PanicAttack",
        "post_title": "I overcame panic attacks and anxiety years ago, now I'm helping others overcome as well! Ask me anything (AMA)",
        "struggle_tags": "panic attacks, generalized anxiety, OCD, agoraphobia, work stress",
        "what_helped": "long-term recovery; peer support; anxiety education; helping others",
        "support_type": "educational, community",
        "timeframe": "about 2 years",
        "risk_flags": "panic symptoms, OCD",
        "fragments": [
            "The poster said their spiral started with a panic attack at work.",
            "They were diagnosed with generalized anxiety disorder, panic attacks, OCD, and agoraphobia.",
            "They described the experience as a two-year downward spiral before recovery.",
            "Years later, they were using their own recovery to help others understand and overcome panic.",
            "The story is useful because it shows recovery can become peer support after someone regains stability.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Anxiety/comments/1kuix8i",
        "subreddit": "r/Anxiety",
        "post_title": "Can anybody with a success story please share it?",
        "struggle_tags": "anxiety, recovery hope, psychiatry",
        "what_helped": "psychiatrist reassurance; time; complete recovery belief",
        "support_type": "professional, emotional",
        "timeframe": "eventually",
        "risk_flags": "anxiety symptoms",
        "fragments": [
            "A commenter said they eventually recovered completely from anxiety.",
            "Their psychiatrist told them anyone can recover, which became an important hopeful frame.",
            "The story is short, but it directly answers the need for proof that anxiety recovery can happen.",
            "This row is useful for users who search for complete recovery rather than symptom management.",
            "The testimony emphasizes eventual improvement rather than a quick fix.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/HealthAnxiety/comments/hn0ntv",
        "subreddit": "r/HealthAnxiety",
        "post_title": "Success Stories",
        "struggle_tags": "health anxiety, panic attacks, dizziness, body checking",
        "what_helped": "getting a job; coworkers; meditation; healthy diet; exercise; therapy; medication",
        "support_type": "professional, lifestyle, social",
        "timeframe": "2 years",
        "risk_flags": "medication mention, health anxiety",
        "fragments": [
            "One commenter fully recovered after two years of health anxiety and later relapsed, which taught them recovery requires ongoing effort.",
            "They originally believed their physical symptoms could not possibly be anxiety because the sensations felt too real.",
            "Getting a job and having coworkers they loved helped pull attention back into life.",
            "Daily meditation, healthy eating, exercise, self-care, patience, hard work, and acceptance were part of their recovery.",
            "Another commenter improved after normal medical tests, therapy, low-dose Zoloft, psychiatry support, and exercise helped clear constant health focus.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/OCDRecovery/comments/1gj6y1i/id_like_to_hear_some_success_stories/",
        "subreddit": "r/OCDRecovery",
        "post_title": "I'd like to hear some success stories!",
        "struggle_tags": "OCD, intrusive thoughts, uncertainty, anxiety, depression",
        "what_helped": "uncertainty tolerance; ignoring compulsions; medication; exposure and CBT-based therapy",
        "support_type": "professional, educational, practical",
        "timeframe": "1 year or more",
        "risk_flags": "intrusive thoughts, medication mention",
        "fragments": [
            "One commenter said OCD is less about beating it forever and more about learning to be okay with uncertainty.",
            "They described recovery as identifying when OCD is acting up and learning to ignore thoughts and compulsions.",
            "Another person with OCD since childhood used exposure and CBT-based therapy to let intrusive thoughts pass into the background.",
            "The thread includes medication-supported recovery where obsessions became manageable and then largely disappeared.",
            "A key shared idea was that OCD may present new themes, but skills can make future flare-ups less consuming.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/OCD/comments/17idy3c",
        "subreddit": "r/OCD",
        "post_title": "Success Story",
        "struggle_tags": "OCD, mental compulsions, health anxiety, benzo withdrawal",
        "what_helped": "ERP; reducing compulsions; persistence; OCD education",
        "support_type": "professional, practical",
        "timeframe": "half a year of ERP",
        "risk_flags": "benzodiazepine withdrawal, intrusive thoughts",
        "fragments": [
            "The poster said ERP let them live life again and become mostly free of compulsions.",
            "Their compulsions had become mental, which they described as the worst OCD they had experienced.",
            "After about half a year of ERP, intrusive thoughts and anxiety headaches began reducing.",
            "They had experienced OCD since childhood, including rituals, numbers, pure OCD, just-right OCD, health anxiety, and fear of going blind.",
            "The story highlights that ERP can still help even when someone has tried many things and believes nothing works for them.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/OCDRecovery/comments/1riattu/how_i_recovered_from_ocd/",
        "subreddit": "r/OCDRecovery",
        "post_title": "How i recovered from OCD",
        "struggle_tags": "OCD, pure O, rumination, motherhood",
        "what_helped": "mindset shift; refusing self-pity; family motivation; gradual recovery",
        "support_type": "emotional, practical",
        "timeframe": "years",
        "risk_flags": "medication mention",
        "fragments": [
            "The poster was diagnosed with Pure O OCD in 2016 and said medication did not work for them.",
            "They could not access therapy easily and recovery became a slow self-directed process.",
            "A turning point came when they felt tired of having life revolve around OCD all day.",
            "They described choosing to heal because they had a family and a whole life ahead of them.",
            "This testimony is useful as a self-agency story, with a clear disclaimer that it is personal experience rather than medical advice.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/OCDRecovery/comments/1rn60op/i_have_completely_recovered_from_what_id_consider/",
        "subreddit": "r/OCDRecovery",
        "post_title": "I have completely recovered from what I'd consider extreme OCD",
        "struggle_tags": "OCD, harm OCD, health OCD, magical thinking, fear of delusion",
        "what_helped": "medication; ERP; radical acceptance",
        "support_type": "professional, practical, emotional",
        "timeframe": "not specified",
        "risk_flags": "harm OCD, medication mention",
        "fragments": [
            "The poster described complete recovery from what they considered extreme OCD.",
            "Their themes included harm OCD, health OCD, magical thinking, fear of being delusional, and seeing danger everywhere.",
            "They wanted others to hear a success story because those themes can feel especially frightening.",
            "The recovery ingredients they named were medication, ERP practice, and radical acceptance.",
            "This row is useful for users whose OCD themes feel too severe or unusual to recover from.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/BPDrecovery/comments/yvi78t/any_bpd_recovery_success_stories/",
        "subreddit": "r/BPDrecovery",
        "post_title": "any bpd recovery success stories?",
        "struggle_tags": "BPD, emotional dysregulation, trauma, relationships, self-compassion",
        "what_helped": "DBT; group therapy; self-compassion; routine; sleep; exercise; medication; journaling",
        "support_type": "professional, lifestyle, emotional",
        "timeframe": "years",
        "risk_flags": "BPD, suicidal thoughts, medication mention",
        "fragments": [
            "One commenter said their first year of therapy and DBT helped them accomplish more than years of trying alone.",
            "They emphasized wanting help and doing the work to become better to themselves and the people around them.",
            "Another person said DBT in a group setting twice helped them survive the worst period of their life.",
            "A commenter used Headspace, journaling prompts, exercise, medication, ADHD treatment, and self-compassion work.",
            "A longer recovery story described BPD, CPTSD, anxiety, and dissociation improving through supportive counseling, CBT, DBT twice, and trauma therapy.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/BPDrecovery/comments/yvi78t/any_bpd_recovery_success_stories/",
        "subreddit": "r/BPDrecovery",
        "post_title": "any bpd recovery success stories? - DBT remission comment",
        "struggle_tags": "BPD, DBT, emotion regulation, relationships, self-confidence",
        "what_helped": "3 years DBT; boundaries; self-compassion; emotional communication; therapy commitment",
        "support_type": "professional, emotional, relational",
        "timeframe": "3 years",
        "risk_flags": "BPD, alcohol mention, disordered eating mention",
        "fragments": [
            "One commenter said they recovered after three years of DBT and no longer qualified for a BPD diagnosis.",
            "They described becoming more emotionally regulated and stable, with enough confidence to follow dreams.",
            "The recovery required hard work setting boundaries and working through conflicts.",
            "Their relationship became healthier after both partners learned emotional regulation and communication skills.",
            "The turning point was moving from being scared of change to wanting recovery enough to surrender fully to therapy.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/BPD/comments/pyvazq",
        "subreddit": "r/BPD",
        "post_title": "Our success story",
        "struggle_tags": "BPD, ADHD, PTSD, anxiety, marriage",
        "what_helped": "diagnosis; DBT; medication management; eating; hydration; sunlight; walks; trusted friends; boundaries",
        "support_type": "professional, lifestyle, relational",
        "timeframe": "7.5 years",
        "risk_flags": "BPD, PTSD, medication mention",
        "fragments": [
            "A spouse shared that their husband became more stable after diagnosis and DBT.",
            "Medication consistency and staying in touch with mental health providers were counted as concrete successes.",
            "Small body-care habits such as eating, drinking water, sunlight, and walks became part of the progress.",
            "Telling two trusted friends about the diagnosis and being accepted was a major relationship win.",
            "The couple framed openness about symptoms and respect for boundaries as ongoing recovery markers.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/CPTSD_NSCommunity/comments/1d551rc",
        "subreddit": "r/CPTSD_NSCommunity",
        "post_title": "I have officially recovered from CPTSD!!",
        "struggle_tags": "CPTSD, trauma, shame, emotional processing",
        "what_helped": "5 years therapy; late-stage CPTSD recovery; assessment tools; shame processing; trauma education",
        "support_type": "professional, educational, emotional",
        "timeframe": "5 years",
        "risk_flags": "trauma, CSA mention in source context",
        "fragments": [
            "The poster celebrated that after five years of therapy, their therapist officially considered them recovered from CPTSD.",
            "They had spent the past couple years in late-stage recovery: fewer active symptoms but still processing large emotions and shame.",
            "The previous year brought major positive shifts in several areas of life.",
            "They and their therapist reviewed assessment tools that reflected recovery as well.",
            "The story stresses that trauma is not erased, but the label no longer accurately described their current life.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/CPTSD/comments/1o5t5cn/anyone_have_a_success_story_where_they_went_from/",
        "subreddit": "r/CPTSD",
        "post_title": "Anyone have a success story where they went from totally alone...",
        "struggle_tags": "CPTSD, isolation, family estrangement, support, independence",
        "what_helped": "outside help; becoming the driver of recovery; rebuilding support; refusing isolation",
        "support_type": "emotional, practical, community",
        "timeframe": "years",
        "risk_flags": "trauma, family estrangement",
        "fragments": [
            "A commenter described cutting off contact with family and friends in 2018 before rebuilding.",
            "They emphasized that they absolutely did not recover alone, but they were the driving force behind their own recovery.",
            "The story pushes back on the idea that needing help makes a recovery less valid.",
            "A core lesson was that getting on one's feet usually requires real support rather than total isolation.",
            "This row is useful for users who feel ashamed that they cannot recover entirely by themselves.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/leaves/comments/p3mv8e",
        "subreddit": "r/leaves",
        "post_title": "One year!",
        "struggle_tags": "cannabis withdrawal, anxiety, depression, anhedonia, dissociation",
        "what_helped": "one year abstinence; time; stress tolerance; creativity returning; recovery timeline",
        "support_type": "lifestyle, practical",
        "timeframe": "1 year",
        "risk_flags": "substance recovery, depression",
        "fragments": [
            "After several attempts to quit weed, the poster reached one year and estimated they were eighty to ninety percent recovered.",
            "The clearest improvement was being able to handle stress without becoming horribly dissociated and anxious.",
            "Months one to three involved constant anhedonia, depression, anxiety, and low stress tolerance.",
            "Months six to nine brought a slow return of creativity, sociability, and joy.",
            "By months nine to twelve, life felt increasingly normal with only minimal negative symptoms.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/leaves/comments/15dogke",
        "subreddit": "r/leaves",
        "post_title": "1 Year Weed & Cigarette Free Withdrawal Timeline",
        "struggle_tags": "cannabis withdrawal, nicotine, caffeine, panic attacks, depression, grief",
        "what_helped": "quit weed/caffeine/cigarettes; gym; time; symptom tracking; remembering progress",
        "support_type": "lifestyle, practical",
        "timeframe": "1 year",
        "risk_flags": "substance recovery, ER visits, panic symptoms",
        "fragments": [
            "The poster quit weed, caffeine, and cigarettes after using weed to numb grief from losing their father.",
            "They had severe withdrawal symptoms including anxiety, panic, depersonalization, brain fog, heart palpitations, dizziness, and depression.",
            "They went to the ER three times thinking something was wrong with their heart, but monitoring showed nothing.",
            "After three months symptoms slowly tapered, and at five months returning to the gym helped remaining symptoms disappear except mild waves.",
            "At one year, symptoms were basically gone and they could recognize mild waves without being consumed by them.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/leaves/comments/x6yvgj",
        "subreddit": "r/leaves",
        "post_title": "I quit weed and it has been the best decision of my life",
        "struggle_tags": "cannabis addiction, depression, anxiety, social isolation",
        "what_helped": "asking for help; AA; 12 steps; sober living; honesty; consistency; social effort",
        "support_type": "community, lifestyle, emotional",
        "timeframe": "10 months",
        "risk_flags": "suicidal depression, substance recovery",
        "fragments": [
            "Before quitting, the poster was suicidally depressed, extremely anxious, socially isolated, and smoking weed daily.",
            "Ten months later they described no severe depression or severe anxiety and a return of passion, spirituality, work, and dating.",
            "They asked for help and threw away everything related to weed.",
            "AA meetings, the twelve steps, sober living, taking suggestions, honesty, and consistency were part of the recovery structure.",
            "The poster named realizing they were addicted to weed as a crucial step.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/leaves/comments/z63r0y",
        "subreddit": "r/leaves",
        "post_title": "One year weed free",
        "struggle_tags": "cannabis use, depression, anxiety, paranoia, abusive relationship",
        "what_helped": "quitting cannabis; clarity; social reconnection; freelancing; clean home; confronting avoidance",
        "support_type": "lifestyle, practical, emotional",
        "timeframe": "1 year",
        "risk_flags": "substance recovery, abusive relationship mention",
        "fragments": [
            "After twenty years of smoking and ten years of daily use, the poster quit for a year.",
            "Quitting gave them clarity and confidence to leave an abusive relationship and quit a toxic job.",
            "They became more successful freelancing because they could set schedules and actually stick to goals.",
            "They described socializing as easier because conversations and connection became possible again.",
            "The paranoia disappeared, their apartment became clean, and they could care for plants they used to neglect.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/stopdrinking/comments/wzk5fs",
        "subreddit": "r/stopdrinking",
        "post_title": "1 year sober",
        "struggle_tags": "alcohol, binge drinking, depression, suicidal thoughts",
        "what_helped": "sobriety; recognizing post-drinking depression; resisting social pressure",
        "support_type": "lifestyle, emotional",
        "timeframe": "1 year",
        "risk_flags": "suicidal thoughts, alcohol recovery",
        "fragments": [
            "The poster had struggled with alcohol since their teenage years, especially binge drinking.",
            "They had quit for a year twice before, then slowly slid back from occasions to weekends to binge drinking.",
            "They noticed alcohol made them feel great briefly but burned through happiness and left them very low for weeks.",
            "After one last night of drinking led to weeks of misery and suicidal thoughts, they stopped.",
            "At one year sober, they still faced work and social issues from not drinking, but sobriety was the line they chose.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/depression_partners/comments/1hy5p4r",
        "subreddit": "r/depression_partners",
        "post_title": "Success stories",
        "struggle_tags": "major depression, marriage, caregiver burnout, boundaries",
        "what_helped": "coping tools; favorite sports; friends; laughter; boundaries; rebalanced responsibilities",
        "support_type": "relational, emotional, practical",
        "timeframe": "about 5 years",
        "risk_flags": "emotional abuse mention, depression",
        "fragments": [
            "A spouse described their marriage recovering after the husband developed major depressive disorder about five years earlier.",
            "For nearly a year, depression showed up as anger, blame, criticism, and refusal to get help.",
            "During recovery, they rebuilt the relationship through favorite sports, time with friends, and finding ways to laugh again.",
            "The partner had to set firm boundaries and stop falling fully into the caregiver role.",
            "Today the husband is mostly back to himself, with added empathy and self-awareness from the fight with depression.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/depression_help/comments/103m0uf",
        "subreddit": "r/depression_help",
        "post_title": "Success stories",
        "struggle_tags": "depression, anxiety, leaving the house, triggers",
        "what_helped": "coping skills; functional goals; perspective; pushing through triggers",
        "support_type": "practical, emotional",
        "timeframe": "not specified",
        "risk_flags": "depression, anxiety symptoms",
        "fragments": [
            "A commenter said they went from being unable to leave the house because of anxiety and depression to becoming functional.",
            "They described being successful in some areas while still working on others, like anyone else.",
            "Triggers can still set them off, but coping skills help them push through.",
            "Their perspective changed as they got older and realized everyone is trying to function with their own difficulties.",
            "This story is useful because it frames success as functioning and improving, not becoming perfect.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Psychosis/comments/1cdgvjc/recoverysuccess_story/",
        "subreddit": "r/Psychosis",
        "post_title": "Recovery/Success Story :)",
        "struggle_tags": "psychosis recovery, depression, anxiety, shame, social reconnection",
        "what_helped": "time; community; journaling; art; returning to work; social situations; active recovery",
        "support_type": "professional, creative, community",
        "timeframe": "about 1 year",
        "risk_flags": "psychosis, medication mention, trauma",
        "fragments": [
            "The poster returned to the psychosis community to say they had recovered after previously having no hope.",
            "At the worst point they felt like a zombie on medication and stayed alive mainly to spare loved ones more trauma.",
            "A year later they were sociable, hopeful, growing, and finding meaning in what happened.",
            "They described recovery milestones: leaving hospital, coming off medication with care, getting a normal cafe job, and returning to the place where the episode occurred.",
            "They recommended journaling, making art, seeking community, and being active in recovery rather than waiting passively.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Psychosis/comments/1q4fd4q/recovery_success_stories_please/",
        "subreddit": "r/Psychosis",
        "post_title": "Recovery success stories please",
        "struggle_tags": "psychosis recovery, rumination, anxiety, nightmares, work return",
        "what_helped": "medication; hospital care; step-down care; work; time; rest; exercise; sobriety",
        "support_type": "professional, lifestyle, emotional",
        "timeframe": "5 months to 2 years",
        "risk_flags": "psychosis, medication mention, cannabis recovery",
        "fragments": [
            "One commenter had a seven-month episode and said treatment helped them become employed, dating, housed, and mostly symptom-free within months.",
            "They still had anxiety and nightmares, showing recovery can include remaining symptoms.",
            "Another person said their first terrible psychosis episode lasted five days, with cognitive recovery taking six to eight months.",
            "A longer reply emphasized time, rest, and taking things slowly because the brain went through something intense.",
            "One person nearly two years out said they were working, meditating, spending time with their dog, and grateful to be alive after thinking recovery would never come.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Psychosis/comments/1gyb6ee",
        "subreddit": "r/Psychosis",
        "post_title": "Success stories",
        "struggle_tags": "psychosis, post-psychotic depression, medication, drug-induced episode",
        "what_helped": "psychiatrist; medication; time; avoiding drugs; gentle self-talk; recovery stories",
        "support_type": "professional, community, emotional",
        "timeframe": "not specified",
        "risk_flags": "psychosis, medication mention, drug-induced psychosis",
        "fragments": [
            "The poster said psychosis success stories acted like sparks when they were fresh out of the psych ward.",
            "Reading every comment where people described recovery helped them feel hope again.",
            "Their message to others was that recovery is possible, but it needs time and should not be rushed.",
            "They advised being gentle, listening to a psychiatrist, taking medication, and waiting.",
            "For drug-induced psychosis, they emphasized avoiding drugs to lower relapse risk.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Psychosis/comments/1kl6e4d",
        "subreddit": "r/Psychosis",
        "post_title": "How many of you are high functioning? Tell me your success stories",
        "struggle_tags": "psychosis recovery, work, exercise, routine, socializing",
        "what_helped": "medication; work; exercise; socializing; gardening; writing; routine",
        "support_type": "professional, lifestyle, practical",
        "timeframe": "2020 to 2025",
        "risk_flags": "psychosis, medication mention",
        "fragments": [
            "A commenter had severe psychosis in 2020, a milder one in 2022, and was recently described by their doctor as fully recovered.",
            "They remained on medication as part of a long-term plan but described themselves as high-functioning.",
            "They had bad days like anyone else but were generally doing great.",
            "Recent wins included exercise, a raise at work, socializing, plans for a garden, and writing an autobiography.",
            "The story is useful because it shows recovery can include work, creativity, and long-term medication management.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Psychosis/comments/1mivmaq",
        "subreddit": "r/Psychosis",
        "post_title": "Fully Recovered",
        "struggle_tags": "psychosis recovery, school disruption, low motivation, medication",
        "what_helped": "psych ward care; medication; therapy; time; reading recovery posts",
        "support_type": "professional, community",
        "timeframe": "about 16 months",
        "risk_flags": "psychosis, medication mention",
        "fragments": [
            "The poster had bad psychosis about sixteen months earlier and had to miss a year of school.",
            "They went through the psych ward, medication, therapy, and a long recovery process.",
            "They used to check the subreddit daily to stay optimistic and read every relevant recovery post and comment.",
            "By the update, they considered themselves fully recovered and no longer needed to check the community often.",
            "They remembered sleeping more than twelve hours a day with no motivation, but that phase eventually passed.",
        ],
    },
    {
        "source_url": "https://www.reddit.com/r/Psychosis/comments/w74bjd",
        "subreddit": "r/Psychosis",
        "post_title": "Success Stories Please",
        "struggle_tags": "psychosis recovery, trauma, suicidal ideation, work, education",
        "what_helped": "weekly therapy; support system; coping skills; routine; work; college",
        "support_type": "professional, community, practical",
        "timeframe": "about 10 years",
        "risk_flags": "suicidal thoughts, trauma, psychosis",
        "fragments": [
            "A commenter had a first psychotic episode before sixteen after ongoing sexual abuse and multiple hospitalizations.",
            "They were given a poor prognosis but later graduated college with good grades.",
            "They became a teacher, were approaching tenure, and bought their own condo.",
            "They still had manageable passive suicidal ideation, hallucinations, and paranoia, but life became relatively normal.",
            "Weekly therapy, friends, coping skills, and a regular routine helped keep symptoms manageable.",
        ],
    },
]

bulk_rows = []
next_story_id = 83
for thread in []:
    for fragment_index, fragment in enumerate(thread["fragments"], start=1):
        bulk_rows.append(
            extension_row(
                f"reddit_original_{next_story_id:03d}",
                thread["source_url"],
                thread["subreddit"],
                "bulk_comment_or_post_excerpt",
                thread["post_title"],
                fragment,
                thread["struggle_tags"],
                thread["what_helped"],
                thread["support_type"],
                thread["timeframe"],
                "bulk excerpt, recovery-focused",
                thread["risk_flags"],
                f"Bulk expansion row {fragment_index} from a Reddit recovery thread; source should be re-reviewed before production use.",
                "visible_via_search_result_or_direct_open_excerpt",
                "needs_human_review",
            )
        )
        next_story_id += 1

ROWS.extend(bulk_rows)


FULL_TEXT_ADDITIONAL_ROWS = [
    extension_row(
        "reddit_original_083",
        "https://www.reddit.com/r/HealthAnxiety/comments/hn0ntv",
        "r/HealthAnxiety",
        "post",
        "Success Stories",
        """I would love to hear a success story about anyone who’s struggled with health anxiety and has been able to overcome this. There was one point in my life were I wasn’t so afraid of getting sick or having some serious terminal illness. A time where I used to enjoy life and not obsess about health....Ever since I had to get a baseline mammogram I’ve had this fear of getting ill and I keep wanting to check my body to catch things early. I wish I could get back to that time where I wasn’t always worried. Are there any positive success stories for people who have struggled with health anxiety but can now live life happily and healthy without worrying all day? Hearing that would be so helpful for me so I know that I will move beyond this daily fear one day.""",
        "health anxiety, illness fear, body checking",
        "seeking success stories; hope; wanting to move beyond daily fear",
        "emotional",
        "ongoing",
        "vulnerable, seeking hope",
        "health anxiety",
        "Question post included because it captures the user need for recovery stories.",
    ),
    extension_row(
        "reddit_original_084",
        "https://www.reddit.com/r/HealthAnxiety/comments/hn0ntv",
        "r/HealthAnxiety",
        "comment",
        "Success Stories",
        """After suffering for 2 years I fully recovered. 100% back to my normal self. Unfortunately since the end of April I relapsed and have been in a rut. But it’s taught me a lot about my anxiety and how it’s a constant effort, and not a one and done kind of thing. But I will tell you that I was one of those people before I recovered the first time that thought it was absolutely impossible. I also thought that there was just no way my symptoms were caused by anxiety. They were too real. They never went away, even when I wasn’t feeling anxious. But they were anxiety, as I know now. I recovered because I finally forced myself to get a job, and got a group of coworkers that I loved. I also practiced daily mediation, and ate a healthy diet. Exercised everyday. Basically took care of myself to the best of my ability. This whole quarantine is what sparked my HA again. Not because of the virus, I’m not scared of it at all tbh...but because I have too much time to sit and think at home since I got laid off. My advice is to find what works for you, and be kind to yourself. It takes patience, a lot of hard work, and acceptance. But you can do it. I did, and a lot of other people have. You and me both will recover and be smiling one day knowing we did""",
        "health anxiety, relapse, physical symptoms, work",
        "job; supportive coworkers; daily meditation; healthy diet; exercise; patience; acceptance",
        "lifestyle, social, emotional",
        "2 years",
        "realistic, hopeful",
        "health anxiety, relapse",
        "Full visible recovery comment from search/open result.",
    ),
    extension_row(
        "reddit_original_085",
        "https://www.reddit.com/r/HealthAnxiety/comments/hn0ntv",
        "r/HealthAnxiety",
        "comment",
        "Success Stories",
        """The symptoms when you don’t feel anxious thing is the reason I feel like most can’t recover, and it makes sense. Your mind builds the picture that if the symptoms happen even when you’re not worried, then it has to be real, as that’s what a lot of people even tell us. My major symptom when I first had health anxiety that I couldn’t get passed was dizziness. I was OBSESSED with my balance. I would be off balance from the moment I woke up to when I went to sleep, and felt completely out of my body. I always told myself that there is no possible way this is all in my head. Then after two years I just got sick of it and said “fuck it, let it kill me. or let me be stuck like this. But I need to live”. It stayed like that for another 3-4 months or so. Then one day I noticed that I hadn’t felt dizzy for weeks. It wasn’t like it went away one morning. I guess it just slowly went away and I was so involved in enjoying life that I didn’t even notice""",
        "health anxiety, dizziness, balance obsession, depersonalization",
        "acceptance; living despite symptoms; stopping symptom monitoring",
        "practical, emotional",
        "2 years plus 3-4 months",
        "blunt, experiential",
        "health anxiety, physical symptoms",
        "Full visible follow-up comment about dizziness and acceptance.",
    ),
    extension_row(
        "reddit_original_086",
        "https://www.reddit.com/r/HealthAnxiety/comments/hn0ntv",
        "r/HealthAnxiety",
        "comment",
        "Success Stories",
        """At the end of last year into this year, I had multiple panic attacks about my health. I was focused on my heart from random palpitations, heart rate slowness, high BP, tingling, etc. It was horrible. I felt miserable and couldn't focus on the present moment. Even if I could focus on the present for 2-3 min, my mind would eventually drift back to if everything was ok. I was convinced I had some sort of issue.
I first went to my doctor to check to make sure everything was ok - they took EKG, Echo, and blood. All normal. I felt better after the tests, but after those came back, I had a headache and was convinced I had a brain tumor. This is when I realized I had a problem with my anxiety and needed to fix it.
I went to therapy and started a low dose of Zoloft (25mg). After a few weeks, it really helped me and I started to feel much better. Head is clear and I don't focus on my health 24/7. I'm still on it and working with my psychiatrist, but hoping to maybe get off it soon and see how that goes. I'm 100% happier, calmer, and my head is clear.
I'm not saying SSRI is for you, but worth checking out. I wouldn't have gone straight to that if I knew I hadn't had an anxiety filled past. Exercise is worth a try since it helps me a lot too.""",
        "health anxiety, panic attacks, heart fears, brain tumor fear",
        "medical tests; therapy; low-dose medication; psychiatrist; exercise",
        "professional, lifestyle",
        "weeks to months",
        "specific, practical",
        "medication mention, health anxiety",
        "Full visible comment; needs medical disclaimer.",
    ),
    extension_row(
        "reddit_original_087",
        "https://www.reddit.com/r/socialanxiety/comments/1k64a5r",
        "r/socialanxiety",
        "post",
        "My story of how I cured from social-anxiety (and keep going every day!)",
        """I will tell you my experience as a person that had a very deep social-anxiety.

Before I'll start with the "success story", I'll start with how my life looked like before I overcame my fears.

I couldn't look in the eyes of others. Everything I did or said felt "Wrong", "Weird", "Weak"...   
I was afraid of people judging me and it made people judge me even more. I've been judged or even bullied by almost every person I met. (I had some terrible social circle)  
Every bad feedback I got made me locked-inside even more.  
I was even on watch for actual medicine since I've started to develop obsessive thoughts.(nothing harmful, just non-stop thinking of why I might not succeed...)

I've tried basically everything, looked for that "Magic" solution that'll make me confident, I thought I had to "become confident" in order to not GAF, and that was the trap.  
I've been waiting for that "magically confident" cure to come and heal me, and nothing changed for years.

I've realized the ONLY way to cure my fear is through the fear itself.""",
        "social anxiety, bullying, obsessive thoughts, confidence",
        "facing fear directly; dropping search for magic confidence",
        "practical, emotional",
        "years",
        "reflective, motivational",
        "bullying, obsessive thoughts",
        "Visible portion of social anxiety recovery post.",
    ),
    extension_row(
        "reddit_original_088",
        "https://www.reddit.com/r/Meditation/comments/1inbqmo/success_stories_overcoming_anxiety/",
        "r/Meditation",
        "post",
        "Success stories overcoming anxiety?",
        """Im in my late 30’s and have just started meditating. 20 minutes in the morning for the past 6 months. I sit in silence and come back to my breath. Haven’t noticed it helping much, but I’m hoping that I’m watering seeds and just can’t see them sprouting. Have mostly taken care of worry and anxiety with drugs and alcohol for past 20 years. 8 months sober and trying things different. Anxiety is crippling a lot of the time. I really want to get out of the northern cold for a month or two and stay in the Caribbean, but am worried I will freeze up and succumb to substance use if heavy anxiety sets in. I also attend a variety of meetings, listen to different spiritual teachers like Michael singer, eckhart tolle, sadguru, Krishnamurti, etc… I’m still early in sobriety and learning to be still, so trying to be patient, but feeling like I can’t even take a trip like I used to is depressing. Looking for people who overcame major anxiety through meditation, awareness, presence. I know everyone’s journey will be different, but I always love a good success story.""",
        "anxiety, sobriety, meditation, substance use",
        "meditation; meetings; spiritual teachers; patience in sobriety",
        "spiritual, community, lifestyle",
        "6-8 months",
        "seeking, honest",
        "substance recovery, anxiety",
        "Question post; useful for matching sobriety plus meditation needs.",
    ),
    extension_row(
        "reddit_original_089",
        "https://www.reddit.com/r/Meditation/comments/1inbqmo/success_stories_overcoming_anxiety/",
        "r/Meditation",
        "comment",
        "Success stories overcoming anxiety?",
        """I had crippling social anxiety that was triggered by LSD, I had never learnt to deal with my emotions so I never understood where all my insecurities came from and it was very overwhelming, I couldn’t even go to uni and I lost my job thanks to it. This was about the time that the pandemic hit so I got the chance to start slowly incorporating myself into society, I tried everything from therapy, antidepressants , Buddhist groups, you name it and nothing could work, I thought for a brief period of time that I was honestly loosing my intelligence and it was pretty scary since I always thought that the only thing good in me was my intelligence
I think I got to experience hitting rock bottom emotionally, at least for what I can handle (I know brave people who suffer from truly horrible personal experiences hundreds of times worse than mine and my hearth goes out to them) and there what pushed me through was compassion for myself and as I started forgiving myself it snowballed into a lot of love and a journey into a spiritual awakening, through my thoughts not only did I pushed through depression (which almost killed me) but I created an extroverted, abundant life with literally my dream job from back then.
The way through is always love.""",
        "social anxiety, depression, job loss, university, substance-triggered anxiety",
        "self-compassion; self-forgiveness; spiritual awakening; gradual social reintegration",
        "spiritual, emotional",
        "pandemic period",
        "intense, spiritual, hopeful",
        "drug mention, depression, suicidal implication",
        "Full visible comment from meditation thread.",
    ),
    extension_row(
        "reddit_original_090",
        "https://www.reddit.com/r/Meditation/comments/1inbqmo/success_stories_overcoming_anxiety/",
        "r/Meditation",
        "comment",
        "Success stories overcoming anxiety?",
        """I’ve honestly always had awful anxiety and I recently started meditating 
ill say that yes it has worked to help I guess free space in the mind not sure how to word it but I’ll add that my anxiety has shown in chest pain in heart palpitations shortness of breath and the things that helped with that was just doing what made them come 
Id purposely say let me fkin go do this thing out the house because I feel nervous about doing it! I remember a time when I was so chill and didn’t give a fuck (prescribed Xanax) HAHA and I’d try replicate how it felt to not give a fuck and not be anxious and how i enjoyed myself out in public without caring so much and it helped me to put myself in those situations as much as i could, id even pretend and hold that feeling that someone safe was with me like my sister, id often leave those outings feeling great ful that i did it and knowing it wasn’t half as bad as my mind thought it was, and that there’s such an initial build up of anxiety it’s just so important to get out of our heads sometimes 
I think knowing that your mind is making up bs and proving yourself wrong over and over again is a really great teacher 
But I think meditation has helped massively aswell but it’s definitely hand in hand with actually experiencing the real world and the things that make us uncomfortable""",
        "anxiety, chest pain, palpitations, avoidance",
        "meditation; exposure; leaving the house; imagined safe person; proving anxious thoughts wrong",
        "practical, lifestyle",
        "recent",
        "casual, practical",
        "benzodiazepine mention, physical anxiety symptoms",
        "Full visible comment; informal language preserved.",
    ),
    extension_row(
        "reddit_original_091",
        "https://www.reddit.com/r/Meditation/comments/1inbqmo/success_stories_overcoming_anxiety/",
        "r/Meditation",
        "comment",
        "Success stories overcoming anxiety?",
        """I have an anxiety disorder which is greatly alleviated by meditation. When I meditate consistently, the change and the relief is dramatic.""",
        "anxiety disorder, meditation",
        "consistent meditation",
        "lifestyle, spiritual",
        "ongoing",
        "brief, clear",
        "anxiety",
        "Short but complete visible comment.",
    ),
    extension_row(
        "reddit_original_092",
        "https://www.reddit.com/r/PanicAttack/comments/dhsety",
        "r/PanicAttack",
        "post",
        "I have finally overcome my Panic Disorder.",
        """Hi all. Celebrating a victory on my part as I have finally overcome my Panic Disorder.


I was diagnosed with PTSD-C, GAD, ADD and Panic disorder. I can, with confidence, scratch off Panic Disorder on my list and soon to follow will be GAD.


I am posting to shine hope on those who think they'll never get out of the darkness that is Anxiety/panic attacks. I've posted here before, asking for help and I've been in all your shoes.


This past May I experienced my first major (10/10) panic attack that left me out of work for an entire week. I had all the symptoms, to the point my anxiety was so high I would vomit. My panic attacks were so bad I had 3 a day, I didn't sleep for 36hrs, I lost over 15lbs, went to the ER twice, and saw my doctor atleast 10 times. I would go to her ask "What's wrong with me??" Leaving stumped when it wasn't a legit medical condition. I've been there, felt that.


There's Hope for all of you. You just have to believe in yourself. I believe in you.""",
        "panic disorder, PTSD-C, GAD, ADD, ER visits, work absence",
        "hope; persistence; doctor visits; belief in recovery",
        "professional, emotional",
        "about 5 months",
        "celebratory, hopeful",
        "panic symptoms, weight loss, ER visits",
        "Visible full post text from search/open result.",
    ),
    extension_row(
        "reddit_original_093",
        "https://www.reddit.com/r/Anxiety/comments/1ktr2yq",
        "r/Anxiety",
        "post",
        "Finally I overcame anxiety",
        """Hi everyone,
I just wanted to share a success story here, because this sub helped me so much during the hardest moments of my life. When I was struggling, reading the posts here gave me hope, so maybe my story will do the same for someone else.

My first panic attack happened when I was around 16. I struggled with anxiety on and off ever since (mostly due to high school stress) but I managed to cope without any medication or therapy.
That changed in May 2023, when I had a full-blown panic attack that was so intense I ended up in the ER. That wasn't the first time but this time I really thought I was dying. The doctors told me it was "just" a panic attack, but that traumatic experience changed my life.
From that day on, I felt constant anxiety. I was terrified I would die. Every single day, I experienced extreme physical symptoms: dizziness, headaches, chest pain, racing heart, breathlessness, derealization, depersonalization – you name it.
I developed health anxiety. Every single day I was convinced I had cancer, a hidden heart disease, a brain tumor, that I would have a stroke any moment.""",
        "anxiety, panic attacks, health anxiety, derealization, depersonalization",
        "reading recovery posts; ER reassurance; eventually overcoming anxiety",
        "emotional, professional",
        "since age 16; major episode in 2023",
        "long-form, hopeful",
        "ER visit, health anxiety, panic symptoms",
        "Visible portion of high-vote anxiety success post.",
    ),
    extension_row(
        "reddit_original_094",
        "https://www.reddit.com/r/socialanxiety/comments/1k1ti3i",
        "r/socialanxiety",
        "post",
        "I spent 10 years doing exposure therapy and recorded most wins/losses",
        """I started doing exposure therapy and stuck with it for 10 years. I'm a big journaler, so I also ended up writing down stories of my wins and demoralizing losses -- in detail.

Ask me anything about exposure therapy, facing fear, setbacks or building confidence.

I’m happy to share what helped me (and what didn’t).""",
        "social anxiety, exposure therapy, confidence",
        "10 years exposure therapy; journaling wins and losses",
        "practical",
        "10 years",
        "concise, open",
        "social anxiety",
        "Full visible post text.",
    ),
    extension_row(
        "reddit_original_095",
        "https://www.reddit.com/r/socialanxiety/comments/1k1ti3i",
        "r/socialanxiety",
        "comment",
        "I spent 10 years doing exposure therapy and recorded most wins/losses",
        """Thanks! Honestly I still have things that make me anxious, but I think I've come a long way since then.
**The biggest thing that helped me:** keeping a journal of my wins and reading it often. Every time I overcame a fear, I'd write down what I did and why it mattered. You'd be amazed how much confidence you get from revisiting your own wins. Social anxiety wires us to focus on future failure — but our past wins fuel present confidence. I have a journal of 10 years of wins to scroll through now :)
**My best success story:** after six years of knowing someone, I finally shared one of the most pivotal stories of my life. I had thought about it for *years* but always backed out, afraid she'd think I was weird and it would ruin our relationship. It was a long story too, so I had a lot of anxiety about sitting through it without rushing or quitting halfway. But on Feb 23, 2020, I finally did it — and it actually brought us closer. I spent the next two months riding that high. One of the best moments of my life.
**The worst experience**: not opening up to a neighbor I had a crush on. I never thought we were a great match long-term, but because I was too scared to talk about anything personal, our relationship stayed surface-level. She shared real parts of herself with me, and I wanted to do the same — about my projects, my faith, my struggles — but most of the time, I couldn't.""",
        "social anxiety, vulnerability, confidence, journaling",
        "journal of wins; repeated exposure; opening up to someone trusted",
        "practical, emotional",
        "10 years",
        "specific, reflective",
        "social anxiety",
        "Full visible answer comment with success and setback.",
    ),
    extension_row(
        "reddit_original_096",
        "https://www.reddit.com/r/leaves/comments/p3mv8e",
        "r/leaves",
        "post",
        "One year!",
        """After several attempts to quit weed for good, I've finally hit the one year mark. At this point I would estimate that I'm about 80-90% recovered. I still have some other bad habits to deal with (periodically when these habits are under control I feel so peaceful I can hardly believe it). The clearest difference between then and now is in my ability to handle stress. In the beginning, if I slept poorly, drank too much or got into an argument (to name a few), I would feel horribly dissociated and anxious. These days almost nothing brings me anxiety. I feel much more grounded and creative. I have no doubt that at some point during the next year I will feel as if I had never smoked!

A rough timeline of my recovery:

months 1-3: constant anhedonia, depression, anxiety, low tolerance to stress / fragile nervous system

months 3-6: above symptoms peaked at the 3 month mark. The worst was over by the 6 month mark.

months 6-9: slow return of creativity, sociability and joy.

months 9-12: increased sense of normalcy, only minimally and periodically affected by any negative symptoms.

In retrospect, this seems to be a pretty typical timeline for a post-acute withdrawal syndrome, with the majority making a full recovery within 1-2 years. Since all my previous attempts to quit failed at me smoking "just once", I've decided that I will not ever smoke again. Earlier this decision felt like a loss, but at this point it feels like it doesn't make any difference (in a good way). Thank you for reading.""",
        "cannabis withdrawal, anxiety, depression, anhedonia, dissociation",
        "abstinence; timeline tracking; stress tolerance; never smoking again",
        "lifestyle, practical",
        "1 year",
        "timeline, hopeful",
        "substance recovery, depression",
        "Full visible post text.",
    ),
    extension_row(
        "reddit_original_097",
        "https://www.reddit.com/r/stopdrinking/comments/117h3vy",
        "r/stopdrinking",
        "post",
        "1 year sober!",
        """1 year ago I was feeling the following:

* No appetite and little sleep.
* Fatigue and low energy.
* Feeling sad, empty, and anxious.
* Feeling guilty, worthless, and helpless.
* Feeling hopeless and pessimistic.
* Having trouble remembering things.
* Irritable
* No interest in activities I enjoyed doing.
* Aches and pains.
* Numerous health problems.

Alcohol helped relieve all of what I just mentioned……until it didn’t. It created a monster within.

“I took a drink, the drink took a drink, and the drink took me.”

1 year ago I had finally had enough of self medicating in search for a temporary escape or the ‘quick fix.’

My life had become unmanageable.

1 year ago I chose Recovery.

“Recovery is not simple abstinence. It’s about healing the brain, remembering how to feel, learning how to make good decisions, becoming the kind of person who can engage in healthy relationships, cultivating the willingness to accept help from others, daring to be honest, and opening up to doing.”""",
        "alcohol, anxiety, depression, self-medication",
        "recovery; abstinence; honesty; accepting help; healthier relationships",
        "community, emotional",
        "1 year",
        "reflective, recovery-focused",
        "alcohol recovery, depression symptoms",
        "Full visible post text.",
    ),
    extension_row(
        "reddit_original_098",
        "https://www.reddit.com/r/leaves/comments/1tpa3ap/anxiety/",
        "r/leaves",
        "post",
        "Anxiety?!?!",
        """I've been using weed to mask my anxiety. It helps slow my thoughts and aids in avoiding the things that stress me.
When I've quit before (currently attempting again bc addiction ya know) I did have an increase in my anxiety, since I wasn't able to self medicate with weed.
However, using weed did NOT actually improve my anxiety. It just helped me avoid it. You enter back into the "real world" when you quit, and are faced with the serious/scary nature of life you originally sought to avoid.
I would recommend seeing a psychiatrist and going to therapy for it, as those are the ways you can actually address it. Of course it will be uncomfortable since you have to face it head on. But the grass IS greener on the other side.
Edit: I should add that at a certain point the weed stopped helping with anxiety, as I wasn't actually addressing it or dealing with the things that make me stressed. Overcoming anxiety is largely about desensitizing your brain to the things that instigate it through exposure.""",
        "cannabis use, anxiety, avoidance, self-medication",
        "therapy; psychiatrist; facing stressors; exposure",
        "professional, practical",
        "ongoing quitting attempts",
        "honest, advisory",
        "substance recovery, anxiety",
        "Full visible thread text.",
    ),
    extension_row(
        "reddit_original_099",
        "https://www.reddit.com/r/leaves/comments/1tpa3ap/anxiety/",
        "r/leaves",
        "comment",
        "Anxiety?!?!",
        """Work issues... Anxiety. You get the point.
One year sober... What even is anxiety? After a few months sober, I started to notice the changes. Things just stopped affecting me like they used it. It's been a huge relief not getting anxious over everything like I once did and my confidence is coming back slowly but surely.
Don't let other peoples experience scare you. Could yours get worse? Maybe. Could it get better? Maybe. You will never know until you get there and should it get worse, seek professional help. Weed certainly won't fix it but maybe talking with a professional can.""",
        "cannabis use, anxiety, confidence, work stress",
        "sobriety; time; professional help if needed",
        "lifestyle, professional",
        "1 year",
        "reassuring, grounded",
        "substance recovery",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_100",
        "https://www.reddit.com/r/leaves/comments/1tpa3ap/anxiety/",
        "r/leaves",
        "comment",
        "Anxiety?!?!",
        """I take medication for anxiety and depression.   
I think weed was interfereing with it because it just felt like it was not working.   
Tried several doses, types, etc.
My (doctor) brother told me I need to quit. I quit, cold turkey.   
First 2-3 weeks anxiety was worse than before, like panic-attacks-crawl-into a-ball bad.   
I knew, from everything I read, that it was common.
2 months in, my anxiety is still there, but it is SOOOO much better than it has been the past year.   
I am so glad I took his advice.""",
        "cannabis use, anxiety, depression, panic attacks",
        "quitting weed; medication; doctor advice; waiting through withdrawal",
        "professional, lifestyle",
        "2 months",
        "relieved, practical",
        "medication mention, substance recovery, panic attacks",
        "Full visible comment; medical disclaimer needed.",
    ),
    extension_row(
        "reddit_original_101",
        "https://www.reddit.com/r/leaves/comments/1tpa3ap/anxiety/",
        "r/leaves",
        "comment",
        "Anxiety?!?!",
        """I’m nearly a month in and my anxiety has settled down significantly but it was really bad for the first couple weeks. Granted I have an anxiety disorder already and have been going through some stressful times, but quitting made it all kind of hit me at once. Pounding heart, sleep troubles, feeling like I was going CRAZY.
Some things I did to help: cut caffeine, use the “Quit Weed” app to help track my symptoms and realize “this sucks but it’s normal and will be over eventually, lots of breathing exercises, and learning to catch myself when I realize I’m spiraling.
It’s really hard but it does get easier! You can do it!""",
        "cannabis withdrawal, anxiety disorder, sleep trouble, panic symptoms",
        "cut caffeine; Quit Weed app; breathing exercises; catching spirals",
        "digital tool, lifestyle, practical",
        "nearly 1 month",
        "practical, encouraging",
        "substance recovery, anxiety symptoms",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_102",
        "https://www.reddit.com/r/leaves/comments/1tpa3ap/anxiety/",
        "r/leaves",
        "comment",
        "Anxiety?!?!",
        """as time went on, these periods became less frequent, but the intensity is still higher now that i’m not smoking. but, my baseline is still much lower.
It’s also hard to say if my anxiety got more intense, or if taking away my major and in fact only coping skill for anxiety made it more difficult to manage.
On the whole, a little more than a year out from quitting, I’d say quitting made my anxiety better — with the nuance of “much lower baseline anxiety, but intense periods of anxiety right after i quit”""",
        "cannabis use, anxiety, coping skills",
        "long-term sobriety; tolerating intense early anxiety; baseline tracking",
        "lifestyle, reflective",
        "a little over 1 year",
        "nuanced, realistic",
        "substance recovery",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_103",
        "https://www.reddit.com/r/leaves/comments/f6za8n",
        "r/leaves",
        "post",
        "1 year clean, still have withdrawals (post acute withdrawal)",
        """Wow, so one year ago I quit smoking weed. Its been a really shit, terrible year and Ive suffered a lot but I can see the light at the end of the tunnel now.  Ive been going through post acute withdrawal, which I think people on here are probably familiar with.  I posted this on another sub, which is a sub just for post acute withdrawal from weed.  Just thought id post it here to share; in case others are suffering like I have been.

I used to smoke a lot. For 5 years - the first 2 years just a small bowl every night; then increased that a lot the next two years; I smoked a few bowls a day. By the end of that I was smoking 1-3 grams a day. There was a time I quit once; and I felt mild paws symptoms - headache, insomnia, depression, anxiety; but all very mild. Then I started smoking dab pens and concentrates for a year or two - thats when the health symptoms I observed went off the chain. They were things like: headache, muscle aches, poor sleep (had to smoke multiple times a night to sleep), erectile dysfunction, depression, constantly tired.""",
        "cannabis withdrawal, PAWS, depression, anxiety, insomnia",
        "quitting weed; sharing PAWS experience; seeing light at end of tunnel",
        "community, lifestyle",
        "1 year",
        "raw, detailed",
        "substance recovery, depression, sexual health mention",
        "Visible portion of post.",
    ),
    extension_row(
        "reddit_original_104",
        "https://www.reddit.com/r/stopdrinking/comments/1cy349k",
        "r/stopdrinking",
        "post",
        "1 year sober",
        """Officially made it to 1 year sober today! I’ve struggled with sobriety for years but May 22nd 2023 I finally made the decision to fully commit to it and stick through. I saw everything in my future and all the good stuff leading up for me falling apart if I kept drinking.

Now, I’m better than I’ve ever been. My mental health / depression has gotten a lot better. I’ve been able to just have clearer vision, understanding my values, where I want to focus my energy, etc.

Some things have definitely been tough. Once going sober I realized how much social anxiety I actually had that was being masked with alcohol the whole time. Been learning how to listen to my body, allow it to heal, learning socialization in a healthy matter, etc.

Feels great to finally be in control""",
        "alcohol recovery, depression, social anxiety",
        "sobriety; values; listening to body; healthy socialization",
        "lifestyle, emotional",
        "1 year",
        "clear, proud",
        "alcohol recovery, depression",
        "Full visible post.",
    ),
    extension_row(
        "reddit_original_105",
        "https://www.reddit.com/r/stopdrinking/comments/x9lk0v",
        "r/stopdrinking",
        "post",
        "1 year sober, wanted to share",
        """1 year sober from a bottle or more a day of whiskey or rum, just wanted to share. 26M almost 27, had a terrible time with sobriety after 2 years of agony, and this last year has been hellacious, but a good learning experience. Hopefully year 2 is better, but I'll be damned what I thought would be impossible somehow has happened.

If I can do it, anyone can do it. Laying in the hospital detoxing after 2 bottles and a 12 pack on my last night drinking a year ago surely was a wake up call.

I hope everyone keeps going and realizes you are loved and what you are doing now others may not understand, hell you may not understand, but it is the right thing to be done. Long term will show the results.

I still deal with PAWS episodes but they have very much improved, what appears to be a delayed wake cycle issue, and depression, but you know what, im not drinking, im not in any debt anymore had 10,000 of credit card debt, have a little savings, been working a little bit, have atleast a few long term goals, paid down medical bills tho a few remain, taken on new responsibilities having new bills and responsibilities, and things appear to be moving in the right direction.

Despite my remaining issues, im here and im sober even tho depression and anxiety/possible adhd and  confusion makes me borderline crazy at times.

Wishing everyone the best, IWNDWYT, 1 year complete! Cheers!""",
        "alcohol recovery, PAWS, depression, anxiety, debt",
        "sobriety; hospital wake-up call; long-term goals; debt repayment; responsibilities",
        "lifestyle, practical, emotional",
        "1 year",
        "raw, hopeful",
        "alcohol detox, depression, anxiety",
        "Full visible post.",
    ),
    extension_row(
        "reddit_original_106",
        "https://www.reddit.com/r/stopdrinking/comments/1r1q67r/1_year_sober_missing_my_old_life/",
        "r/stopdrinking",
        "post",
        "1 year sober, missing my old life",
        """This is the ramblings of a terribly anxious alcoholic and drug addict, though alcohol was my primary poison. I would love to hear from others about there experiences with hitting the year milestone. How you felt, how you kept going.

I don’t miss the years of alcohol dependence, sitting in my room alone and drunk 24hrs a day. Crying when I walked to the bottle shop, wanting to stop so bad but too mentally fucked to ask for help. The seizures and detoxes, each time feeling the fleeting hope of finally escaping alcohol’s grip, before falling straight back into my old pattern because I felt like I didn’t belong in the world.

I do miss the parties, festivals, social gatherings, taking drugs and dancing, whimsy. I miss when alcohol made me care free and pushed my anxiety away, allowing me to express myself without crippling fear of rejection. Even though those times were leading me straight into addiction, I was genuinely so happy, I felt like I belonged.

Don’t get me wrong, I know I can’t get those times back.""",
        "alcohol recovery, drug use, anxiety, social fear, belonging",
        "sobriety; honesty about grief; community milestone support",
        "community, emotional",
        "1 year",
        "ambivalent, honest",
        "alcohol recovery, drug use, seizures/detox",
        "Visible post text.",
    ),
    extension_row(
        "reddit_original_107",
        "https://www.reddit.com/r/stopdrinking/comments/wzk5fs",
        "r/stopdrinking",
        "post",
        "1 year sober",
        """I’ve been having problems with alcohol since I was a teenager. I didn’t drink all the time but when I did I’d binge.

Twice in the past I’d given up drinking for a year, my drinking had become a problem where I was drinking a bottle of brandy or whisky on a weekday evening and would still be drunk the next morning.

Eventually I would have a drink and would then say “I only drink on occasions”, then “I only drink on weekends” and soon enough I was binge drinking again.

I have pretty severe depression and when I drink I feel great, however when I drink I seem to burn through all my happiness and I get unbelievably low for weeks after.

The last time I drank I had one of the best nights of my life. Then the day after and for weeks following I was miserable, I wanted to die. I got pretty close to making it happen.

So I stopped, cold turkey.

I’ve had issues with my job from not drinking, friends have distanced themselves, I’ve been laughed at and I get a lot of people giving me shit “if you’re not an alcoholic, why can’t you drink?”""",
        "alcohol, binge drinking, depression, suicidal thoughts",
        "sobriety; recognizing post-drinking crash; cold turkey; social pressure tolerance",
        "lifestyle, emotional",
        "1 year",
        "raw, serious",
        "suicidal thoughts, alcohol recovery",
        "Visible post text.",
    ),
    extension_row(
        "reddit_original_108",
        "https://www.reddit.com/r/leaves/comments/15dogke",
        "r/leaves",
        "post",
        "1 Year Weed & Cigarette Free Withdrawal Timeline",
        """1 year withdrawal timeline:

I've been trying to update everyone as time goes on. On July 28th 2022 I quit weed, caff3ine & cigarettes simultaneously. In 2019 I lost my father and it devastated me. I didn't want to feel the hurt of losing him so I started smoking weed just to numb myself. It wasnt long before I was puffing on my vape from the time I got home at 5pm until I went to bed at 11pm. Sometimes I would even wake up in the middle of the night and smoke some.

I experienced every major withdrawal symptom known to man and was convinced I was dying or that I had given myself brain damage. I had never experienced anxiety or panic attacks or any of the other symtpoms I had prior to quitting.

I experienced nausea, vomiting, muscle spasms/twitches, depersonalization, severe brain fog, severe panic & anxiety, tinnitus, heart palpitations, body zaps, tingling in my limbs, night sweats, night terrors, extreme dizziness, memory troubles, depression & vivid dreams and I was an emotional mess. I had no previous mental illness or history of depression & anxiety. I ended up in the ER three times because I thought I was having a heart attack and ultimately ended up on a heart monitor which showed absolutely nothing wrong.

The first 3 months were hell on earth. After 3 months the symptoms started to taper slowly. At 5 months the symptoms lessened significantly and I started going back to the gym. As soon as I started back to the gym all of the remaining symptoms disappeared with the exception of tinnitus, some mild dizziness and a bit of mild anxiety.

As of Friday I am one year weed and cigarette free and my symptoms are basically gone. I still get PAWS waves that last anywhere from 1 week to a month but they are incredibly mild and only consist of some random mild tinnitus, the odd small bit of dizziness and anxiety which is very manageable. I know what they are now so it really doesn't bother me much and I remind myself how far I've come and how mild these symptoms are in comparison to the beginning.""",
        "cannabis withdrawal, nicotine, caffeine, grief, panic, depression",
        "quitting substances; gym; time; recognizing PAWS waves; remembering progress",
        "lifestyle, practical",
        "1 year",
        "timeline, intense, hopeful",
        "substance recovery, ER visits, panic symptoms",
        "Visible post text.",
    ),
    extension_row(
        "reddit_original_109",
        "https://www.reddit.com/r/leaves/comments/1b0qbx9",
        "r/leaves",
        "comment",
        "Has anyone in this sub actually quit weed successfully?",
        """My withdrawals are basically unnoticeable if I keep my self busy and active.
The big question to ask yourself is: why do I smoke?
Is it self-medicating anxiety/depression/PTSD, is it boredom, what is it?
Once you have the answer to that you can begin to truly do the work of figuring out why you smoke in the first place.
For me it was being 13 and having undiagnosed PTSD/anxiety mixed with childhood trauma and neglect. When I smoked all of the anxiety went away and it felt like I was getting the hug I never got from my parents.""",
        "cannabis use, PTSD, anxiety, depression, childhood trauma",
        "staying busy; asking why; identifying self-medication; trauma work",
        "practical, emotional",
        "not specified",
        "reflective, vulnerable",
        "substance recovery, trauma",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_110",
        "https://www.reddit.com/r/leaves/comments/1b0qbx9",
        "r/leaves",
        "comment",
        "Has anyone in this sub actually quit weed successfully?",
        """7 years in June.  The first month was rough.  The next few months were better but not good.  It took me maybe 6 months to feel like  myself again.  After one year I didn't want it anymore and have never looked back.""",
        "cannabis recovery, withdrawal",
        "time; abstinence; waiting for desire to fade",
        "lifestyle",
        "7 years",
        "brief, steady",
        "substance recovery",
        "Short complete visible comment.",
    ),
    extension_row(
        "reddit_original_111",
        "https://www.reddit.com/r/leaves/comments/1b0qbx9",
        "r/leaves",
        "comment",
        "Has anyone in this sub actually quit weed successfully?",
        """Hello! Yes, I quit weed for over a year now (a year and six months)! I actually didn’t view it as something that was hurting me and instead thought it was a great coping skill and the only thing getting me through the day. I blamed depression and anxiety.
Then, I saw this group on Reddit and decided to follow it. This sparked me to do my own research about weed and how it affects  mental health and my body. I was scared and anxious to quit but I was broke, my mental health and memory were worse than ever and I was losing friends.  I began to read about peoples experiences with quitting and it gave me hope and courage that I could also do it.
I did it and had an amazing support system of friends and my (then) boyfriend. After I quit, I slowly noticed my motivation, confidence, memory, diet and savings improved. I wasn’t anxious about where and when I’m going to smoke next. I had to set boundaries with friends that still smoked but I was able to do it with a clear mind. My skin cleared up and I smelled better.""",
        "cannabis use, depression, anxiety, memory, friendship loss",
        "Reddit group; research; support system; boundaries with smoking friends",
        "community, emotional, practical",
        "1.5 years",
        "reflective, hopeful",
        "substance recovery",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_112",
        "https://www.reddit.com/r/leaves/comments/h9plc4",
        "r/leaves",
        "post",
        "One. Year. Free.",
        """2 days ago, on June 13, 2020 marked my 1 year weed and other subtance-free.

I spent more than 5h browsing this forum with people having the same difficulty as I had when I stopped smoking weed. I stopped because of a panic attack when I was high.

The first 6 months were the most difficult of my entire life. The first 3 the most horrendous. But, everything was a downhill after that.

I experienced extreme anxiety, panics, fear, weight gain and thousands of unpleasantnesses. I went to psychotherapy because of my anxiety. I started working out because of my weight gain. I started getting out of my comfort zone because of my fear and I started to enjoy life rather than smoking weed every night.

I do not regret it. It made me better. I didn't know who I was, what my limits were and what I was capable to do before I stopped smoking. I was always picked by others as "the lazy one". Turns out I'm not a lazy guy. Turns out I can get the job done and I can wake up early. After I stopped, my life was back on track.""",
        "cannabis recovery, panic attack, anxiety, fear, identity",
        "psychotherapy; working out; comfort-zone exposure; abstinence",
        "professional, lifestyle",
        "1 year",
        "proud, reflective",
        "substance recovery, panic symptoms",
        "Full visible post.",
    ),
    extension_row(
        "reddit_original_113",
        "https://www.reddit.com/r/leaves/comments/1ron5g7/1_year_without_weed_my_mind_finally_came_back/",
        "r/leaves",
        "post",
        "1 Year Without Weed My Mind Finally Came Back",
        """My anxiety had gotten a lot better at that point, but my depression and boredom were still rough.

But slowly… something started changing.

My sleep became deeper.

The constant tightness in my chest disappeared.

That terrifying racing heart at night stopped showing up.

The fear of dying randomly that used to haunt my mind began fading.

It felt like my nervous system was finally breathing again.

Today it has been one year.

One full year without weed.

No panic attacks.

No constant anxiety cloud hanging over my head.

Sometimes I still get anxious like any human does, but it passes. It no longer controls my life.

My heart feels calm again.

And the biggest realization?

The life I thought weed was enhancing… it was actually stealing from me.

Quietly. Slowly. Without me noticing.

If you’re someone reading this who feels stuck in the same loop I was in, listen to me carefully.

Your brain is not broken.

It just needs time to heal.

And when it does, the peace that comes back is something you didn’t even realize you had lost.

One year later I’m just grateful.

Grateful for a quiet mind.""",
        "cannabis recovery, anxiety, panic attacks, depression, insomnia",
        "one year without weed; time; sleep recovery; nervous system healing",
        "lifestyle, emotional",
        "1 year",
        "soothing, hopeful",
        "substance recovery, panic symptoms",
        "Visible post text.",
    ),
    extension_row(
        "reddit_original_114",
        "https://www.reddit.com/r/Psychosis/comments/1mivmaq",
        "r/Psychosis",
        "post",
        "Fully Recovered",
        """You can check my past posts I had bad psychosis ~16 months ago had to miss a year of school went thru the psych ward, meds, therapy and all that.

I can now say I’ve fully recovered. There was a time I used to check this sub every day to make myself feel optimistic about recovery and relate to other users. I don’t want to brag I just remember desperately looking for success stories and seeing how many people don’t come back to this sub once they’ve made it to the other side. Tbh I don’t check this sub very often anymore but I swear I obsessively read every relevant post and comment for months.

Just wanted to share how I’m doing. You will get through this! I was going through the worst year and a half of my life. Your resilience will make you stronger!""",
        "psychosis recovery, school disruption, therapy, medication",
        "psych ward; medication; therapy; reading success stories; time",
        "professional, community",
        "16 months",
        "hopeful, concise",
        "psychosis, medication mention",
        "Full visible post.",
    ),
    extension_row(
        "reddit_original_115",
        "https://www.reddit.com/r/Psychosis/comments/1mivmaq",
        "r/Psychosis",
        "comment",
        "Fully Recovered",
        """Gym walks talking with friends I think time was the most important just need time to heal and put the episode behind""",
        "psychosis recovery, isolation, healing",
        "gym; walks; talking with friends; time",
        "lifestyle, emotional",
        "gradual",
        "brief, practical",
        "psychosis",
        "Short complete visible reply on what helped.",
    ),
    extension_row(
        "reddit_original_116",
        "https://www.reddit.com/r/Psychosis/comments/1mivmaq",
        "r/Psychosis",
        "comment",
        "Fully Recovered",
        """Not anymore yes I was sleeping 12+ hours a day w no motivation tired all the time""",
        "psychosis recovery, low motivation, hypersomnia",
        "time; recovery from low motivation",
        "lifestyle",
        "months",
        "brief, honest",
        "psychosis, low motivation",
        "Short complete visible reply.",
    ),
    extension_row(
        "reddit_original_117",
        "https://www.reddit.com/r/Psychosis/comments/1mivmaq",
        "r/Psychosis",
        "comment",
        "Fully Recovered",
        """I’m going back to school in two weeks I study mechanical engineering thank you!""",
        "psychosis recovery, school return, engineering",
        "returning to school; rebuilding identity",
        "practical",
        "16 months",
        "brief, hopeful",
        "psychosis",
        "Short complete visible reply.",
    ),
    extension_row(
        "reddit_original_118",
        "https://www.reddit.com/r/Psychosis/comments/1mivmaq",
        "r/Psychosis",
        "comment",
        "Fully Recovered",
        """Yeah definitely prob went away a few months after I got off the meds happened gradually""",
        "psychosis recovery, anhedonia, medication",
        "time after medication; gradual emotional return",
        "professional, emotional",
        "a few months",
        "brief, practical",
        "medication mention, anhedonia",
        "Short complete visible reply.",
    ),
    extension_row(
        "reddit_original_119",
        "https://www.reddit.com/r/Psychosis/comments/1mivmaq",
        "r/Psychosis",
        "comment",
        "Fully Recovered",
        """I think my psychosis was from weed and stress they had a bunch of diagnoses but no one was really sure. I got general anxiety, schizophreniform, bipolar 1, manic depressive, diagnoses but I still don’t know what I have. Hopefully it was just a one time thing. I was on risperidone for 9 months and Vraylar for 3 months then I’ve been off for 4 months""",
        "psychosis, cannabis, stress, diagnosis uncertainty",
        "medication; time off medication; avoiding recurrence",
        "professional",
        "16 months",
        "uncertain, candid",
        "psychosis, medication mention, cannabis mention",
        "Full visible reply; needs clinical review.",
    ),
    extension_row(
        "reddit_original_120",
        "https://www.reddit.com/r/Psychosis/comments/1cc727e/i_made_a_full_recovery_from_psychosis/",
        "r/Psychosis",
        "comment",
        "I made a full recovery from psychosis",
        """I always thought I had made a full recovery
I thought I did back when I had my first episode which lasted a month at 18 (2015) - i slowly stopped medication with the monitoring of my psychiatrist, I thought I did after it relapsed again for a month when I was 23 (2020) i stopped medication with the monitoring of my psychiatrist again, and I think I have fully recovered from my most recent episode which lasted the longest when I was 25 (2023)
Now I dont dare to stop the medication anymore, I can't afford to have another episode I will literally kill myself if I do - I just have to manage it like the diabetes thats all""",
        "psychosis recovery, relapse, medication management",
        "psychiatrist monitoring; staying on medication; relapse prevention",
        "professional",
        "2015-2023",
        "serious, candid",
        "suicidal statement, psychosis, medication mention",
        "Visible comment; high-risk content needs moderation.",
    ),
    extension_row(
        "reddit_original_121",
        "https://www.reddit.com/r/Psychosis/comments/1cc727e/i_made_a_full_recovery_from_psychosis/",
        "r/Psychosis",
        "comment",
        "I made a full recovery from psychosis",
        """I took medication for a long time so that helped. Staying on your medication consistently will help a lot. I also got a lot of sleep. Avoid drugs including caffeine.""",
        "psychosis recovery, medication, sleep, drug avoidance",
        "consistent medication; sleep; avoiding drugs and caffeine",
        "professional, lifestyle",
        "long time",
        "brief, practical",
        "medication mention, drug avoidance",
        "Short complete visible comment.",
    ),
    extension_row(
        "reddit_original_122",
        "https://www.reddit.com/r/Psychosis/comments/1cc727e/i_made_a_full_recovery_from_psychosis/",
        "r/Psychosis",
        "comment",
        "I made a full recovery from psychosis",
        """Yes and yes. I remember doing a IQ test while i was in the hospital ward and being in the 11 percentile. Then about 6 months later I was in the 50th percentile.
But yeah everything returned to normal pretty much.""",
        "psychosis recovery, cognition, hospital",
        "time; cognitive recovery",
        "professional",
        "about 6 months",
        "brief, reassuring",
        "psychosis, cognitive symptoms",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_123",
        "https://www.reddit.com/r/Psychosis/comments/1eoymnk/do_people_recover_from_psychosis/",
        "r/Psychosis",
        "comment",
        "Do people recover from psychosis?",
        """A full recovery took a couple of years but she has a normal life now, I’m still here, her family is a big part of her life again, she’s going back to school and her mental health is stronger than ever. Don’t give up hope ❤️""",
        "psychosis recovery, family, school return",
        "time; family support; school return; hope",
        "emotional, practical",
        "a couple of years",
        "brief, hopeful",
        "psychosis",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_124",
        "https://www.reddit.com/r/Psychosis/comments/1eoymnk/do_people_recover_from_psychosis/",
        "r/Psychosis",
        "comment",
        "Do people recover from psychosis?",
        """She needed to get on the right meds (took two months of trial and error during her hospitalization) then subsequently have support from friends and family. After getting medicated / stabilized she was on meds that flattened her personality quite heavily for about 6 months. When she stopped the meds it took another 6 months for her mental state to return to what I’d consider normal for her.""",
        "psychosis recovery, hospitalization, medication, family support",
        "right medication; hospitalization; friends and family support; time after medication",
        "professional, emotional",
        "about 14 months",
        "specific, caregiver perspective",
        "medication mention, hospitalization",
        "Full visible follow-up comment.",
    ),
    extension_row(
        "reddit_original_125",
        "https://www.reddit.com/r/Psychosis/comments/1eoymnk/do_people_recover_from_psychosis/",
        "r/Psychosis",
        "comment",
        "Do people recover from psychosis?",
        """I had my last episode in March this year and I'm what can be deemed as fully recovered. Recovery usually takes a couple months for me and I was always able to work again fulltime.""",
        "psychosis recovery, work",
        "time; returning to full-time work",
        "practical",
        "a couple months",
        "brief, hopeful",
        "psychosis",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_126",
        "https://www.reddit.com/r/Psychosis/comments/1eoymnk/do_people_recover_from_psychosis/",
        "r/Psychosis",
        "comment",
        "Do people recover from psychosis?",
        """Hi it's been 4 mo for me And I've been doing group therapy for 3 weeks. Something clicked this week and I started feeling myself again.""",
        "psychosis recovery, group therapy, identity",
        "group therapy; time; feeling like self again",
        "professional, emotional",
        "4 months",
        "brief, hopeful",
        "psychosis",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_127",
        "https://www.reddit.com/r/Psychosis/comments/1eoymnk/do_people_recover_from_psychosis/",
        "r/Psychosis",
        "comment",
        "Do people recover from psychosis?",
        """I lost my last job due to Psychosis but thanks to my support neteork, I found my footing again. I now have a new job with nice people and a better salary.""",
        "psychosis recovery, job loss, work return",
        "support network; new job; better environment",
        "practical, emotional",
        "not specified",
        "brief, encouraging",
        "psychosis, job loss",
        "Full visible comment with typo preserved.",
    ),
    extension_row(
        "reddit_original_128",
        "https://www.reddit.com/r/Psychosis/comments/1qm7pxl/is_full_recovery_possible/",
        "r/Psychosis",
        "comment",
        "Is full recovery possible?",
        """Full recovery is absolutely possible. It takes more time than we tend to expect but it can definitely eventually happen. It's very important that you avoid further episodes by being mindful of your triggers especially drug use.
If it was cannabis induced there's lots of information on recovery in the community hub at r/cannabis_psychosis.""",
        "psychosis recovery, cannabis-induced psychosis, triggers",
        "time; trigger awareness; avoiding drug use; community resources",
        "educational, practical",
        "longer than expected",
        "brief, reassuring",
        "psychosis, drug use mention",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_129",
        "https://www.reddit.com/r/Psychosis/comments/1iz8wme/do_people_who_get_psychosis_once_recover_fully/",
        "r/Psychosis",
        "comment",
        "Do people who get psychosis once recover fully?",
        """I’ve had psychosis twice. One of the lucky ones that fully recovered. Mine were purely substance induced, however. Had to stop all drugs and alcohol. But I’m off antipsychotics and have been for almost two years (when I got sober) with no recurrent issues.""",
        "psychosis recovery, substance-induced psychosis, sobriety",
        "stopping drugs and alcohol; sobriety; time off antipsychotics",
        "lifestyle, professional",
        "almost 2 years",
        "brief, hopeful",
        "psychosis, substance recovery, medication mention",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_130",
        "https://www.reddit.com/r/Psychosis/comments/1iz8wme/do_people_who_get_psychosis_once_recover_fully/",
        "r/Psychosis",
        "comment",
        "Do people who get psychosis once recover fully?",
        """Have had bipolar psychosis a bunch of times .. fully recovered and thriving. Yes for me it will come back if not taking meds or healthy lifestyle.
Edit: maybe not full recovery because I’ll take meds my entire life but very well managed.""",
        "bipolar psychosis, medication management, lifestyle",
        "medication; healthy lifestyle; long-term management",
        "professional, lifestyle",
        "ongoing",
        "brief, realistic",
        "psychosis, medication mention, bipolar",
        "Full visible comment.",
    ),
    extension_row(
        "reddit_original_131",
        "https://www.reddit.com/r/Psychosis/comments/15iz4ct",
        "r/Psychosis",
        "comment",
        "Have you guys fully recovered?",
        """I stressed for a few years it would happen again. But alas, I recovered fully, have my own place, a great career I built without a college degree, and an amazing boyfriend. It took a lot of trial and error, a lot of therapy (which I am still in), medication, and patience. I stay mindful that it can happen again and have developed skills to recognize if I feel myself getting too overwhelmed. I try not to stress about it and stay honest with myself.
Weed/stress/ptsd induced for reference. I was blazed out of my mind for months on end to deal with undiagnosed and untreated ptsd. I have smoked extremely minimally since, after I have a few beers. The babiest of baby hits I can handle. Probably not the best advice for this sub but being honest with yall.""",
        "psychosis recovery, PTSD, cannabis, stress, career, relationships",
        "therapy; medication; patience; overwhelm recognition; honesty",
        "professional, practical, emotional",
        "years",
        "candid, hopeful",
        "psychosis, cannabis, PTSD, medication mention",
        "Full visible comment; needs moderation for cannabis mention.",
    ),
    extension_row(
        "reddit_original_132",
        "https://www.reddit.com/r/Psychosis/comments/1e0kz0t",
        "r/Psychosis",
        "comment",
        "How long it took for you to recover? Did you fully recovered cognitively?",
        """It took me 2 years to make a full recovery. I went through multiple psychotic episodes and such. Thankfully I made a full recovery, am happy and have no thoughts of suicide.""",
        "psychosis recovery, suicidal thoughts, cognition",
        "time; full recovery after multiple episodes",
        "emotional",
        "2 years",
        "brief, hopeful",
        "psychosis, suicidal thoughts",
        "Full visible comment.",
    ),
]

ROWS.extend(FULL_TEXT_ADDITIONAL_ROWS)


README_ROWS = [
    ["Purpose", "Guiden MVP workbook built from Reddit stories that were directly accessible through normal search/open paths, without Reddit API access."],
    ["Content", "The original_story column contains visible Reddit post/comment text copied for a closed prototype dataset with source_url attribution."],
    ["Use in app", "Search over original_story, struggle_tags, what_helped, support_type, and tone."],
    ["Review", "Rows marked needs_human_review include medication, self-harm, diagnosis, or other sensitive details."],
    ["Replacement plan", "For production, replace these rows with opt-in Guiden submissions or licensed/permissioned stories."],
]


def autosize(sheet, widths: dict[str, int]) -> None:
    for idx, header in enumerate(HEADERS, start=1):
        sheet.column_dimensions[get_column_letter(idx)].width = widths.get(header, 18)


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Original Stories"
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
            "story_id": 20,
            "source_platform": 16,
            "source_url": 52,
            "subreddit": 24,
            "source_kind": 20,
            "post_title": 46,
            "original_story": 110,
            "struggle_tags": 36,
            "what_helped": 52,
            "support_type": 30,
            "timeframe": 24,
            "tone": 24,
            "risk_flags": 36,
            "retrieval_status": 28,
            "moderation_status": 24,
            "prototype_notes": 52,
        },
    )
    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 220

    table = Table(displayName="GuidenOriginalStories", ref=f"A1:{get_column_letter(len(HEADERS))}{ws.max_row}")
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
    tag_ws.column_dimensions["A"].width = 34
    tag_ws.column_dimensions["B"].width = 14
    tag_ws.sheet_view.showGridLines = False
    tag_table = Table(displayName="GuidenOriginalTagSummary", ref=f"A1:B{tag_ws.max_row}")
    tag_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True)
    tag_ws.add_table(tag_table)

    readme_ws = wb.create_sheet("README")
    readme_ws.append(["Field", "Guidance"])
    for row in README_ROWS:
        readme_ws.append(row)
    for cell in readme_ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    readme_ws.column_dimensions["A"].width = 20
    readme_ws.column_dimensions["B"].width = 120
    for row in readme_ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    readme_ws.sheet_view.showGridLines = False

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(OUTPUT_PATH.resolve())


if __name__ == "__main__":
    main()
