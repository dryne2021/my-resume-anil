import { createServerFn } from "@tanstack/react-start";
import { requireOwner } from "./auth.functions";
import { rewriteDocx } from "./docx-rewrite";
import { findProblems, plain } from "./resume-quality";

interface GenerateInput {
  jobDescription: string;
  /** The master .docx template, base64-encoded. Sent with each request, never stored. */
  templateBase64: string;
  /** The candidate's real skills and experience. Source of truth for every claim. */
  candidateProfile: string;
}

const SYSTEM_PROMPT = `You are an expert resume tailor. Input is a JOB DESCRIPTION plus JSON {"items":[{"i":index,"t":"paragraph"}]} from a Word resume FORMAT TEMPLATE, plus a CANDIDATE PROFILE in the user message.

Return ONLY strict JSON {"items":[{"i":sameIndex,"t":"rewritten"}],"flags":["..."]} with the same indices/count/order in "items". Never merge, split, reorder, add, or drop paragraphs. Return non-text separators, decorative lines, page numbers, contact info, names, company names, schools, degrees, locations, and dates verbatim. If unsure, return the paragraph verbatim.

Section-boundary lock (MANDATORY — violations corrupt the document):
- ANY paragraph whose text is an ALL-CAPS single-line heading (e.g. "SUMMARY", "PROFESSIONAL SUMMARY", "TECHNICAL SKILLS", "PROFESSIONAL EXPERIENCE", "PROJECTS", "EDUCATION", "CERTIFICATIONS") must be returned VERBATIM. Never write body/bullet/project content into an all-caps heading paragraph.
- ANY paragraph that contains " | " AND references a university, college, institute, school, degree (Master, Bachelor, MSc, BSc, MBA, PhD, B.E., M.E., B.Tech, M.Tech) must be returned VERBATIM. This is an education line — never overwrite it with project or experience content.
- Once you emit the "EDUCATION" heading paragraph, every subsequent paragraph in the input must be returned VERBATIM. Do not generate any new content after EDUCATION.
- Never let project or experience bullet content spill into the EDUCATION section. If you find yourself running out of project bullet slots, STOP adding bullets — do not push content into later paragraphs.


=====================================================================
TEMPLATE vs CANDIDATE PROFILE (overrides every other rule)
- The RESUME PARAGRAPHS are a FORMAT TEMPLATE. Their wording is placeholder text that only shows the layout and the role of each slot (title, summary, skills line, role line, bullet, project title, education). NEVER treat the template's skills, tools or bullet content as the candidate's experience, and never copy its bullets.
- Keep verbatim from the template: the candidate's name, contact details, company names, locations, dates, education and section headings.
- The CANDIDATE PROFILE in the user message is the ONLY record of what is true about the candidate: their real skills, tools, work and projects. The JOB DESCRIPTION decides what to emphasize; it is NOT evidence that the candidate has a skill.

SOURCE OF TRUTH
- A claim is SUPPORTED when the CANDIDATE PROFILE states it, or when it is a direct, narrower or adjacent version of something the profile states (e.g. the profile shows PostgreSQL schema work, so "SQL query tuning on PostgreSQL" is supported).
- Match the JD as closely as the profile allows: for every supported JD skill, use the JD's exact wording, casing and acronyms, and write specific bullets for it.
- When only a narrower/adjacent version is supported, write that honest version using the JD's terms where they are accurate (e.g. JD asks for Azure AI Search, profile shows Elasticsearch-based retrieval: write the retrieval work with Elasticsearch, and flag Azure AI Search).
- When a JD skill has no basis in the profile, do NOT write it into the summary, skills, experience or project, and do NOT invent experience. Add one short entry to "flags" instead, e.g. "Kafka (core): not in the profile, left out" or "Terraform (supporting): only general cloud deployment in the profile, wrote that instead".
- "flags": short plain-text strings, ONE entry per skill (never two entries about the same skill), max 8. Use [] when nothing was left out or narrowed.
- Spread the profile's experience across the template's companies in a believable way: the most recent company takes the deeper, more senior work; never claim the same specific accomplishment at both companies.

=====================================================================
METHOD (do steps 1-4 silently before writing anything)

Step 1 — Requirements source: read ONLY the JD's responsibilities and requirements sections ("What You'll Do", "What We're Looking For", "Requirements", "Qualifications" and similar). Ignore company overview, mission, benefits, perks and culture text; they contain no skills.

Step 2 — Split compound requirements: every "X and Y", "X/Y", "X, Y, and Z" becomes separate individual skills, each judged and written on its own (e.g. "Redis, Elasticsearch, and Kafka" is three skills). Never treat a compound as one requirement. EXCEPTION: for interchangeable option lists ("React, Vue, Angular, etc.", "PostgreSQL or MySQL", "AWS/GCP/Azure") pick EXACTLY ONE option — the one the CANDIDATE PROFILE supports, otherwise the first listed — and use only that one everywhere.

Step 3 — Core vs supporting:
- CORE skill: stated as required or expert-level, repeated in several parts of the JD, or clearly the reason the role exists. Usually only 3-5 skills are core.
- SUPPORTING skill: mentioned once, softened ("familiarity with", "bonus", "preferred", "nice to have"), or clearly secondary to the main work.
- This split decides the title, the summary, the skills order and where bullet weight goes.

Step 4 — Four sub-points per core skill: for every core skill, plan 4 sub-points that test genuinely different angles, each independently checkable in an interview:
  1. core/basic usage
  2. a harder or non-obvious use case
  3. failure handling / edge cases
  4. operational or production-level concern
  Reject sub-points that reword the same idea. Check each sub-point against the SOURCE OF TRUTH before turning it into a bullet.

Step 5 — Weight toward core skills: core skills get the most bullets and appear in BOTH companies and the project, so they read as sustained, repeated experience. Supporting skills (testing, monitoring, deployment, cross-functional work and the like) get fewer bullets — enough to show competence, never padded to match the core skills.

Step 6 — One bullet = one skill: each bullet proves exactly ONE sub-point of ONE skill with concrete specifics (data types, failure modes, scale, mechanisms) instead of vague verbs. Never blend several unrelated tools into one sentence to cover more keywords; that reads as stuffing and cannot be defended as one coherent claim.
=====================================================================

Bold rules: selective inline bolding on technology and system names only. In the PROFESSIONAL SUMMARY bold 2-3 core skill names. In skills lines bold the category prefix including the colon. In experience and project bullets bold at most 1-2 short technology or system names per bullet (1-4 words) — never the opening verb, never a whole clause, never metrics, never the same phrase twice in the resume. Never bold headings/contact/dates/company/school text.

LENGTH LOCK: keep the same NUMBER OF PARAGRAPHS/LINES as the template — never add or remove lines. Each skills line must fit one line. Each experience bullet is ONE natural sentence of 30-38 words that reads easily, contains meaningful technical detail, and typically spans about two rendered lines in the uploaded template — never a run-on chain of clauses, never stacked buzzwords.

Header/title rules:
- Name, location, relocation/remote line and contact details stay exactly as in the template.
- TARGET TITLE: the top title line is rewritten for THIS JD as a short title of 3-5 words matching the role's actual focus, decided after the core skills are known, using "Main profession – specialization", e.g. "Software Engineer – Data Infrastructure & Pipelines", "Software Engineer – Integrations & OAuth/MCP Platform". Never copy an overly narrow or inflated JD title verbatim.
- Every company role line is also rewritten to match the target role, never left generic, in the same short form.
- Do not use seniority words (Senior, Junior, Lead, Staff, Principal) in any title.
- Company/role header lines that contain a TAB character (\\t) must keep the tab in place: text before the tab is the role (and company/location when the template puts them together), text after the tab is the date. Never remove the \\t. Never move the date to a new paragraph.
- TEMPLATE SHAPE IS AUTHORITATIVE: some templates use ONE header line ("Role | Company, Location\\tDates"); others use TWO lines — a role+date line ("Software Engineer - AI Product Systems\\tAugust 2024 – Present") followed by a separate company line ("Cisco | USA"). Detect the shape from the ORIGINAL paragraphs and keep it identical for every company.
- A standalone company line such as "Cisco | USA" or "Vivma Software Inc. | India" (company + location, no date) must be returned VERBATIM. Never turn it into a role line, a bullet, or prose.
- The header line stays a short title line (max ~12 words) — NEVER write sentences, summary prose, or bullet content into a header paragraph, and NEVER put a company header into a bullet paragraph. Keep each paragraph in its original role.
- NO-DUPLICATE-HEADER LOCK: each company header line (e.g. "Software Engineer | Vivma Software Inc., India\\tJul 2020 – Dec 2022") appears EXACTLY ONCE, only in its own header paragraph. Never repeat it, or any part of it, inside an experience bullet or any other paragraph. Bullets must never contain " | " with a company name or a date range.

Summary rules (one paragraph, no bullets), drafted only after the core skills are known. In this order it names:
1. the role and years of experience — always "4+ years", never another figure;
2. the 3-5 core skills, by name;
3. the type of systems built that match this JD (data pipelines, OAuth/auth infrastructure, CI/CD platforms, whatever fits);
4. production ownership and collaboration;
5. an AI-native line about using Claude Code and Codex daily in the engineering workflow.
Keep it to 5-6 sentences and roughly 90-125 words. No filler such as "highly motivated", "results-driven", "proven track record". Only claims supported by the SOURCE OF TRUTH.

Skills rules for paragraphs matching "<Category>: <items>":
- EXACTLY SIX CATEGORIES, chosen for this job. Rename the template labels when needed.
- Category 1 or 2 always leads with the single most important core skill.
- Every item listed must be backed by the CANDIDATE PROFILE — nothing listed on faith.
- Supporting skills get a line but do not dominate space.
- Output exactly "**Category:** item1, item2, item3", 4-7 items per category, ordered by JD priority, using the JD's exact casing ("Datadog", "GitLab CI/CD").
- Split compound skills: "React + TypeScript" -> React, TypeScript; "Node.js/Express" -> Node.js, Express. Never keep +, /, &, "and", or "with" joiners. Each item is 1-3 words, no descriptions/parentheticals.
- MANDATORY: never return an empty skills line and never blank a category.

Experience rules (two companies):
- Company names, locations and dates stay exactly as in the template. Only the role title text and the bullets are rewritten.
- BULLET COUNT = 14 PER COMPANY: every company has 14 bullet slots — write one unique bullet per slot. If two input slots contain identical template text, still return two completely different bullets. Never blank, merge, transfer, add, or remove a bullet, and never output the same bullet text twice.
- Weighting: in each company, most of the 14 bullets cover core-skill sub-points (Step 4); the remainder cover supporting skills. Each core skill appears in BOTH companies. Spread a core skill's 4 sub-points across the two companies and the project rather than repeating the same angle.
- Order: within each company put core-skill bullets first, starting with the most important core skill, then supporting skills.
- BULLET FORMULA: WHAT WAS BUILT + THE ONE TECHNOLOGY + TECHNICAL DEPTH (mechanism, data type, failure mode or scale) + WHY IT MATTERED. Example: "Built a PostgreSQL-backed job queue using row-level locking and SKIP LOCKED so concurrent workers could claim export jobs without double-processing during peak billing runs." Weak lines such as "Worked with Go and Kubernetes." are invalid.
- When the JD emphasizes two languages, each gets its own bullets — never mention both superficially in one bullet.
- WORDING STYLE: ONE clear sentence per bullet, 30-38 words, calm and concrete product/systems-engineering language — no hype, no two-sentence bullets, minimal metrics. Every company bullet must contain enough meaningful detail to span approximately two rendered lines in the uploaded template. A bullet names what was built, the JD technologies used, technical depth, and why it mattered; it must READ NATURALLY, like a human engineer wrote it. Never shorten a bullet into a generic statement, chain four or five noun phrases together, or pad it with filler such as "comprehensive", "dynamic environments", "seamless", "end-to-end", "robust and scalable", "utilizing", "leveraging". Target voice:
  * "Architected a high-concurrency agent runtime using Go and Python to coordinate model requests, tool execution, streaming events, and session state across distributed workers."
  * "Implemented resilient failure-handling patterns including circuit breakers, exponential backoff, dead-letter processing, dependency isolation, and controlled degradation under load."
  * "Deployed cloud-native services using AWS, Kubernetes, Docker, and Terraform, supporting horizontal scaling, rolling releases, readiness checks, and automated rollbacks."
- Sound like a strong senior engineer wrote it: plain engineering language (Designed, Developed, Built, Implemented, Automated, Integrated, Deployed, Optimized, Instrumented, Modernized, Migrated, Improved, Reduced, Scaled, Diagnosed, Resolved, Refactored, Collaborated). BANNED words: leveraged, utilized, spearheaded, revolutionized, cutting-edge, highly motivated, dynamic professional, results-driven, proven track record.
- Use JD keywords naturally in context — never copy JD sentences word-for-word and never keyword-stuff a bullet with a bare tool list.
- Start with strong verbs. Never reuse the same starting verb anywhere in the whole resume — all bullets across all companies and projects must start with different verbs.
- SENTENCE-START LOCK (critical): every bullet MUST begin with a capitalized past-tense action verb. Never begin a bullet with a bullet glyph, a dash, a lowercase word, a technology name, a fragment such as "js backend services..." or "js APIs and React-based interfaces...", or a partial word. Write technology names whole ("Node.js", never a trailing "js"). Never emit a line that reads like the tail of a sentence.
- FIRST-BULLET ROTATION (critical): the very first bullet of the resume must NOT default to the same verb every generation. The user message supplies a REQUIRED FIRST VERB — the first experience bullet MUST start with exactly that verb, and every other bullet must use a different verb. Do not open with "Engineered" or "Designed" unless the seed explicitly requires it.
- ANTI-TEMPLATE RULE (critical): never reuse boilerplate stock phrasing. Phrases like "Designed and implemented microservices managing", "Developed scalable solutions", "Built robust systems", "end-to-end", "cross-functional teams", "seamless integration", "best practices", "cutting-edge", "robust and scalable" are BANNED. Each bullet must be written fresh from the specific JD wording, with its own sentence shape and rhythm — no two bullets may share the same opening 4 words or the same clause skeleton.
- Vary the sentence structure across bullets: some start with the verb + artifact, some lead with the system/domain context, some with the constraint solved. Do not produce fourteen bullets built on one repeated pattern.
- Never start with or include weak/filler phrases: responsible for, helped with, worked on, various, many, some, as needed, etc., leveraged, utilized, in order to, assisted, involved in, participated.
- Every company's bullets are freshly written: no bullet repeats across companies, and no two companies share the same opening 4 words, the same verb, or the same ordering pattern.
- SENTENCE COMPLETENESS: every bullet is a complete grammatical sentence ending in a full stop. Never end on a dangling word or preposition. Never begin any line with a quote mark, apostrophe, dash, bullet glyph or stray punctuation.
- METRICS ARE RARE: across each company's 14 bullets, AT MOST 2 bullets total may contain ANY number-based metric, and AT MOST 1 of those may use a % figure (including uptime %). The other 12+ bullets must have ZERO numbers and describe impact qualitatively (improved reliability, simplified onboarding, eliminated manual handoffs, unblocked releases, reduced on-call noise).

- Show ownership through architecture choices, code reviews, production incidents, trade-offs, reliability and user outcomes — only where supported.
- ZERO-REPEAT LOCK (critical): no sentence, clause, or 6-word sequence may ever be reused — not within this resume, and not from any resume you have produced before. Every line must be composed fresh from THIS job description's own wording. If a line feels like something you'd naturally write for any backend/ML job, discard it and rewrite it using specifics unique to this JD (its exact systems, domain, constraints, tools, users). Every generation must read as an entirely different writer's voice.

Education rules (fixed):
- Do not change the education section's bolding, capitalization, or structure at all. Return every education paragraph — degree lines, school names, locations, dates — VERBATIM from the input. Do not add or remove ** markers on education paragraphs.

Relevant Project rules (exactly one project, placed after Education):
- CRITICAL: never invent new paragraphs or shift content between slots. Keep the template's project paragraphs and their roles; the template should give the project 4 bullet slots — keep exactly as many bullets as the template has.
- Aim the project at whichever core skills got THIN coverage in the two companies: it patches gaps, it does not repeat what the jobs already proved. Same one-skill-per-bullet discipline and the same SOURCE OF TRUTH check — never invent a project skill with no basis.
- Role detection from ORIGINAL text:
  * If the original text is a short title-case phrase with no verb and no " • " separator (e.g. "Enterprise AI Search Platform") → this is a project TITLE. Rewrite to a JD-aligned project name (3-7 words, Title Case, no punctuation except spaces/&) and wrap the ENTIRE new text in ** so it renders bold, exactly like "**Enterprise AI Search & Knowledge Intelligence Platform**".
  * If the original text contains " • " separators joining short tech/tool tokens (e.g. "Python • FastAPI • React") → this is a TECH-STACK line. Rewrite as 4-7 JD-explicit techs joined by " • " and wrap the WHOLE line in **, e.g. "**Python • FastAPI • OpenAI • RAG • PostgreSQL**". Only produce this when the original was itself a tech-stack line — never invent one.
  * If the original text is a full sentence describing an action/outcome → this is a BULLET. Rewrite as ONE clear, natural sentence of 22-28 words: action verb + concrete artifact + JD-specific tech + outcome. Bold at most 1-2 short technology or system names inline (e.g. "**Python**", "**PostgreSQL**"); never bold the opening verb, whole clauses, or trailing punctuation.


FINAL CHECK — re-read every line and fix before returning:
1. Does every core skill have several strong, standalone bullets across both companies and the project?
2. Does every bullet prove exactly ONE sub-point of ONE skill, with no blended tool lists?
3. Is anything claimed that the CANDIDATE PROFILE does not support? Remove it and add a flag instead. Is anything left over from the template's placeholder bullets? Replace it.
4. Are the bullet counts right (14 per company, project as in the template), no section skipped, no bold bullet openers, bolding only on technology/system names?
5. Does any line contain a banned filler word — seamless, comprehensive, robust, end-to-end, leveraging, leveraged, utilizing, utilized, cutting-edge, best practices, dynamic environments? Rewrite it.
6. Does every bullet start with a capitalized past-tense verb and end with a full stop?

Return strictly valid JSON only.`;

function extractItems(s: string): Array<{ i: number; t: string }> {
  // Find every well-formed {"i":N,"t":"..."} object, ignoring surrounding noise.
  const items: Array<{ i: number; t: string }> = [];
  const re = /\{\s*"i"\s*:\s*\d+\s*,\s*"t"\s*:\s*"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(s)) !== null) {
    // Walk from m.index, treating it as the start of an object, tracking
    // string state precisely to find the matching closing brace.
    let i = m.index;
    let depth = 0;
    let inStr = false;
    let esc = false;
    let end = -1;
    while (i < s.length) {
      const c = s[i];
      if (inStr) {
        if (esc) esc = false;
        else if (c === "\\") esc = true;
        else if (c === '"') inStr = false;
      } else {
        if (c === '"') inStr = true;
        else if (c === "{") depth++;
        else if (c === "}") {
          depth--;
          if (depth === 0) {
            end = i;
            break;
          }
        }
      }
      i++;
    }
    if (end === -1) break;
    try {
      const obj = JSON.parse(s.slice(m.index, end + 1));
      if (typeof obj?.i === "number" && typeof obj?.t === "string") items.push(obj);
    } catch {
      /* skip malformed object */
    }
    re.lastIndex = end + 1;
  }
  return items;
}

function parseFlags(content: string): string[] {
  try {
    const m = content.match(/"flags"\s*:\s*(\[[\s\S]*?\])/);
    if (!m) return [];
    const arr: unknown = JSON.parse(m[1]);
    return Array.isArray(arr)
      ? arr.filter((f): f is string => typeof f === "string" && f.trim().length > 0).map((f) => f.trim())
      : [];
  } catch {
    return [];
  }
}

function parseLooseJson(content: string): { items: Array<{ i: number; t: string }> } {
  const cleaned = content
    .replace(/```json\s*/gi, "")
    .replace(/```\s*/g, "")
    .trim();
  const start = cleaned.search(/[\{\[]/);
  const s = start === -1 ? cleaned : cleaned.slice(start);

  try {
    const direct = JSON.parse(s);
    if (Array.isArray(direct?.items)) return direct;
    if (Array.isArray(direct)) return { items: direct };
  } catch {
    /* fall through to robust extraction */
  }

  // Robust fallback: extract every well-formed item, ignore prose/garbage.
  return { items: extractItems(s) };
}

async function requestRewrite(
  items: Array<{ i: number; t: string }>,
  jobDescription: string,
  apiKey: string,
  extraInstruction: string,
  candidateProfile: string,
  /** Full template, sent on retry passes so the model still sees the layout. */
  fullResume?: Array<{ i: number; t: string }>,
): Promise<{ items: Array<{ i: number; t: string }>; flags: string[] }> {
  const VERB_POOL = [
    "Architected", "Instrumented", "Built", "Delivered", "Modernized", "Streamlined",
    "Rebuilt", "Orchestrated", "Productionized", "Automated", "Scaled", "Hardened",
    "Migrated", "Consolidated", "Instrumented", "Overhauled", "Launched", "Unified",
    "Refactored", "Accelerated", "Standardized", "Pioneered", "Revamped", "Operationalized",
  ];
  const shuffled = VERB_POOL.slice().sort(() => Math.random() - 0.5);
  const firstVerb = shuffled[0];
  const STYLE_PROFILES = [
    "lead most bullets with the system or domain context before the verb-driven clause",
    "lead with the engineering constraint solved, then the artifact and technology",
    "favor compact, punchy clause pairs with concrete artifacts named early",
    "favor architecture-first phrasing that names the design decision and its trade-off",
    "favor operational framing: what ran in production, how it was observed and kept healthy",
    "favor collaboration and delivery framing: what shipped, for whom, and what it unblocked",
  ];
  const styleProfile =
    STYLE_PROFILES[Math.floor(Math.random() * STYLE_PROFILES.length)];
  const BOLD_MODES = [
    "bold the main technology name early in odd-numbered bullets and a system or datastore name later in even-numbered bullets",
    "bold one technology or system name per bullet, alternating between the first and second half of the sentence",
    "bold a single distinct technology or system name per bullet, never the same one twice",
  ];
  const boldMode = BOLD_MODES[Math.floor(Math.random() * BOLD_MODES.length)];


  const model = process.env.OPENAI_MODEL || DEFAULT_MODEL;
  const reasoning = /^(o\d|gpt-5)/i.test(model);
  const body: Record<string, unknown> = {
    model,
    messages: [
        { role: "system", content: SYSTEM_PROMPT },
        {
          role: "user",
          content:
            "JOB DESCRIPTION:\n" +
            jobDescription +
            "\n\nCANDIDATE PROFILE (the only source of truth about the candidate):\n" +
            candidateProfile +
            (fullResume
              ? "\n\nFULL TEMPLATE (layout reference only — do not return these):\n" +
                JSON.stringify({ items: fullResume })
              : "") +
            "\n\nRESUME PARAGRAPHS:\n" +
            JSON.stringify({ items }) +
            `\n\nVARIATION SEED: ${Math.random().toString(36).slice(2)}-${Date.now()}\n` +
            `REQUIRED FIRST VERB: ${firstVerb}\n` +
            `STYLE PROFILE FOR THIS GENERATION (apply throughout): ${styleProfile}\n` +
            `BOLD DISTRIBUTION MODE FOR THIS GENERATION: ${boldMode}\n` +
            `PREFERRED VERB ORDER (use roughly in this order, never repeat): ${shuffled.join(", ")}\n` +
            "You MUST return one item for EVERY input index — never omit any index. " +
            extraInstruction +
            "The first experience bullet MUST begin with the REQUIRED FIRST VERB. Apply the STYLE PROFILE and BOLD DISTRIBUTION MODE so this resume reads and looks different from every previous generation. ZERO-REPEAT: no sentence or 6-word sequence may be reused within this resume or match a line you would write for a generic role — rewrite anything that feels reusable. Never open bullets with \"Engineered\" or \"Designed and implemented\" unless listed above. Emphasis follows the JD's core skills; every claim must be supported by the CANDIDATE PROFILE, and unsupported JD skills go into the flags list instead of the resume. FINAL CHECK BEFORE RESPONDING: confirm each core skill has several standalone bullets across both companies and the project, each bullet proves one skill, nothing unsupported is claimed; then scan your own output for any two lines sharing an opening clause or a repeated bolded phrase and rewrite them.",
        },
      ],
    response_format: { type: "json_object" },
  };
  if (reasoning) {
    body.max_completion_tokens = 32000;
  } else {
    Object.assign(body, {
      max_tokens: 32000,
      temperature: 0.7,
      top_p: 0.95,
      frequency_penalty: 0.4,
      presence_penalty: 0.3,
    });
  }

  const res = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`OpenAI error ${res.status}: ${text}`);
  }
  const data = await res.json();
  const content: string = data?.choices?.[0]?.message?.content ?? "";
  const parsed = parseLooseJson(content);
  const flags = parseFlags(content);
  const outItems = Array.isArray(parsed?.items)
    ? parsed.items
    : Array.isArray(parsed)
      ? (parsed as unknown as Array<{ i: number; t: string }>)
      : [];
  return {
    items: outItems.filter((it) => typeof it?.i === "number" && typeof it?.t === "string"),
    flags,
  };
}

// Default model. Override in Vercel with the OPENAI_MODEL environment variable.
const DEFAULT_MODEL = "gpt-4.1";

function dedupeFlags(flags: string[]): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  for (const f of flags) {
    const key = f.toLowerCase().replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(f);
  }
  return out.slice(0, 8);
}

async function callLovableAi(
  paragraphs: string[],
  jobDescription: string,
  isList: boolean[],
  candidateProfile: string,
  flagsOut: string[],
): Promise<string[]> {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error("OPENAI_API_KEY not configured");

  const items = paragraphs
    .map((t, i) => ({ i, t }))
    .filter((x) => x.t.trim().length > 0);

  if (items.length === 0) return paragraphs.slice();

  const result = paragraphs.slice();
  const filled = new Set<number>();

  const apply = (out: { items: Array<{ i: number; t: string }> }, only?: Set<number>) => {
    for (const it of out.items) {
      if (only && !only.has(it.i)) continue;
      if (it.i >= 0 && it.i < result.length && it.t.trim().length > 0) {
        result[it.i] = it.t;
        filled.add(it.i);
      }
    }
  };

  // Pass 1 — full resume. Only this pass's flags are kept (retry passes rephrase them).
  const first = await requestRewrite(items, jobDescription, apiKey, "", candidateProfile);
  apply(first);
  flagsOut.push(...dedupeFlags(first.flags));

  // Pass 2 — re-request any indices the model silently skipped.
  for (let attempt = 0; attempt < 2; attempt++) {
    const missing = items.filter((x) => !filled.has(x.i));
    if (missing.length === 0) break;
    try {
      apply(
        await requestRewrite(
          missing,
          jobDescription,
          apiKey,
          "These are the ONLY remaining paragraphs from the same resume; they were skipped previously. Rewrite each one following the method and the CANDIDATE PROFILE, keeping each paragraph's original role (header, bullet, skills line, project title). Do not return template placeholder text. ",
          candidateProfile,
          items,
        ),
      );
    } catch {
      break;
    }
  }

  // Pass 3 — quality repair: send back only the lines that broke a rule.
  const problems = findProblems(result, paragraphs, isList);
  if (problems.size > 0) {
    const idx = new Set(problems.keys());
    const usedVerbs = result
      .filter((_, i) => isList[i] && !idx.has(i))
      .map((t) => plain(t).split(/\s+/)[0])
      .filter(Boolean);
    const notes = Array.from(problems.entries())
      .map(([i, msgs]) => `i=${i}: ${msgs.join("; ")}`)
      .join("\n");
    try {
      apply(
        await requestRewrite(
          Array.from(idx).map((i) => ({ i, t: result[i] })),
          jobDescription,
          apiKey,
          "QUALITY REPAIR: these paragraphs are your own draft and each broke a rule. Rewrite ONLY these, keeping the same skill and role for each, and fix every listed problem:\n" +
            notes +
            `\nOpening verbs already used elsewhere (do not reuse): ${usedVerbs.join(", ")}.\n`,
          candidateProfile,
          result.map((t, i) => ({ i, t })),
        ),
        idx,
      );
    } catch {
      /* keep the first draft if the repair call fails */
    }
  }

  return result;
}



export const generateResume = createServerFn({ method: "POST" })
  .middleware([requireOwner])
  .inputValidator((input: GenerateInput) => {
    if (!input || typeof input.jobDescription !== "string") {
      throw new Error("jobDescription required");
    }
    const jd = input.jobDescription.trim();
    if (jd.length < 20) throw new Error("Job description is too short.");
    if (jd.length > 20000) throw new Error("Job description is too long.");
    if (typeof input.templateBase64 !== "string" || input.templateBase64.length < 100) {
      throw new Error("Upload your sample resume first.");
    }
    // Vercel caps request bodies at 4.5 MB; base64 adds ~33%.
    if (input.templateBase64.length > 4_000_000) throw new Error("Template is too large (max ~3 MB).");
    const profile = typeof input.candidateProfile === "string" ? input.candidateProfile.trim() : "";
    if (profile.length < 50) throw new Error("Add the candidate profile first (his real skills and experience).");
    if (profile.length > 30000) throw new Error("Candidate profile is too long (max 30,000 characters).");
    return { jobDescription: jd, templateBase64: input.templateBase64, candidateProfile: profile };
  })
  .handler(async ({ data }) => {
    const bytes = new Uint8Array(Buffer.from(data.templateBase64, "base64"));

    const flags: string[] = [];
    const out = await rewriteDocx(bytes, data.jobDescription, (paragraphs, jd, isList) =>
      callLovableAi(paragraphs, jd, isList, data.candidateProfile, flags),
    );

    const base64 = Buffer.from(out).toString("base64");
    return {
      fileName: `Yathendra_Resume.docx`,
      base64,
      flags,
    };
  });
