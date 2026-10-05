// Quality checks for generated resume text. No dependencies so it can be tested on its own.

const BANNED_RE =
  /\b(seamless(?:ly)?|comprehensive|robust|end-to-end|leverag(?:e|ed|es|ing)|utiliz(?:e|ed|es|ing)|cutting-edge|best practices|dynamic environments?|spearhead(?:ed)?|revolutioniz(?:e|ed)|results-driven|proven track record|highly motivated)\b/gi;

const CONNECTORS = new Set(
  "a an the to for with by of in on and or that so from across into through during while which using via under at as when without before after its their".split(" "),
);

export const plain = (t: string) => t.replace(/\*\*/g, "").trim();

export const BULLET_MIN_WORDS = 25;
export const BULLET_MAX_WORDS = 42;
export const SUMMARY_MIN_WORDS = 85;
export const SUMMARY_MAX_WORDS = 130;

/** Index of the summary paragraph: the first paragraph after a "SUMMARY" heading. */
export function findSummaryIndex(originals: string[]): number {
  const h = originals.findIndex((t) => /^\s*(professional\s+)?summary\s*:?\s*$/i.test(plain(t)));
  if (h < 0) return -1;
  for (let i = h + 1; i < originals.length; i++) if (plain(originals[i])) return i;
  return -1;
}

const countSentences = (t: string) =>
  (t.replace(/\b(?:e\.g|i\.e|etc|vs|Inc|Ltd|Jr|Sr)\./gi, "x").match(/[.!?](?=\s|$)/g) ?? []).length;

/**
 * Server-side quality checks run on every generation. Returns problems keyed by
 * paragraph index so the bad lines can be sent back for one repair pass.
 */
export function findProblems(
  result: string[],
  originals: string[],
  isList: boolean[],
): Map<number, string[]> {
  const problems = new Map<number, string[]>();
  const add = (i: number, msg: string) => {
    const arr = problems.get(i) ?? [];
    arr.push(msg);
    problems.set(i, arr);
  };
  const seenVerbs = new Map<string, number>();
  const summaryIdx = findSummaryIndex(originals);

  result.forEach((raw, i) => {
    const t = plain(raw);
    if (!t || t === plain(originals[i] ?? "")) return;

    const banned = Array.from(new Set((t.match(BANNED_RE) ?? []).map((w) => w.toLowerCase())));
    if (banned.length) add(i, `remove the banned word(s): ${banned.join(", ")}`);

    if (i === summaryIdx) {
      const n = t.split(/\s+/).length;
      const sentences = countSentences(t);
      if (n < SUMMARY_MIN_WORDS || n > SUMMARY_MAX_WORDS)
        add(i, `summary is ${n} words; write 5-6 sentences totalling 90-125 words`);
      else if (sentences < 5 || sentences > 6)
        add(i, `summary has ${sentences} sentences; write exactly 5-6 sentences (90-125 words)`);
      return;
    }

    if (!isList[i]) return; // the remaining rules are for bullets only

    const words = t.split(/\s+/);
    const verb = (words[0] ?? "").replace(/[^A-Za-z-]/g, "").toLowerCase();
    if (verb) {
      if (seenVerbs.has(verb)) add(i, `starts with "${words[0]}", already used by another bullet; use a different past-tense verb`);
      else seenVerbs.set(verb, i);
    }
    if (!/^[A-Z]/.test(t)) add(i, "must start with a capitalized past-tense verb");
    if (!/[.]$/.test(t)) add(i, "must be one complete sentence ending with a full stop");

    // Run-on keyword strings ("OAuth 2.0 token validation JWT schemes mutual TLS…") have
    // almost no connector words; a normal sentence scores around 0.2-0.35.
    const tokens = t.match(/[A-Za-z0-9.+#/'-]+/g) ?? [];
    const connectors = tokens.filter((w) => CONNECTORS.has(w.toLowerCase())).length;
    if (tokens.length >= 15 && connectors / tokens.length < 0.1)
      add(i, "reads as a run-on list of keywords; rewrite as one natural sentence about one skill");

    const commas = (t.match(/,/g) ?? []).length;
    if (commas >= 4) add(i, "lists too many tools or items; make it prove ONE skill with one concrete mechanism");
    if (words.length > BULLET_MAX_WORDS) add(i, `too long (${words.length} words); keep it to 30-38 words`);
    if (words.length < BULLET_MIN_WORDS)
      add(i, `too short (${words.length} words); expand to 30-38 words by adding the concrete mechanism and why it mattered, so it spans about two lines`);
  });
  return problems;
}

// ---------------------------------------------------------------------------
// Structure helpers
// ---------------------------------------------------------------------------

export interface CoreSkill {
  name: string;
  match: string[];
}

/** Consecutive runs of bullet paragraphs, e.g. [company 1 bullets, company 2 bullets, project bullets]. */
export function bulletGroups(isList: boolean[]): number[][] {
  const groups: number[][] = [];
  let cur: number[] = [];
  isList.forEach((l, i) => {
    if (l) cur.push(i);
    else if (cur.length) {
      groups.push(cur);
      cur = [];
    }
  });
  if (cur.length) groups.push(cur);
  return groups;
}

const DATE_RANGE_RE =
  /((?:19|20)\d{2})\s*[–—-]\s*(?:(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*)?(Present|Current|Now|(?:19|20)\d{2})/i;

/** End year of the company that owns the bullet group starting at `start` (9999 = present). */
export function companyEndYear(paragraphs: string[], start: number): number | null {
  for (let i = start - 1; i >= 0 && i >= start - 4; i--) {
    const m = DATE_RANGE_RE.exec(plain(paragraphs[i] ?? ""));
    if (m) return /^\d/.test(m[2]) ? Number(m[2]) : 9999;
  }
  return null;
}

// Public release years, used to stop tools appearing under companies that ended earlier.
const TIMELINE: Array<[RegExp, number, string]> = [
  [/\bAzure AI Foundry\b/i, 2024, "Azure AI Foundry"],
  [/\bClaude Code\b/i, 2025, "Claude Code"],
  [/\bCodex\b/i, 2025, "Codex"],
  [/\bModel Context Protocol\b|\bMCP\b/, 2025, "MCP"],
  [/\bLangGraph\b/i, 2024, "LangGraph"],
  [/\bGemini\b/i, 2024, "Gemini"],
  [/\bAzure AI Search\b/i, 2023, "Azure AI Search"],
  [/\bAI Studio\b/i, 2023, "Azure AI Studio"],
  [/\bAzure OpenAI\b/i, 2023, "Azure OpenAI"],
  [/\bChatGPT\b/i, 2023, "ChatGPT"],
  [/\bGPT-4/i, 2023, "GPT-4"],
  [/\bLangChain\b/i, 2023, "LangChain"],
  [/\bLlamaIndex\b/i, 2023, "LlamaIndex"],
  [/\bBedrock\b/i, 2023, "Amazon Bedrock"],
  [/\b(?:AI|LLM)[- ]agents?\b|\bagentic\b/i, 2023, "AI agents"],
  [/(?<!GitHub )\bcopilots?\b/i, 2023, "AI copilots"],
  [/\bGitHub Copilot\b/i, 2023, "GitHub Copilot"],
  [/\bgenerative AI\b|\bGenAI\b/i, 2022, "generative AI"],
  [/\bLLMs?\b|\blarge language models?\b/i, 2022, "LLMs"],
];

const METRIC_RE =
  /(?:\d[\d,.]*\s?(?:%|x\b|×|ms\b|milliseconds|s\b|seconds?|minutes?|hours?|days?|weeks?|k\b|K\b|M\b|B\b|million|billion|TB\b|GB\b|requests|req\/s|rps\b|QPS\b|users|customers|services|microservices|teams|engineers|pipelines|models|endpoints|transactions|events|documents|records|agents|workflows|deployments|releases|tickets|incidents|clusters|nodes|applications|apps|APIs|\+))|\$\s?\d/i;

const hasMetric = (t: string) => METRIC_RE.test(plain(t));

const ANGLES =
  "(1) core usage, (2) a harder or non-obvious use, (3) failure handling or edge cases, (4) an operational or production concern";

/**
 * Structure checks that need the whole resume: core-skill coverage, number counts
 * and the timeline. Returns problems keyed by paragraph index.
 */
export interface StructureOptions {
  /** Template paragraphs (used to locate the Skills section). */
  originals?: string[];
  /** Every tool the JD names; each must be in Skills and in a bullet. */
  tools?: string[];
  /** Hiring company and product names that must never appear. */
  hiringNames?: string[];
}

const HEADING_WORD_RE = /^(experience|education|projects?|summary|skills|professional experience|work experience|certifications?)$/i;
const SKILLS_LINE_RE = /^[A-Za-z][A-Za-z0-9 &\/+\-]{1,40}:\s*\S/;
// "predating Azure AI Foundry", "before Azure AI Foundry's release" — needs a product name,
// so ordinary phrases like "before public releases" are not flagged.
const TIMING_TEXT_RE =
  /\bpredat(?:e|ed|es|ing)\b|\b(?:before|prior to|ahead of)\s+(?:[A-Z][\w.+-]*\s?){1,4}(?:'s)?\s*(?:release|launch|adoption|availability|introduction)\b/;

/** Indices of "Category: items" lines inside the template's SKILLS section. */
export function skillsLineIndices(originals: string[], isList: boolean[]): number[] {
  const start = originals.findIndex((t) => /^\s*(technical\s+|core\s+)?skills\s*:?\s*$/i.test(plain(t)));
  if (start < 0) return [];
  const out: number[] = [];
  for (let i = start + 1; i < originals.length; i++) {
    const t = plain(originals[i]);
    if (!t) continue;
    if (/^[A-Z][A-Z &\/-]{2,40}$/.test(t)) break; // next ALL-CAPS heading
    if (!isList[i] && SKILLS_LINE_RE.test(t)) out.push(i);
  }
  return out;
}

const lc = (t: string) => plain(t).toLowerCase();

export function findStructureProblems(
  result: string[],
  isList: boolean[],
  core: CoreSkill[],
  opts: StructureOptions = {},
): Map<number, string[]> {
  const problems = new Map<number, string[]>();
  const add = (i: number, msg: string) => {
    const arr = problems.get(i) ?? [];
    arr.push(msg);
    problems.set(i, arr);
  };
  const groups = bulletGroups(isList);
  const companies = groups.filter((g) => g.length >= 6);
  const top = core.slice(0, 3).filter((c) => c.match.length > 0);
  const mentions = (i: number, c: CoreSkill) =>
    c.match.some((m) => m && plain(result[i]).toLowerCase().includes(m.toLowerCase()));

  // Timeline — every bullet group, using the dates of the company above it.
  for (const g of groups) {
    const end = companyEndYear(result, g[0]);
    if (end === null || end === 9999) continue;
    for (const i of g) {
      for (const [re, year, label] of TIMELINE) {
        if (year > end && re.test(plain(result[i]))) {
          add(
            i,
            `mentions ${label} (released ${year}) but this company ended in ${end}; rewrite the same skill with the technology that existed then`,
          );
          break;
        }
      }
    }
  }

  companies.forEach((g) => {
    const end = companyEndYear(result, g[0]) ?? 9999;
    const released = (name: string) => TIMELINE.find(([re]) => re.test(name))?.[1] ?? 0;
    // Core-skill coverage: 4 bullets per top-3 core skill.
    const used = new Set<number>();
    for (const c of top) {
      const hits = g.filter((i) => mentions(i, c));
      hits.forEach((i) => used.add(i));
      let missing = 4 - hits.length;
      if (missing <= 0) continue;
      const spare = g.filter((i) => !used.has(i) && !top.some((t) => mentions(i, t))).reverse();
      for (const i of spare) {
        if (missing <= 0) break;
        // Keep at least two supporting-skill bullets per company.
        const supporting = g.filter((j) => !used.has(j) && !top.some((t) => mentions(j, t)));
        if (supporting.length <= 2) break;
        used.add(i);
        missing--;
        const older = released(c.name) > end;
        const label = older
          ? `${c.name}, written with its period-correct equivalent because this company ended in ${end} (e.g. ${c.match.filter((m) => m.toLowerCase() !== c.name.toLowerCase()).join(" / ") || "the technology that existed then"})`
          : c.name;
        add(
          i,
          `rewrite this bullet to focus solely on ${label} (it needs 4 bullets in this company, one per angle: ${ANGLES}); pick an angle not already used by the other bullets on this skill and name the technology explicitly`,
        );
      }
    }

    // Numbers: 5-7 bullets per company with one concrete figure.
    const withNum = g.filter((i) => hasMetric(result[i]));
    if (withNum.length < 5) {
      let need = 5 - withNum.length;
      for (const i of g) {
        if (need <= 0) break;
        if (hasMetric(result[i]) || problems.has(i)) continue;
        add(
          i,
          "end this bullet with ONE concrete, modest, believable number showing scale or impact (latency, volume, time saved, error rate, users), different in type from the other numbers",
        );
        need--;
      }
    } else if (withNum.length > 7) {
      withNum.slice(7).forEach((i) => add(i, "remove the number from this bullet; describe the impact in words"));
    }
  });

  const bullets = groups.flat();

  // Bullet hygiene: no semicolon second clauses, no release-timing commentary.
  for (const i of bullets) {
    const t = plain(result[i]);
    if (/;/.test(t)) add(i, "remove the semicolon; write ONE sentence");
    if (TIMING_TEXT_RE.test(t)) add(i, "do not write about release dates or timing; just describe the work with the period-correct technology");
  }

  // Hiring company / product names and trademarks must never appear.
  const banned = (opts.hiringNames ?? []).filter((n) => n.length > 2).map((n) => n.toLowerCase());
  result.forEach((raw, i) => {
    const t = lc(raw);
    if (!t || (opts.originals && raw === opts.originals[i])) return;
    const hit = banned.find((n) =>
      new RegExp(`(^|[^a-z0-9])${n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}([^a-z0-9]|$)`).test(t),
    );
    if (hit || /[™®]/.test(raw)) {
      add(
        i,
        `remove ${hit ? `"${hit}"` : "the trademarked product name"}: never mention the hiring company or its products; describe the work in this employer's own business instead`,
      );
    }
  });

  // Skills section: every line filled with real items; every JD tool present.
  const skillIdx = opts.originals ? skillsLineIndices(opts.originals, isList) : [];
  for (const i of skillIdx) {
    const t = plain(result[i]);
    const items = t.includes(":") ? t.slice(t.indexOf(":") + 1).trim() : "";
    const parts = items.split(",").map((x) => x.trim()).filter(Boolean);
    if (parts.length < 2 || parts.some((x) => HEADING_WORD_RE.test(x)) || /^[A-Z\s]+$/.test(items)) {
      add(i, "this skills line is empty or contains a heading/placeholder; fill it with 3-6 items this JD needs, renaming the category if it does not fit the JD (e.g. 'Programming Languages')");
    }
  }
  const tools = (opts.tools ?? []).filter((x) => x.length > 1);
  if (tools.length && skillIdx.length) {
    const skillsText = skillIdx.map((i) => lc(result[i])).join(" | ");
    const missing = tools.filter((x) => !skillsText.includes(x.toLowerCase()));
    if (missing.length) {
      for (const i of skillIdx) {
        add(i, `the Skills section is missing these JD tools: ${missing.join(", ")}; add each one to the most fitting skills line (rename a category that has no JD-relevant items), 3-6 items per line`);
      }
    }
  }
  if (tools.length && bullets.length) {
    const projectGroups = groups.filter((g) => g.length < 6);
    const coreMention = (i: number) => top.some((c) => mentions(i, c));
    const candidates = [
      ...projectGroups.flat(),
      ...companies.flatMap((g) => g.filter((i) => !coreMention(i)).reverse()),
    ];
    const taken = new Set<number>();
    for (const tool of tools) {
      const key = tool.toLowerCase();
      if (bullets.some((i) => lc(result[i]).includes(key))) continue;
      const year = TIMELINE.find(([re]) => re.test(tool))?.[1] ?? 0;
      const pick = candidates.find((i) => {
        if (taken.has(i) || problems.has(i)) return false;
        const g = groups.find((gg) => gg.includes(i));
        const end = g ? (companyEndYear(result, g[0]) ?? 9999) : 9999;
        return year <= end;
      });
      if (pick === undefined) continue;
      taken.add(pick);
      add(pick, `rewrite this bullet so it shows concrete, hands-on use of ${tool} as its subject (it is a JD tool with no bullet yet); keep it to that one skill`);
    }
  }
  return problems;
}

// ---------------------------------------------------------------------------
// Last-resort fixes applied in code after the AI repair rounds.
// ---------------------------------------------------------------------------

const REPLACEMENTS: Array<[RegExp, string]> = [
  [/\ban end-to-end\b/gi, "a full"],
  [/\bend-to-end\b/gi, "full"],
  [/\bseamlessly\b/gi, "smoothly"],
  [/\bseamless\b/gi, "smooth"],
  [/\bcomprehensive\b/gi, "full"],
  [/\brobust\b/gi, "reliable"],
  [/\bleveraging\b/gi, "using"],
  [/\bleveraged\b/gi, "used"],
  [/\bleverages\b/gi, "uses"],
  [/\bleverage\b/gi, "use"],
  [/\butilizing\b/gi, "using"],
  [/\butilized\b/gi, "used"],
  [/\butilizes\b/gi, "uses"],
  [/\butilize\b/gi, "use"],
  [/\bcutting-edge\b/gi, "modern"],
  [/\bbest practices\b/gi, "engineering standards"],
  [/\bdynamic environments\b/gi, "changing environments"],
  [/\bdynamic environment\b/gi, "changing environment"],
  [/\bspearheaded\b/gi, "led"],
  [/\bspearhead\b/gi, "lead"],
  [/\brevolutionized\b/gi, "reworked"],
  [/\brevolutionize\b/gi, "rework"],
  [/\bresults-driven\s*/gi, ""],
  [/\bproven track record\b/gi, "track record"],
  [/\bhighly motivated\s*/gi, ""],
];

const keepCase = (orig: string, rep: string) =>
  rep && orig[0] === orig[0].toUpperCase() ? rep[0].toUpperCase() + rep.slice(1) : rep;

export function replaceBannedWords(t: string): string {
  let out = t;
  for (const [re, rep] of REPLACEMENTS) out = out.replace(re, (m) => keepCase(m, rep));
  return out.replace(/ {2,}/g, " ");
}

const VERB_POOL = [
  "Built", "Designed", "Developed", "Implemented", "Delivered", "Engineered", "Automated",
  "Integrated", "Deployed", "Optimized", "Instrumented", "Modernized", "Migrated", "Refactored",
  "Scaled", "Hardened", "Streamlined", "Launched", "Introduced", "Established", "Created",
  "Configured", "Extended", "Shipped", "Standardized", "Consolidated", "Rebuilt", "Tuned",
  "Secured", "Reduced", "Improved", "Accelerated", "Simplified", "Containerized", "Orchestrated",
  "Authored", "Prototyped", "Productionized", "Diagnosed", "Resolved", "Validated", "Trained",
  "Evaluated", "Expanded", "Upgraded", "Restructured", "Assembled", "Drove", "Owned", "Partnered",
];

/** Replaces banned words everywhere and swaps any repeated opening verb for an unused one. */
export function autoFix(result: string[], originals: string[], isList: boolean[]): string[] {
  const out = result.map((t, i) => (plain(t) && t !== originals[i] ? replaceBannedWords(t) : t));
  const seen = new Set<string>();
  const firstWord = (t: string) => (t.match(/^\s*(\*\*)?([A-Za-z-]+)/)?.[2] ?? "");
  out.forEach((t, i) => {
    if (isList[i]) seen.add(firstWord(t).toLowerCase());
  });
  const counted = new Set<string>();
  out.forEach((t, i) => {
    if (!isList[i]) return;
    const w = firstWord(t);
    const key = w.toLowerCase();
    if (!key) return;
    if (!counted.has(key)) {
      counted.add(key);
      return;
    }
    const replacement = VERB_POOL.find((v) => !seen.has(v.toLowerCase()));
    if (!replacement) return;
    seen.add(replacement.toLowerCase());
    counted.add(replacement.toLowerCase());
    out[i] = t.replace(w, replacement);
  });
  return out;
}
