// Step 0 — a short, focused AI call that reads ONLY the job description and
// decides the core skills, supporting skills and tool list. The writing step
// and the automatic checks both use this, so the main skills are chosen once,
// carefully, instead of inside the long writing prompt.

import { chatJson } from "./openai";
import type { CoreSkill } from "./resume-quality";

export interface JdAnalysis {
  titleFocus: string;
  core: CoreSkill[];
  supporting: string[];
  tools: string[];
  hiringNames: string[];
  /** One title per candidate employer, most recent first. */
  roleTitles: string[];
}

const ANALYSIS_PROMPT = `You analyze a job description (JD) so a resume can be tailored to it. Read ONLY the responsibilities, requirements, qualifications, "key technologies" and skills sections; ignore company overview, mission, benefits, perks, working conditions and EEO text.

Return ONLY strict JSON:
{
  "titleFocus": "3-5 word specialization for the resume title, e.g. 'Embedded Linux & Test Automation'",
  "core": [{"name": "C++", "match": ["C++"], "olderEquivalent": ""}],
  "supporting": ["..."],
  "tools": ["..."],
  "hiringNames": ["..."],
  "roleTitles": ["...", "..."]
}

Rules:
- "core": EXACTLY 3 items, most important first. A core skill is a hands-on technical skill the JD states as required, repeats in several places, or is clearly the reason the role exists (e.g. a required programming language, the main platform, the main engineering activity such as "Software Testing & Debugging" or "Retrieval-Augmented Generation").
  * Programming languages the JD explicitly requires are strong core candidates; when the JD requires two languages (e.g. "Python, C++"), the less obvious one must not be dropped — include both if they are among the top needs.
  * NEVER core: anything the JD does not name; soft skills; documentation; version control / Git; developer tooling such as GitHub Copilot, Claude Code, Codex, Jira, IDEs; items softened with "familiarity", "working knowledge", "exposure", "nice to have", "bonus", "preferred". Those go in "supporting" (if technical) or nowhere.
  * "match": 1-4 exact words/phrases a resume bullet about this skill would contain (case-insensitive substrings), e.g. ["C++"], ["retrieval-augmented", "RAG"], ["test", "debug"].
  * "olderEquivalent": for a tool released 2022 or later (e.g. Azure AI Foundry, LangChain, GitHub Copilot), the closest technology that existed in 2020-2021 (e.g. "Azure Machine Learning"); otherwise "".
- "supporting": the other technical skills/responsibilities from the JD, split into single skills (max 12).
- "tools": EVERY named language, tool, platform, OS, framework, protocol and product from the JD, split into single items with official spelling ("Git Hub Copilot" -> "GitHub Copilot", "Azure Dev Ops" -> "Azure DevOps", "Foundry" -> "Azure AI Foundry"). For "X, Y, etc." keep each named item; for "X or Y or similar" keep only X.
- "hiringNames": the hiring company's name and any of its own product, brand or trademark names (anything with ™ or ®, or clearly proprietary). The resume must never mention these.
- "roleTitles": one job title for EACH candidate employer listed in the user message, in the same order (most recent first). Rules:
  * It is a real job title for the JD's job FUNCTION (e.g. JD for a healthcare business analyst -> "Business Systems Analyst"; JD for an AI engineer -> "AI Engineer"; JD for IT audit -> "IT Audit Analyst"), 2-6 words, optionally "Function – specialization" (e.g. "Software Engineer – Embedded Systems").
  * It must be believable at THAT employer's industry: never put the hiring company's industry words into it (no "Clinical", "Healthcare", "Battery", "Energy" in a title at a bank, a networking company or a food-delivery platform). Use the function, or a specialization that fits both the JD and the employer (e.g. "Business Systems Analyst – Data & Reporting").
  * The titles must all be DIFFERENT from each other. The earlier employer gets the same function at an earlier career stage or a closely related role (e.g. "Business Systems Analyst" then "Business Analyst"; "Software Engineer – Platform" then "Software Engineer"; "IT Audit Analyst" then "IT Compliance Analyst").
  * No "Senior", "Junior", "Lead", "Staff", "Principal", "Head" or "Manager" unless the JD itself is a management role.
  * Never reuse the template's current titles (given in the user message) unless they genuinely fit this JD.`;

function cleanList(v: unknown, max: number): string[] {
  if (!Array.isArray(v)) return [];
  return Array.from(
    new Set(v.filter((x): x is string => typeof x === "string").map((x) => x.trim()).filter((x) => x.length > 1)),
  ).slice(0, max);
}

export interface EmployerInfo {
  company: string;
  years: string;
  currentTitle: string;
}

export async function analyzeJob(
  jobDescription: string,
  apiKey: string,
  employers: EmployerInfo[] = [],
): Promise<JdAnalysis | null> {
  try {
    const content = await chatJson({
      apiKey,
      system: ANALYSIS_PROMPT,
      user:
        "JOB DESCRIPTION:\n" +
        jobDescription +
        (employers.length
          ? "\n\nCANDIDATE EMPLOYERS (most recent first) — give one roleTitle each:\n" +
            employers
              .map((e, k) => `${k + 1}. ${e.company || "(company)"} (${e.years}) — template title now: "${e.currentTitle}"`)
              .join("\n")
          : ""),
      maxTokens: 2000,
      temperature: 0.1,
    });
    const raw = JSON.parse(content.replace(/```json\s*|```/g, "").trim()) as Record<string, unknown>;
    const core: CoreSkill[] = (Array.isArray(raw.core) ? raw.core : [])
      .map((c) => {
        const o = c as { name?: unknown; match?: unknown; olderEquivalent?: unknown };
        const name = typeof o?.name === "string" ? o.name.trim() : "";
        const match = cleanList(o?.match, 4);
        const older = typeof o?.olderEquivalent === "string" ? o.olderEquivalent.trim() : "";
        const all = [...(match.length ? match : name ? [name] : []), ...(older ? [older] : [])];
        return { name, match: all };
      })
      .filter((c) => c.name)
      .slice(0, 3);
    if (core.length === 0) return null;
    return {
      titleFocus: typeof raw.titleFocus === "string" ? raw.titleFocus.trim() : "",
      core,
      supporting: cleanList(raw.supporting, 12),
      tools: cleanList(raw.tools, 30),
      hiringNames: cleanList(raw.hiringNames, 10),
      roleTitles: distinctTitles(cleanList(raw.roleTitles, 6), employers),
    };
  } catch {
    return null; // fall back to the writer choosing core skills itself
  }
}

/** Keeps a title per employer only when it is new and different from the others. */
function distinctTitles(titles: string[], employers: EmployerInfo[]): string[] {
  const seen = new Set<string>();
  return employers.map((e, k) => {
    const t = (titles[k] ?? "").trim();
    const key = t.toLowerCase();
    if (!t || t.split(/\s+/).length > 8 || seen.has(key)) return "";
    seen.add(key);
    return t;
  });
}

/** Text block given to the writing step. */
export function analysisBlock(a: JdAnalysis | null): string {
  if (!a) return "";
  return (
    "\n\nJD ANALYSIS (decided already — follow it exactly):\n" +
    `- Title specialization: ${a.titleFocus || "(choose)"}\n` +
    `- CORE SKILLS in priority order (the LAYOUT PLAN says how many bullets each gets per company; for older companies use the period-correct equivalent shown): ${a.core
      .map((c, k) => `#${k + 1} ${c.name} [match: ${c.match.join(" / ")}]`)
      .join("; ")}\n` +
    `- Supporting skills: ${a.supporting.join(", ") || "(none)"}\n` +
    `- TOOLS that must each appear in Skills AND in at least one bullet: ${a.tools.join(", ") || "(none)"}\n` +
    `- NEVER mention these hiring-company names or products anywhere: ${a.hiringNames.join(", ") || "(none)"}\n` +
    (a.roleTitles.some(Boolean)
      ? `- ROLE TITLES (already set in code — write each company's bullets as work done in that role): ${a.roleTitles
          .map((t, k) => `company ${k + 1}: ${t || "(keep)"}`)
          .join("; ")}\n`
      : "")
  );
}
