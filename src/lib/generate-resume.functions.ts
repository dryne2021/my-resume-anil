import { createServerFn } from "@tanstack/react-start";
import { requireOwner } from "./auth.functions";
import { rewriteDocx } from "./docx-rewrite";

interface GenerateInput {
  jobDescription: string;
  /** The master .docx template, base64-encoded. Sent with each request, never stored. */
  templateBase64: string;
}

const SYSTEM_PROMPT = `You are an expert resume tailor. Input is a JOB DESCRIPTION plus JSON {"items":[{"i":index,"t":"paragraph"}]} from a Word resume.

Return ONLY strict JSON {"items":[{"i":sameIndex,"t":"rewritten"}]} with the same indices/count/order. Never merge, split, reorder, add, or drop paragraphs. Return non-text separators, decorative lines, page numbers, contact info, names, company names, schools, degrees, locations, and dates verbatim. If unsure, return the paragraph verbatim.

Section-boundary lock (MANDATORY — violations corrupt the document):
- ANY paragraph whose text is an ALL-CAPS single-line heading (e.g. "SUMMARY", "PROFESSIONAL SUMMARY", "TECHNICAL SKILLS", "PROFESSIONAL EXPERIENCE", "PROJECTS", "EDUCATION", "CERTIFICATIONS") must be returned VERBATIM. Never write body/bullet/project content into an all-caps heading paragraph.
- ANY paragraph that contains " | " AND references a university, college, institute, school, degree (Master, Bachelor, MSc, BSc, MBA, PhD, B.E., M.E., B.Tech, M.Tech) must be returned VERBATIM. This is an education line — never overwrite it with project or experience content.
- Once you emit the "EDUCATION" heading paragraph, every subsequent paragraph in the input must be returned VERBATIM. Do not generate any new content after EDUCATION.
- Never let project or experience bullet content spill into the EDUCATION section. If you find yourself running out of project bullet slots, STOP adding bullets — do not push content into later paragraphs.


The JOB DESCRIPTION is the only source for rewritten content. Rewrite summary, skills, job title text, and experience bullets from scratch using JD terminology. Preserve the candidate's identity, contact details, employers, education, and all dates. Do not add visa, location, citizenship, sponsorship, or clearance.

ATS/JD lock (100% MATCH — highest priority after structure locks):
- Before writing anything, silently build a COVERAGE LIST: every concrete noun-phrase in the JD — languages, frameworks, libraries, tools, platforms, databases, cloud services, protocols, methodologies, domain terms, and each stated responsibility/requirement bullet. Include niche terms exactly as written (e.g. Rust, Go, WASM, MCP, Claude Code, Codex, OpenAI SDK, Agent SDK, LLM agents, agent orchestration, developer tooling, terminal/IDE tooling, CI/CD, shell scripting, Linux automation, Workday Extend, Workday Studio, EIB, Core Connectors, HCM, Payroll, SOAP, Calculated Fields, DevEx, desktop apps, agent workflows).
- MANDATORY: 100% of that COVERAGE LIST must appear VERBATIM somewhere in the final resume (summary, skills, experience bullets, or projects). Mirror the JD's exact casing, spelling, punctuation and acronym form (write both the acronym and expansion when the JD does, e.g. "Continuous Integration (CI)"). EXCEPTION: for interchangeable option lists (e.g. "React, Vue, Angular, etc."), include ONLY the first-listed option and drop the alternatives from the coverage list entirely.
- Distribute coverage: put every JD tool/tech in Skills, and additionally weave the most important 60% of them into experience bullets and projects so each keyword has context, not just a list entry.
- Every JD responsibility/requirement must have at least one bullet that demonstrably performs it, using the JD's own verbs and nouns.
- Before returning, silently re-check the COVERAGE LIST item by item; if any item is missing, rewrite a bullet or skills line to include it. Do not return output with an uncovered JD keyword.
- Do not use any skill/tech from the original resume unless the JD explicitly names it. Never invent skills the JD does not mention.
- C#, .NET, ASP.NET, .NET Core, and related Microsoft-stack terms are REQUIRED when the JD explicitly mentions them and FORBIDDEN otherwise.

Bold rules (MANDATORY, no exceptions): use **markers** for (a) 2-3 top JD keywords inside the PROFESSIONAL SUMMARY, (b) the Skills category prefix including colon, and (c) 1-2 key JD phrases inside EVERY experience bullet of EVERY company — the second and third company must be bolded exactly as thoroughly as the first. A summary with no bold, or any company whose bullets lack bold, is invalid output. Never bold headings/contact/dates/company/school text.

LENGTH LOCK: keep the same NUMBER OF PARAGRAPHS/LINES as the template — never add or remove lines. Bullets are clear and readable, about 2 rendered lines each.
Length lock: Summary is 105-125 words across exactly six sentences. Each skills line must fit one line. Each experience bullet is ONE natural sentence of 30-38 words that reads easily, contains meaningful technical detail, and typically spans at least two rendered lines in the uploaded template — never a run-on chain of clauses, never stacked buzzwords.

Header/title rules:
- TARGET TITLE: the top title line and every company role line become a concise title describing the role's function for THIS JD, using the formula "Main profession + specialization matching the JD" — e.g. "Software Engineer – AI Platform Infrastructure", "Software Engineer – Banking Product Systems", "AI Engineer – Agentic Clinical Systems", "Backend Engineer – Real-Time Agent Infrastructure", "Software Engineer – CI/CD and Build Infrastructure". Never copy the JD's overly narrow or inflated title verbatim.
- Do not use seniority words (Senior, Junior, Lead, Staff, Principal) in any title.
- Company/role header lines that contain a TAB character (\\t) must keep the tab in place: text before the tab is the role (and company/location when the template puts them together), text after the tab is the date. Never remove the \\t. Never move the date to a new paragraph.
- TEMPLATE SHAPE IS AUTHORITATIVE: some templates use ONE header line ("Role | Company, Location\\tDates"); others use TWO lines — a role+date line ("Software Engineer - AI Product Systems\\tAugust 2024 – Present") followed by a separate company line ("Cisco | USA"). Detect the shape from the ORIGINAL paragraphs and keep it identical for every company.
- A standalone company line such as "Cisco | USA" or "Vivma Software Inc. | India" (company + location, no date) must be returned VERBATIM. Never turn it into a role line, a bullet, or prose.
- The header line stays a short title line (max ~12 words) — NEVER write sentences, summary prose, or bullet content into a header paragraph, and NEVER put a company header into a bullet paragraph. Keep each paragraph in its original role.
- NO-DUPLICATE-HEADER LOCK: each company header line (e.g. "Software Engineer | Vivma Software Inc., India\\tJul 2020 – Dec 2022") appears EXACTLY ONCE, only in its own header paragraph. Never repeat it, or any part of it, inside an experience bullet or any other paragraph. Bullets must never contain " | " with a company name or a date range.


JD analysis (do this silently before writing anything):
- Extract from the JD: exact target job title, seniority, required years, primary and secondary programming languages, frameworks, cloud platforms, databases, DevOps/CI-CD/IaC tooling, containers and Kubernetes, monitoring/observability/logging, incident management, AI/ML and LLM tech, architecture and distributed-systems requirements, security, performance, reliability engineering concepts, testing, Agile methodology, product/leadership/collaboration expectations, soft skills, education, nice-to-haves, and the most frequently repeated keywords.
- Rank every requirement HIGH PRIORITY / MEDIUM / NICE-TO-HAVE, and identify the top 5-10 problems this employer is actually hiring the person to solve (e.g. improving reliability, reducing MTTR, automating operations, building CI/CD, scaling distributed systems, AI-assisted workflows).
- Order the Summary content, the Skills categories, and the experience bullets by those priorities: must-have requirements first, then core responsibilities, then frequently repeated technologies, then architecture, cloud/DevOps, AI, and finally nice-to-haves. Never spend space on content the JD does not care about.

Summary rules (EXACTLY SIX SENTENCES — no more, no less):
- Sentence 1 — experience level + top JD technologies: "<Target discipline> Engineer with 4+ years of experience building production systems using <top 4-6 JD technologies>."
- Sentence 2 — systems and products built: "Experienced in developing <3-4 concrete JD system types, e.g. Kubernetes operators, GitOps platforms, distributed data services, AI-agent infrastructure>."
- Sentence 3 — ownership and production responsibilities: "Owns systems from architecture and implementation through deployment, monitoring, performance optimization, and production operations." (rephrase to the JD's own lifecycle wording; never copy this sentence verbatim every time).
- Sentence 4 — architecture and technical depth: describe one JD-priority design concern, such as scalable application components, data modeling, reliability, security, or performance, using only concepts explicitly named in the JD.
- Sentence 5 — collaboration and delivery: explain how the candidate works with the JD-named partners and contributes to the JD's planning, documentation, review, testing, or delivery practices.
- Sentence 6 — business or user impact: connect the candidate's JD-aligned engineering work to the employer's stated product, customer, operational, or business purpose without inventing metrics.
- Always use "4+ years" only; never any other year count. Keep the six sentences between 105 and 125 words total, with one clear idea per sentence.
- Bold 2-3 top JD keywords. No filler such as "highly motivated", "results-driven", "proven track record".

Skills rules for paragraphs matching "<Category>: <items>":
- EXACTLY SIX CATEGORIES, chosen to support this specific job. Rename the template labels when needed so the six categories match the JD's domain. Reference sets:
  * infrastructure role: Programming; Kubernetes; Cloud and Infrastructure; Distributed Systems; CI/CD; Observability
  * AI role: AI and Machine Learning; Generative AI; Backend Engineering; Data Engineering; Cloud and DevOps; Testing and Observability
  * full-stack role: Frontend; Backend; Applied AI; Data; Cloud and DevOps; Testing and Monitoring
- MANDATORY: never return an empty skills line and never blank a category.
- Output exactly "**Category:** item1, item2, item3". Each category carries 5-7 important keywords for that job. Remove unrelated technologies entirely.
- Prefer JD-explicit skills/tools/methodologies, using the JD's exact terminology and casing ("Datadog", "Prometheus", "GitLab CI/CD"); include acronym + expansion when the JD does, e.g. "Infrastructure as Code (IaC)". If the JD names too few items for a category, complete it with industry-standard skills that fit both the label and the JD's domain.
- Order categories and items by JD priority: the technologies the JD repeats most come first.
- Split compound skills: "React + TypeScript" -> React, TypeScript; "Node.js/Express" -> Node.js, Express. Never keep +, /, &, "and", or "with" joiners.
- Each item is 1-3 words, comma-separated, plain text, no descriptions/parentheticals.


Experience rules:
- BULLET COUNT = 14 PER COMPANY: every company is given 14 bullet paragraph slots — rewrite one unique bullet per slot. If two input bullet slots contain identical template text (duplicated placeholder slots), you MUST still return two completely different bullets. Never blank, merge, transfer, add, or remove an experience bullet, and never output the same bullet text twice.
- Collectively cover every required JD skill/responsibility across the experience section, following the JD's priority order.
- BULLET FORMULA (fixed): WHAT WAS BUILT + TECHNOLOGY + TECHNICAL DEPTH + BUSINESS PURPOSE. Example: "Built a Go-based Kubernetes operator using reconciliation loops, CRDs, finalizers, leader election, and work queues to manage distributed AI workloads." Weak lines such as "Worked with Go and Kubernetes." are invalid.
- FIRST FIVE BULLETS ARE PROJECT-FOCUSED: for every company, bullets 1-5 describe significant systems that were built, not daily responsibilities. Weak: "Responsible for backend development." Strong: "Built a real-time risk-processing platform using Python, Kafka, and PostgreSQL to evaluate customer activity and route suspicious cases for investigation."
- 14-BULLET STRUCTURE — pick the plan matching the JD's role type:
  * platform/infrastructure: 1 primary-language production service; 2 primary-language async/automation; 3 primary-language tooling or data processing; 4 second-language production service; 5 second-language concurrency or operator; 6 second-language performance or CLI; 7 Kubernetes; 8 GitOps; 9 AI/platform infrastructure; 10 networking or edge; 11 observability; 12 databases and messaging; 13 incident response; 14 IaC/Terraform, ownership and collaboration.
  * AI: 1-3 AI and machine learning; 4-6 generative AI and agentic systems; 7-9 backend engineering; 10-14 data, cloud, governance, evaluations, production reliability.
  * full-stack: 1-3 frontend and UI/UX; 4-6 backend services and APIs; 7-10 AI or major product workflows; 11-14 data, deployment, testing, performance, ownership.
- ALTERNATIVES LOCK (critical): when the JD lists interchangeable options (e.g. "React, Vue, Angular, etc.", "PostgreSQL or MySQL", "AWS/GCP/Azure"), pick EXACTLY ONE — the first-listed option — and use only that one across Summary, Skills, Experience and Projects. Never mention the other alternatives anywhere in the resume.
- CORE-SKILL BULLET ALLOCATION (critical): identify the JD's top three core technologies (the ones stated as required fluency, e.g. TypeScript, Node.js, and the chosen front-end framework such as React). For EVERY company, dedicate exactly 3 bullets to each of those three technologies (9 bullets total), each bullet focused on that one technology with real depth, and use the remaining 5 bullets to cover the rest of the JD's skills and responsibilities (data modeling, testing, CI/CD, documentation, collaboration, delivery/monitoring). Never mix two of the three core technologies inside the same bullet.
  * React bullets: components, hooks, state management, rendering performance, forms, routing, accessibility, design-system reuse.
  * TypeScript bullets: typed domain models, generics, discriminated unions, strict mode, shared types across client and server, refactoring safety, lint/type tooling.
  * Node.js bullets: services and REST APIs, middleware, async patterns, streams, authentication, background jobs, error handling, performance.
- TWO IMPORTANT LANGUAGES: when the JD emphasizes two programming languages, give each its own detailed bullets — never mention both superficially in one bullet. Python bullets cover FastAPI/Pydantic, asyncio, decorators, context managers, typing, APIs, automation, pipelines, workers, validation, exception handling, testing. Go bullets cover goroutines, channels, interfaces, contexts, services, controllers, CLIs, graceful shutdown, error wrapping, retries, worker pools, memory use, concurrency. Apply the same split to any other language pair the JD emphasizes.
- WORDING STYLE: ONE clear sentence per bullet, 30-38 words, calm and concrete product/systems-engineering language — no hype, no two-sentence bullets, minimal metrics. Every company bullet must contain enough meaningful detail to span approximately two rendered lines in the uploaded template. A bullet names what was built, the JD technologies used, technical depth, and why it mattered; it must READ NATURALLY, like a human engineer wrote it. Never shorten a bullet into a generic statement, chain four or five noun phrases together, or pad it with filler such as "comprehensive", "dynamic environments", "seamless", "end-to-end", "robust and scalable", "utilizing", "leveraging". Stay strictly inside the JD's technology domain — never introduce tools the JD does not mention. Target voice:
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
- PER-COMPANY UNIQUENESS (critical): the 2nd, 3rd and later companies must NOT reuse a stock bullet set. Recycled generic lines such as "Built Java and TypeScript full-stack applications", "Developed Spring Boot and Node.js backend services", "Built React/Angular interfaces", "Designed MySQL, PostgreSQL and SQL-based data models", "Developed REST and GraphQL integrations", "Implemented asynchronous and event-driven workflows", "Strengthened software quality", "Worked with Product, QA, DevOps and Engineering teams" are BANNED and must never appear.
- Every company's 14 bullets are freshly derived from THIS job description: different JD responsibilities, different JD tools, different sentence shapes than the other companies. No bullet may repeat across companies; no two companies may share the same opening 4 words, the same verb, or the same bullet ordering pattern. Split the JD coverage list across companies: the most recent company takes the JD's senior/architectural and ownership responsibilities, earlier companies take the JD's remaining implementation, tooling, testing, data, automation and collaboration responsibilities - all still 100% JD-sourced with JD keywords verbatim.
- Never fall back to the original resume's old technologies for later companies; if the JD does not name a technology, it must not appear anywhere, including the second and third company.
- TECHNOLOGY WHITELIST (absolute): the ONLY technologies, tools, platforms and techniques allowed anywhere in the resume are those named verbatim in the job description. If the JD does not name it, it is BANNED — including Go, Kafka, Terraform, Kubernetes, Docker, Redis, AWS, GitLab, Prometheus, LLMs, embeddings, vector search, worker pools, message brokers, containers/sandboxes. This applies identically to the summary, skills, EVERY company (especially the 2nd and later) and the PROJECT section. Rewriting an old bullet by keeping its old stack is a failure — rebuild the sentence from JD technologies only.
- The CORE-SKILL BULLET ALLOCATION (3 bullets each for the JD's top three technologies) applies to the 2nd and later companies exactly as it does to the first — never let an earlier company drift into an unrelated stack.
- SENTENCE COMPLETENESS: every bullet must be a complete grammatical sentence ending in a full stop. Never end on a dangling word or preposition ("...in complex.", "...for enhanced product.", "...supporting scalable full-stack."). Keep bullets short enough to finish the thought — never trail off mid-phrase.
- Never begin any line with a quote mark, apostrophe, dash, bullet glyph or stray punctuation.
- METRICS ARE RARE: across each company's 14 bullets, AT MOST 2 bullets total may contain ANY number-based metric, and AT MOST 1 of those may use a % figure (including uptime %). The other 12+ bullets must have ZERO numbers and describe impact qualitatively (improved reliability, simplified onboarding, eliminated manual handoffs, unblocked releases, reduced on-call noise).

- Show ownership and seniority through architecture choices, code reviews, mentoring, cross-team collaboration, production incidents, trade-offs, reliability, cost, velocity, and user outcomes.
- Keyword bolding (distribution matters, applies to ALL companies equally): in EVERY experience bullet of EVERY company, wrap 1-2 short key phrases in **markers** — zero-bold bullets are invalid. Distribute the bolding evenly — never bold the same phrase twice in the resume, never bold the same position in consecutive bullets (alternate between an early-clause tech phrase and a later capability phrase). Bold phrases only (1-4 words), never whole sentences, never metrics, never the leading verb.
- ZERO-REPEAT LOCK (critical): no sentence, clause, or 6-word sequence may ever be reused — not within this resume, and not from any resume you have produced before. Every line must be composed fresh from THIS job description's own wording. If a line feels like something you'd naturally write for any backend/ML job, discard it and rewrite it using specifics unique to this JD (its exact systems, domain, constraints, tools, users). Every generation must read as an entirely different writer's voice.

Projects rules (role-preserving — determine each paragraph's role from its ORIGINAL text, do not assume a fixed template shape):
- CRITICAL: Never invent new paragraphs, never shift content between paragraph slots, and never assume every project has the same structure. Some projects have a tech-stack line, some do not. Some projects have 2 bullets, some have 3, some have more. Match each paragraph's role from its original text.
- Role detection from ORIGINAL text:
  * If the original text is a short title-case phrase with no verb and no " • " separator (e.g. "Enterprise AI Search Platform") → this is a project TITLE. Rewrite to a JD-aligned project name (3-7 words, Title Case, no punctuation except spaces/&) and wrap the ENTIRE new text in ** so it renders bold, exactly like "**Enterprise AI Search & Knowledge Intelligence Platform**".
  * If the original text contains " • " separators joining short tech/tool tokens (e.g. "Python • FastAPI • React") → this is a TECH-STACK line. Rewrite as 4-7 JD-explicit techs joined by " • " and wrap the WHOLE line in **, e.g. "**Python • FastAPI • OpenAI • RAG • PostgreSQL**". Only produce this when the original was itself a tech-stack line — never invent one.
  * If the original text is a full sentence describing an action/outcome → this is a BULLET. Rewrite as ONE clear, natural sentence of 22-28 words: action verb + concrete artifact + JD-specific tech + outcome. Bold 1-2 short inline keyword phrases (1-4 words each) using ** markers — either tech names ("**Python**", "**React**", "**TCP/UDP**") or a key capability phrase ("**client-server architectures**", "**mission critical software**"). Do not bold entire clauses or trailing punctuation.
- Preserve the exact number of bullets each project has in the template — do not add extra bullets to make the section look bigger, and do not drop bullets. If the original project has 2 bullets, output exactly 2 bullets; if it has 3, output 3.
- Do not start project bullets with weak filler. Vary the leading verb across bullets within a project and across projects.
- Collectively, the projects should reinforce JD skills/responsibilities not fully covered by the experience section.

Education rules:
- Do not change the education section's bolding, capitalization, or structure at all. Return every education paragraph — degree lines, school names, locations, dates — VERBATIM from the input. Do not add or remove ** markers on education paragraphs.

FINAL CHECK before returning (re-read every line you wrote and fix violations):
1. Does any line name a technology, tool or platform that is NOT written in the job description (e.g. Prometheus, Grafana, Datadog, Redis, Docker, Kafka, Go, Terraform, Kubernetes)? Remove it and rewrite the sentence with a JD-named technology instead.
2. Does any line contain a banned filler word — seamless, seamlessly, comprehensive, robust, end-to-end, leveraging, leveraged, utilizing, utilized, cutting-edge, best practices, dynamic environments, effective, evolving? Rewrite that line without it.
3. Does every bullet end as a complete sentence with a full stop, and start with a capitalized past-tense verb and no stray punctuation?

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
): Promise<Array<{ i: number; t: string }>> {
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
    "bold mostly technology names in odd-numbered bullets and capability phrases in even-numbered bullets",
    "bold one capability phrase in the first half of each bullet and one tech phrase in the second half, alternating per bullet",
    "bold a single distinct phrase per bullet, alternating between tech stack and domain capability",
  ];
  const boldMode = BOLD_MODES[Math.floor(Math.random() * BOLD_MODES.length)];


  const res = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: "gpt-4.1-mini",
      service_tier: "priority",
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        {
          role: "user",
          content:
            "JOB DESCRIPTION:\n" +
            jobDescription +
            "\n\nRESUME PARAGRAPHS:\n" +
            JSON.stringify({ items }) +
            `\n\nVARIATION SEED: ${Math.random().toString(36).slice(2)}-${Date.now()}\n` +
            `REQUIRED FIRST VERB: ${firstVerb}\n` +
            `STYLE PROFILE FOR THIS GENERATION (apply throughout): ${styleProfile}\n` +
            `BOLD DISTRIBUTION MODE FOR THIS GENERATION: ${boldMode}\n` +
            `PREFERRED VERB ORDER (use roughly in this order, never repeat): ${shuffled.join(", ")}\n` +
            "You MUST return one item for EVERY input index — never omit any index. " +
            extraInstruction +
            "The first experience bullet MUST begin with the REQUIRED FIRST VERB. Apply the STYLE PROFILE and BOLD DISTRIBUTION MODE so this resume reads and looks different from every previous generation. ZERO-REPEAT: no sentence or 6-word sequence may be reused within this resume or match a line you would write for a generic role — rewrite anything that feels reusable. Never open bullets with \"Engineered\" or \"Designed and implemented\" unless listed above. Content must stay strictly JD-accurate. FINAL CHECK BEFORE RESPONDING: re-read the JOB DESCRIPTION and confirm every concrete keyword, tool, technology and stated responsibility in it appears verbatim in your output; then scan your own output for any two lines sharing an opening clause or a repeated bolded phrase and rewrite them; then confirm bolding is spread evenly across bullets.",
        },
      ],
      max_tokens: 32000,
      temperature: 0.85,
      top_p: 0.95,
      frequency_penalty: 0.6,
      presence_penalty: 0.4,
      response_format: { type: "json_object" },
    }),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`OpenAI error ${res.status}: ${text}`);
  }
  const data = await res.json();
  const content: string = data?.choices?.[0]?.message?.content ?? "";
  const parsed = parseLooseJson(content);
  const outItems = Array.isArray(parsed?.items)
    ? parsed.items
    : Array.isArray(parsed)
      ? (parsed as unknown as Array<{ i: number; t: string }>)
      : [];
  return outItems.filter((it) => typeof it?.i === "number" && typeof it?.t === "string");
}

async function callLovableAi(paragraphs: string[], jobDescription: string): Promise<string[]> {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error("OPENAI_API_KEY not configured");

  const items = paragraphs
    .map((t, i) => ({ i, t }))
    .filter((x) => x.t.trim().length > 0);

  if (items.length === 0) return paragraphs.slice();

  const result = paragraphs.slice();
  const filled = new Set<number>();

  const apply = (out: Array<{ i: number; t: string }>) => {
    for (const it of out) {
      if (it.i >= 0 && it.i < result.length && it.t.trim().length > 0) {
        result[it.i] = it.t;
        filled.add(it.i);
      }
    }
  };

  apply(await requestRewrite(items, jobDescription, apiKey, ""));

  // Second pass: re-request any indices the model silently skipped, instead of
  // falling back to the original template boilerplate.
  for (let attempt = 0; attempt < 2; attempt++) {
    const missing = items.filter((x) => !filled.has(x.i));
    if (missing.length === 0) break;
    try {
      apply(
        await requestRewrite(
          missing,
          jobDescription,
          apiKey,
          "These are the ONLY remaining paragraphs from the same resume; they were skipped previously. Rewrite each one fully JD-aligned, keeping each paragraph's original role (header, bullet, skills line, project title). Do not return template boilerplate. ",
        ),
      );
    } catch {
      break;
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
    return { jobDescription: jd, templateBase64: input.templateBase64 };
  })
  .handler(async ({ data }) => {
    const bytes = new Uint8Array(Buffer.from(data.templateBase64, "base64"));

    const out = await rewriteDocx(bytes, data.jobDescription, (paragraphs, jd) =>
      callLovableAi(paragraphs, jd),
    );

    const base64 = Buffer.from(out).toString("base64");
    return {
      fileName: `Yathendra_Resume.docx`,
      base64,
    };
  });
