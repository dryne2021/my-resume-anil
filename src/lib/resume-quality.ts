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
