// Quality checks for generated resume text. No dependencies so it can be tested on its own.

const BANNED_RE =
  /\b(seamless(?:ly)?|comprehensive|robust|end-to-end|leverag(?:e|ed|es|ing)|utiliz(?:e|ed|es|ing)|cutting-edge|best practices|dynamic environments?|spearhead(?:ed)?|revolutioniz(?:e|ed)|results-driven|proven track record|highly motivated)\b/gi;

const CONNECTORS = new Set(
  "a an the to for with by of in on and or that so from across into through during while which using via under at as when without before after its their".split(" "),
);

export const plain = (t: string) => t.replace(/\*\*/g, "").trim();

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

  result.forEach((raw, i) => {
    const t = plain(raw);
    if (!t || t === plain(originals[i] ?? "")) return;

    const banned = Array.from(new Set((t.match(BANNED_RE) ?? []).map((w) => w.toLowerCase())));
    if (banned.length) add(i, `remove the banned word(s): ${banned.join(", ")}`);

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
    if (words.length > 42) add(i, `too long (${words.length} words); keep it to 30-38 words`);
    if (words.length < 18) add(i, `too short (${words.length} words); add the mechanism and why it mattered`);
  });
  return problems;
}
