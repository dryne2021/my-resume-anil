import PizZip from "pizzip";

/**
 * Resume content rewriter that preserves the original .docx formatting.
 *
 * - Extracts text from each <w:p> paragraph.
 * - Sends paragraphs to AI for rewriting (indices preserved).
 * - Writes rewritten text back into the first <w:r> of each paragraph,
 *   cloning its run properties (rPr) so fonts/sizes/colors carry over.
 * - Supports inline **bold** markers from the AI: these become real bold
 *   runs in the .docx by cloning the run's rPr and adding <w:b/>.
 */

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function decodeXml(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

interface Block {
  start: number;
  end: number;
  text: string;
  /** True when the paragraph is a numbered/bulleted list item in the template. */
  isList: boolean;
}

function extractParagraphs(xml: string): Block[] {
  const blocks: Block[] = [];
  const pOpenRe = /<w:p(?:\s[^>]*)?>/g;
  let m: RegExpExecArray | null;
  while ((m = pOpenRe.exec(xml)) !== null) {
    // Skip self-closing empty paragraphs like `<w:p .../>` — they have no
    // matching `</w:p>` and merging them into the next block corrupts XML.
    if (xml[m.index + m[0].length - 2] === "/") continue;
    const start = m.index;
    const closeIdx = xml.indexOf("</w:p>", start);
    if (closeIdx === -1) continue;
    const end = closeIdx + "</w:p>".length;
    const block = xml.slice(start, end);
    // Walk text runs and tab markers in document order so that layout tabs
    // (used for right-aligned dates in resume headers) survive the rewrite.
    const tokenRe =
      /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>/g;
    let text = "";
    let tm: RegExpExecArray | null;
    while ((tm = tokenRe.exec(block)) !== null) {
      if (tm[1] !== undefined) text += decodeXml(tm[1]);
      else text += "\t";
    }
    blocks.push({ start, end, text, isList: /<w:numPr\b/.test(block) });
  }
  return blocks;
}

/** Sentinel: a paragraph marked with this is removed from the document. */
const DELETE_PARAGRAPH = "\u0000__DELETE_PARAGRAPH__";

/**
 * BULLET COUNT LOCK — guarantees every experience bullet group has exactly
 * `target` bullet paragraphs. Groups shorter than `target` (and long enough to
 * be a real experience group) get their last bullet paragraph duplicated, so
 * the rewrite pass has enough slots to fill per company. Groups already at or
 * above the target are left untouched.
 */
function expandBulletGroups(xml: string, target: number): string {
  const blocks = extractParagraphs(xml);
  type Group = { blocks: Block[] };
  const groups: Group[] = [];
  let current: Block[] = [];
  for (const b of blocks) {
    if (b.isList) current.push(b);
    else if (current.length) {
      groups.push({ blocks: current });
      current = [];
    }
  }
  if (current.length) groups.push({ blocks: current });

  let out = xml;
  // Right-to-left so earlier offsets stay valid.
  for (let g = groups.length - 1; g >= 0; g--) {
    const items = groups[g].blocks;
    if (items.length >= target || items.length < 6) continue;
    const last = items[items.length - 1];
    const clone = out.slice(last.start, last.end);
    const extra = clone.repeat(target - items.length);
    out = out.slice(0, last.end) + extra + out.slice(last.end);
  }
  return out;
}


function parseBoldSegments(text: string): { text: string; bold: boolean }[] {
  const segs: { text: string; bold: boolean }[] = [];
  const re = /\*\*([\s\S]+?)\*\*/g;
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) segs.push({ text: text.slice(last, m.index), bold: false });
    segs.push({ text: m[1], bold: true });
    last = m.index + m[0].length;
  }
  if (last < text.length) segs.push({ text: text.slice(last), bold: false });
  if (segs.length === 0) segs.push({ text, bold: false });
  return segs;
}

function injectBold(rPr: string): string {
  if (!rPr) return "<w:rPr><w:b/><w:bCs/></w:rPr>";
  if (/<w:b\s*\/?>/i.test(rPr)) return rPr;
  return rPr.replace(/<\/w:rPr>/, "<w:b/><w:bCs/></w:rPr>");
}

function stripBold(rPr: string): string {
  if (!rPr) return "";
  const cleaned = rPr
    .replace(/<w:b\s*\/>/gi, "")
    .replace(/<w:bCs\s*\/>/gi, "")
    .replace(/<w:b\s+[^/]*\/>/gi, "")
    .replace(/<w:bCs\s+[^/]*\/>/gi, "");
  // If rPr is now empty, drop it
  if (/<w:rPr>\s*<\/w:rPr>/.test(cleaned)) return "";
  return cleaned;
}

function rewriteParagraphXml(block: string, newText: string): string {
  // An intentionally empty rewrite must clear every text run. Previously the
  // fallback below restored the original run, leaving unsupported skill labels
  // such as "Security & Compliance:" visible in the generated document.
  if (newText.length === 0) {
    return block
      .replace(/<w:t(\s[^>]*)?>[\s\S]*?<\/w:t>/g, "<w:t></w:t>")
      .replace(/<w:tab\s*\/>/g, "");
  }

  // Split around tab characters so we can emit <w:tab/> at the same positions.
  const tabParts = newText.split("\t");

  // Find first <w:r ...> ... </w:r> that contains a <w:t>
  const firstRunRe =
    /<w:r(?:\s[^>]*)?>[\s\S]*?<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>[\s\S]*?<\/w:r>/;
  const m = firstRunRe.exec(block);

  if (!m) {
    // Fallback: replace inside first <w:t>
    let first = true;
    return block
      .replace(/<w:t(\s[^>]*)?>[\s\S]*?<\/w:t>/g, (_match, attrs) => {
        if (first) {
          first = false;
          const a =
            attrs && attrs.includes("xml:space") ? attrs : ' xml:space="preserve"';
          return `<w:t${a}>${escapeXml(newText.replace(/\t/g, " "))}</w:t>`;
        }
        return `<w:t></w:t>`;
      })
      .replace(/<w:tab\s*\/>/g, "");
  }

  const firstRun = m[0];
  const before = block.slice(0, m.index);
  const after = block.slice(m.index + firstRun.length);
  const rPrMatch = firstRun.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  const rPr = rPrMatch ? rPrMatch[0] : "";

  const renderPart = (part: string): string => {
    const segments = parseBoldSegments(part);
    return segments
      .filter((s) => s.text.length > 0)
      .map((seg) => {
        const rPrXml = seg.bold ? injectBold(rPr) : stripBold(rPr);
        return `<w:r>${rPrXml}<w:t xml:space="preserve">${escapeXml(
          seg.text,
        )}</w:t></w:r>`;
      })
      .join("");
  };

  const tabRun = `<w:r>${rPr}<w:tab/></w:r>`;
  const newRuns = tabParts
    .map((p) => renderPart(p))
    .reduce((acc, cur, i) => (i === 0 ? cur : acc + tabRun + cur), "");

  // Empty all <w:t> contents and remove leftover <w:tab/> elements from the
  // rest of the paragraph so we don't duplicate tabs.
  const cleanedAfter = after
    .replace(/<w:t(\s[^>]*)?>[\s\S]*?<\/w:t>/g, "<w:t></w:t>")
    .replace(/<w:tab\s*\/>/g, "");

  return before + (newRuns || firstRun) + cleanedAfter;
}


export interface AiRewriteFn {
  (paragraphs: string[], jobDescription: string): Promise<string[]>;
}

export async function rewriteDocx(
  templateBytes: Uint8Array,
  jobDescription: string,
  aiRewrite: AiRewriteFn,
): Promise<Uint8Array> {
  const zip = new PizZip(templateBytes);
  const docXmlFile = zip.file("word/document.xml");
  if (!docXmlFile) throw new Error("Invalid .docx: missing word/document.xml");
  // BULLET COUNT LOCK — every experience company must expose 14 bullet
  // slots. Templates that ship fewer get the last bullet paragraph cloned.
  const originalXml = expandBulletGroups(docXmlFile.asText(), 14);

  const blocks = extractParagraphs(originalXml);
  const paragraphs = blocks.map((b) => b.text);


  const rewritten = await aiRewrite(paragraphs, jobDescription);
  if (rewritten.length !== paragraphs.length) {
    throw new Error(
      `AI returned ${rewritten.length} paragraphs but template has ${paragraphs.length}.`,
    );
  }

  // =====================================================================
  // FORMAT LOCK — server-side guardrails that run on every generation to
  // guarantee identical structure across all resumes, regardless of what the
  // AI returns. Do not weaken these without explicit user approval.
  // =====================================================================

  const isAllCapsHeading = (s: string): boolean => {
    const t = s.trim();
    if (t.length < 3 || t.length > 60) return false;
    if (!/[A-Z]/.test(t)) return false;
    return /^[A-Z0-9 &\/\-]+$/.test(t);
  };
  const isEducationLine = (s: string): boolean => {
    if (!s.includes(" | ")) return false;
    return /\b(Master|Bachelor|MSc|BSc|MBA|Ph\.?D|B\.?E\.?|M\.?E\.?|B\.?Tech|M\.?Tech|University|College|Institute|School)\b/i.test(
      s,
    );
  };
  const stripBoldMarkers = (s: string) => s.replace(/\*\*/g, "");
  const isTechStackLine = (s: string) => stripBoldMarkers(s).includes(" • ");
  const isSkillsLine = (s: string) => {
    const t = stripBoldMarkers(s).trim();
    // "Category: item1, item2" — short label before first colon.
    const idx = t.indexOf(":");
    if (idx <= 0 || idx > 40) return false;
    const label = t.slice(0, idx);
    return /^[A-Za-z][A-Za-z0-9 &\/\-]*$/.test(label) && t.slice(idx + 1).trim().length > 0;
  };
  // Company/role header lines. Some templates right-align the date with a tab,
  // others with a run of spaces — both must be detected as headers, otherwise
  // the slot gets filled with prose and the company loses its heading.
  const DATE_TAIL_RE =
    /(?:\t|\s{2,}|\s(?=(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*\d{4}\b))(?:(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*\d{4}|\d{4})\s*[–—-]\s*(?:Present|Current|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*\d{4}|\d{4})\s*$/i;
  const headerTailIndex = (s: string): number => {
    const m = DATE_TAIL_RE.exec(s);
    return m ? m.index : -1;
  };
  const isHeaderRoleLine = (s: string) =>
    s.includes(" | ") && (s.includes("\t") || headerTailIndex(s) > 0);
  const isProjectTitleLine = (s: string) => {
    const t = stripBoldMarkers(s).trim();
    if (t.length === 0 || t.length > 80) return false;
    if (/[.!?]/.test(t)) return false;
    if (t.includes(" • ") || t.includes(":") || t.includes("\t") || t.includes(" | ")) return false;
    const words = t.split(/\s+/);
    if (words.length < 2 || words.length > 10) return false;
    // Mostly capitalized words (title case) — no verb-led sentence.
    const capish = words.filter((w) => /^[A-Z0-9&]/.test(w)).length;
    return capish >= Math.ceil(words.length * 0.6);
  };

  // Company/role header detection for paragraphs that are NOT header slots.
  const DATE_RE =
    /\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*\d{4}\b|\b\d{4}\s*[–—-]\s*(Present|\d{4})\b/i;
  const norm = (s: string) =>
    stripBoldMarkers(s).replace(/\s+/g, " ").trim().toLowerCase();
  const originalHeaders = paragraphs.filter((p) => isHeaderRoleLine(p)).map((p) => norm(p));
  const originalHeaderHeads = paragraphs
    .filter((p) => isHeaderRoleLine(p))
    .map((p) => norm(p.slice(0, p.indexOf("\t"))))
    .filter((h) => h.length > 6);
  const looksLikeCompanyHeaderText = (s: string): boolean => {
    const t = stripBoldMarkers(s).trim();
    if (t.length === 0) return false;
    if (!t.includes(" | ")) return false;
    if (isEducationLine(t)) return false;
    // A bullet slot containing "Role | Company" plus a date range is a duplicated header.
    return DATE_RE.test(t) || t.includes("\t");
  };
  const matchesAnyOriginalHeader = (s: string): boolean => {
    const t = norm(s);
    if (t.length === 0) return false;
    if (originalHeaders.some((h) => t === h || t.includes(h))) return true;
    return originalHeaderHeads.some((h) => t === h || t.startsWith(h) || t.includes(h));
  };

  // Universal text sanitizer — removes markdown/formatting artifacts the model
  // sometimes emits so every generated resume matches the template exactly.
  const sanitize = (s: string): string => {
    if (!s) return "";
    let t = s;
    t = t.replace(/\r/g, "");
    t = t.replace(/\n+/g, " ");
    // Strip markdown heading marks and stray list glyphs at line start.
    t = t.replace(/^\s*#{1,6}\s*/, "");
    t = t.replace(/^\s*(?:[•·▪◦*\-–—]|\d{1,2}[.)])\s+/, "");
    // Drop leftover markdown emphasis that isn't our ** bold markers.
    t = t.replace(/(^|[^*])\*(?!\*)/g, "$1");
    t = t.replace(/`+/g, "");
    // Balance bold markers — an odd count would leak literal asterisks.
    const stars = (t.match(/\*\*/g) || []).length;
    if (stars % 2 !== 0) t = t.replace(/\*\*([^*]*)$/, "$1");
    // Remove empty bold pairs and normalize whitespace.
    t = t.replace(/\*\*\s*\*\*/g, "");
    t = t.replace(/[ \u00a0]{2,}/g, " ");
    t = t.replace(/ ?\t ?/g, "\t");
    t = t.replace(/\s+([,.;:])/g, "$1");
    return t.trim();
  };

  // Track section context based on ORIGINAL paragraphs.
  let inEducation = false;
  let inExperience = false;
  let inSummary = false;

  for (let i = 0; i < paragraphs.length; i++) {
    const orig = paragraphs[i];
    const origTrim = orig.trim();
    const isListSlot = blocks[i]?.isList === true;

    // 1) Empty paragraphs stay empty — never let the AI fill spacer lines.
    if (origTrim.length === 0) {
      rewritten[i] = orig;
      continue;
    }

    // 2) ALL-CAPS section headings — verbatim.
    if (isAllCapsHeading(orig)) {
      rewritten[i] = orig;
      inEducation = /EDUCATION/i.test(orig);
      inExperience = /EXPERIENCE|EMPLOYMENT|WORK HISTORY/i.test(orig);
      inSummary = /SUMMARY|PROFILE|OBJECTIVE/i.test(orig);
      continue;
    }

    // 3) Education section — every line verbatim, including school/degree lines.
    if (inEducation || isEducationLine(orig)) {
      rewritten[i] = orig;
      continue;
    }

    const line = sanitize(rewritten[i] ?? "");

    // 4) Company/role header lines — role head bolded, date tail preserved
    // verbatim from the ORIGINAL (tab- or space-aligned, whichever it used).
    // Structural detection wins over text heuristics: inside EXPERIENCE, any
    // template paragraph that is NOT a list item but carries "Role | Company"
    // or a tab-aligned date IS a company header slot.
    const isStructuralHeader =
      inExperience && !isListSlot && (orig.includes(" | ") || orig.includes("\t"));
    if (isHeaderRoleLine(orig) || isStructuralHeader) {
      const tabIdx = orig.indexOf("\t");
      const rawSplitIdx = tabIdx >= 0 ? tabIdx : headerTailIndex(orig);
      const origSplitIdx = rawSplitIdx > 0 ? rawSplitIdx : orig.length;
      const origTail = orig.slice(origSplitIdx); // tab/spaces + date, keep exactly.
      const origHead = stripBoldMarkers(orig.slice(0, origSplitIdx)).trim();
      let head: string;
      const lineSplitIdx = line.includes("\t") ? line.indexOf("\t") : headerTailIndex(line);
      if (lineSplitIdx > 0) head = line.slice(0, lineSplitIdx);
      else head = line.trim().length > 0 ? line : origHead;
      head = stripBoldMarkers(head).trim();
      // Guard: header slots must stay short title lines. If the AI wrote prose
      // (sentences, too many words, or a much longer line), restore the
      // original header so every company keeps the same shape.
      const wordCount = head.split(/\s+/).filter(Boolean).length;
      const looksLikeProse =
        /[.!?]\s/.test(head) ||
        /[.!?]$/.test(head) ||
        wordCount > 14 ||
        (origHead.includes(" | ") && !head.includes(" | ")) ||
        head.length > origHead.length * 1.6;
      if (looksLikeProse) head = origHead;
      if (head.length === 0) head = origHead;
      rewritten[i] = `**${head}**${origTail}`;
      continue;
    }

    // 5) Skills lines "Category: items" — force "**Category:** items".
    // The AI may rename the category when the template label does not fit the
    // JD's domain (e.g. "Frontend Engineering: Python, C++"); accept a short,
    // well-formed replacement label, otherwise keep the original label.
    if (isSkillsLine(orig)) {
      const origPlain = stripBoldMarkers(orig);
      const origLabel = origPlain.slice(0, origPlain.indexOf(":")).trim();
      const cleaned = stripBoldMarkers(line);
      const colonIdx = cleaned.indexOf(":");
      const items = colonIdx >= 0 ? cleaned.slice(colonIdx + 1).trim() : cleaned.trim();
      let label = origLabel;
      if (colonIdx > 0) {
        const aiLabel = cleaned.slice(0, colonIdx).trim();
        const words = aiLabel.split(/\s+/).filter(Boolean);
        const ok =
          aiLabel.length > 2 &&
          aiLabel.length <= 40 &&
          words.length <= 4 &&
          !/[.!?]/.test(aiLabel) &&
          !/,/.test(aiLabel);
        if (ok) label = aiLabel;
      }
      if (items.length === 0) {
        // AI blanked it — keep original verbatim rather than emit empty.
        rewritten[i] = orig;
      } else {
        rewritten[i] = `**${label}:** ${items}`;
      }
      continue;
    }


    // 6) Tech-stack lines (original contains " • ") — whole line bolded, joined by " • ".
    if (isTechStackLine(orig)) {
      const cleaned = stripBoldMarkers(line).trim();
      if (cleaned.length === 0) {
        rewritten[i] = orig;
      } else {
        // Normalize separators to " • " and re-wrap.
        const parts = cleaned
          .split(/\s*•\s*/)
          .map((p) => p.trim())
          .filter((p) => p.length > 0);
        rewritten[i] = `**${parts.join(" • ")}**`;
      }
      continue;
    }

    // 7) Project title lines — whole line wrapped in ** so it renders bold.
    if (isProjectTitleLine(orig)) {
      const cleaned = stripBoldMarkers(line).trim();
      const title = cleaned.length > 0 ? cleaned : stripBoldMarkers(orig).trim();
      rewritten[i] = `**${title}**`;
      continue;
    }

    // 8) Anti-duplication: a non-header paragraph (bullet, summary, project line)
    // must never contain a company/role header line. If the AI echoed a header
    // into a bullet slot, restore the original paragraph text.
    if (looksLikeCompanyHeaderText(line) || matchesAnyOriginalHeader(line)) {
      rewritten[i] = orig;
      continue;
    }

    // 9) Body paragraphs (bullets, summary, project bullets): sanitized text,
    // no stray tabs (tabs only belong in role header lines), never blank —
    // a blank body slot would collapse the template's bullet layout.
    let body = line.replace(/\t/g, " ").replace(/ {2,}/g, " ").trim();
    if (body.length === 0) body = orig;

    // SENTENCE-START LOCK — a bullet must never begin mid-sentence (e.g.
    // "js backend services implementing REST APIs..."). Repair the fragment by
    // dropping it up to the next capitalized verb-like word; if nothing usable
    // remains, restore the original template paragraph.
    {
      const plain = stripBoldMarkers(body).trim();
      const firstWord = plain.split(/\s+/)[0] ?? "";
      const badStart =
        firstWord.length > 0 &&
        (/^[a-z0-9]/.test(firstWord) || /^[•·▪◦\-–—,;:.)]/.test(firstWord));
      if (badStart) {
        // Common cause: a technology name lost its head ("Node.js" -> "js ...").
        // Common causes: a technology name lost its head ("Node.js" -> "js ...",
        // "OAuth 2.0" -> "0 ...").
        let repaired = plain.replace(/^(js|ts|py|net|NET|\d+(\.\d+)?)\b[.\s]*/i, "").trim();
        if (/^[a-z0-9]/.test(repaired) || repaired.split(/\s+/).length < 8) {
          const m2 = plain.match(/\b([A-Z][a-z]{2,})\b/);
          repaired = m2 ? plain.slice(plain.indexOf(m2[0])).trim() : "";
        }
        if (repaired.split(/\s+/).length >= 8) {
          // Still starts with a noun/tech word — give it a proper action verb.
          body = /^[A-Z][a-z]+ed\b/.test(repaired) ? repaired : `Built ${repaired.charAt(0).toLowerCase()}${repaired.slice(1)}`;
        } else {
          body = orig;
        }
      }
    }


    // LENGTH LOCK — experience list items are intentionally detailed enough to
    // span about two rendered lines. Non-list paragraphs retain the tighter
    // template-relative limit. Trim at a sentence boundary first, then at a
    // word boundary, keeping bold markers balanced.
    const budget = isListSlot
      ? Math.max(300, Math.round(stripBoldMarkers(orig).length * 1.45))
      : inSummary
        ? Math.max(850, Math.round(stripBoldMarkers(orig).length * 2.1))
      : Math.max(200, Math.round(stripBoldMarkers(orig).length * 1.15));
    if (stripBoldMarkers(body).length > budget) {
      const sentences = body.match(/[^.!?]+[.!?]+(\s|$)/g) ?? [];
      let acc = "";
      for (const s of sentences) {
        if (stripBoldMarkers(acc + s).length > budget) break;
        acc += s;
      }
      let trimmed = acc.trim();
      if (stripBoldMarkers(trimmed).length < budget * 0.5) {
        // No usable sentence boundary — hard-trim on words.
        const words = body.split(" ");
        trimmed = "";
        for (const w of words) {
          const next = trimmed ? `${trimmed} ${w}` : w;
          if (stripBoldMarkers(next).length > budget) break;
          trimmed = next;
        }
        // Drop dangling connectors/modifiers so the sentence never trails off
        // ("...in complex.", "...for enhanced product.", "...scalable full-stack.").
        const DANGLING =
          /\s+(?:a|an|the|and|or|of|for|to|in|on|at|by|with|within|across|through|into|from|that|which|while|using|complex|enhanced|scalable|improved|overall|full-stack|advanced|other|their|its)$/i;
        for (let g = 0; g < 4; g++) {
          const next = trimmed.replace(/[,;:\-–—]$/, "").trim().replace(DANGLING, "");
          if (next === trimmed) break;
          trimmed = next;
        }
        trimmed = trimmed.replace(/[,;:\-–—]$/, "").trim();
        if (trimmed && !/[.!?]$/.test(trimmed)) trimmed += ".";
      }
      if (trimmed.length > 0) body = sanitize(trimmed);
    }

    // BOLD LOCK — summary and experience/project bullets must always carry at
    // least one bolded keyword. If the model returned none, bold the first
    // proper-noun / tech-looking phrase in the sentence (never the lead verb).
    if (!body.includes("**") && body.split(/\s+/).length > 6) {
      const words = body.split(" ");
      for (let w = 1; w < words.length; w++) {
        const raw = words[w];
        const core = raw.replace(/^[^A-Za-z0-9+#./]+|[^A-Za-z0-9+#./]+$/g, "");
        if (core.length < 3) continue;
        if (!/[A-Z0-9]/.test(core)) continue;
        if (/^(The|A|An|And|With|For|To|In|On|Of|By|Across)$/i.test(core)) continue;
        const next = words[w + 1]?.replace(/^[^A-Za-z0-9+#./]+|[^A-Za-z0-9+#./]+$/g, "") ?? "";
        const phrase = /^[A-Z]/.test(next) && next.length > 2 ? `${core} ${next}` : core;
        const idx = body.indexOf(phrase);
        if (idx >= 0) {
          body = body.slice(0, idx) + `**${phrase}**` + body.slice(idx + phrase.length);
        }
        break;
      }
    }
    rewritten[i] = body;

  }

  // TWO-LINE COMPANY HEADER LOCK — validate + repair the experience header
  // shape against the template before writing anything back to the document.
  enforceTwoLineCompanyHeaders(paragraphs, rewritten, blocks);

  // Apply replacements right-to-left so offsets stay valid
  let newXml = originalXml;

  for (let i = blocks.length - 1; i >= 0; i--) {
    const b = blocks[i];
    const newText = rewritten[i] ?? "";
    // Duplicate bullet slots are dropped from the document entirely.
    if (newText === DELETE_PARAGRAPH) {
      newXml = newXml.slice(0, b.start) + newXml.slice(b.end);
      continue;
    }
    if (b.text === "" && newText === "") continue;
    if (b.text === newText) continue;
    const original = newXml.slice(b.start, b.end);
    const replaced = rewriteParagraphXml(original, newText);
    newXml = newXml.slice(0, b.start) + replaced + newXml.slice(b.end);
  }

  newXml = collapseEmptyBeforeHeading(newXml, "PROFESSIONAL EXPERIENCE");

  zip.file("word/document.xml", newXml);
  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

/**
 * Remove empty <w:p> paragraphs immediately preceding a heading paragraph
 * whose text starts with `headingText`. Templates often include blank
 * paragraphs to space sections; those look like an oversized gap once
 * the content above them has been rewritten to be shorter.
 */
function collapseEmptyBeforeHeading(xml: string, headingText: string): string {
  const blocks = extractParagraphs(xml);
  const headingIdx = blocks.findIndex((b) =>
    b.text.trim().toUpperCase().startsWith(headingText.toUpperCase()),
  );
  if (headingIdx <= 0) return xml;

  let firstEmpty = headingIdx;
  for (let i = headingIdx - 1; i >= 0; i--) {
    if (blocks[i].text.trim() === "") firstEmpty = i;
    else break;
  }
  if (firstEmpty === headingIdx) return xml;

  const removeStart = blocks[firstEmpty].start;
  const removeEnd = blocks[headingIdx - 1].end;
  return xml.slice(0, removeStart) + xml.slice(removeEnd);
}

/**
 * Validator: enforces the uploaded template's TWO-LINE company header format
 * in every generated DOCX.
 *
 * Shape (per company, taken from the template itself):
 *   line 1: **Role – Focus | ...**<tab or spaces><Start – End>   (date tail verbatim)
 *   line 2: **Company | Country**                                 (verbatim, no date)
 *
 * Repairs performed:
 *  - the standalone company/location line is restored verbatim from the template
 *    (the model must never rewrite, merge, or blank it);
 *  - the role line keeps the template's exact date tail and stays bolded;
 *  - the company line is never duplicated into a bullet or the role line.
 *
 * Templates that use a ONE-line header are left untouched.
 */
function enforceTwoLineCompanyHeaders(
  originals: string[],
  rewritten: string[],
  blocks: { isList?: boolean }[],
): void {
  const strip = (s: string) => (s ?? "").replace(/\*\*/g, "");
  const norm = (s: string) => strip(s).replace(/\s+/g, " ").trim().toLowerCase();
  const DATE = /\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*\d{4}\b|\b\d{4}\s*[–—-]\s*(?:Present|Current|\d{4})\b/i;

  const isRoleLine = (s: string) => {
    const t = strip(s);
    return t.includes(" | ") || (t.includes("\t") && DATE.test(t));
  };
  // Standalone company line: "Cisco", "Vivma Software Inc. | India" — short,
  // not a list item, carries no date range and no sentence punctuation.
  const isCompanyLine = (s: string, isList?: boolean) => {
    const t = strip(s).trim();
    if (isList) return false;
    if (t.length === 0) return false;
    if (DATE.test(t) || t.includes("\t")) return false;
    if (/[!?]/.test(t)) return false;
    if (!/^[A-Z0-9]/.test(t)) return false;
    const words = t.split(/\s+/).filter(Boolean).length;
    if (words < 1 || words > 10 || t.length > 70) return false;
    if (t.includes(" | ")) return true;
    if (/^[A-Z0-9 &/-]+$/.test(t)) return false; // ALL-CAPS section heading
    if (/[—–]/.test(t)) return false; // project title style "Name — Subtitle"
    // No pipe: a bare company name (no sentence-ending period, few words).
    return words <= 6 && !/\.$/.test(t.replace(/\b(Inc|Ltd|LLC|Corp)\.$/i, ""));
  };

  const companyLines: string[] = [];

  // Collect EVERY standalone company/location line in the template, whether or
  // not it sits directly under a detected role line — these may never be
  // echoed into a bullet.
  for (let i = 0; i < originals.length; i++) {
    if (isCompanyLine(originals[i] ?? "", blocks[i]?.isList)) {
      companyLines.push(norm(originals[i]));
    }
  }

  for (let i = 0; i < originals.length; i++) {
    const roleOrig = originals[i];
    if (!roleOrig || !isRoleLine(roleOrig) || !DATE.test(strip(roleOrig))) continue;

    // Locate the template's standalone company line right after the role line.
    let j = i + 1;
    while (j < originals.length && (originals[j] ?? "").trim() === "") j++;
    if (j >= originals.length) continue;
    const companyOrig = originals[j];
    if (!isCompanyLine(companyOrig, blocks[j]?.isList)) continue; // one-line template

    // 1) Company line: verbatim from the template, bold preserved as authored.
    rewritten[j] = companyOrig;

    // 2) Role line: keep the template's exact date tail; never absorb the company line.
    const tabIdx = roleOrig.indexOf("\t");
    const splitIdx = tabIdx >= 0 ? tabIdx : roleOrig.search(/\s{2,}(?=\S*\d{4})/);
    const tail = splitIdx > 0 ? roleOrig.slice(splitIdx) : "";
    const origHead = strip(roleOrig.slice(0, splitIdx > 0 ? splitIdx : roleOrig.length)).trim();

    let out = rewritten[i] ?? "";
    const outTabIdx = out.indexOf("\t");
    let head = strip(outTabIdx > 0 ? out.slice(0, outTabIdx) : out).trim();
    head = head.replace(/\s*\|\s*$/, "").trim();
    // Strip stray leading punctuation the model sometimes emits ("'Software Developer …").
    head = head.replace(/^['"`’‘“”\-–—•*]+\s*/, "").trim();
    const companyPlain = strip(companyOrig).trim();
    if (companyPlain && head.toLowerCase().includes(companyPlain.toLowerCase())) {
      head = origHead; // merged the two lines back together — revert.
    }
    if (head.length === 0 || DATE.test(head)) head = origHead;
    rewritten[i] = tail ? `**${head}**${tail}` : `**${head}**`;
  }

  // 3) NO-REPEAT LOCK — no bullet may repeat a company line, and no two bullets
  // may carry identical text. Offending bullet paragraphs are deleted outright
  // (reverting them would just re-emit the duplicate).
  const seen = new Set<string>();
  for (let i = 0; i < rewritten.length; i++) {
    if (!blocks[i]?.isList) continue;
    const t = norm(rewritten[i] ?? "");
    if (t.length === 0) continue;
    const repeatsCompany = companyLines.some(
      (c) => c.length > 6 && (t === c || t.startsWith(c)),
    );
    if (repeatsCompany || seen.has(t)) {
      rewritten[i] = DELETE_PARAGRAPH;
      continue;
    }
    seen.add(t);
  }
}

