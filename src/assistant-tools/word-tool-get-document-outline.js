import { Type } from "typebox";

/** @type {readonly ("outlineLevel" | "styleBuiltIn")[]} */
const META_PROPS = ["outlineLevel", "styleBuiltIn"];
/** @type {readonly ("text")[]} */
const TEXT_PROPS = ["text"];

const CHUNK = 80;

/**
 * @param {number | undefined} outlineLevel
 * @param {string | undefined} styleBuiltIn
 * @returns {number | null} Heading level 1–9, or null if not a TOC-style heading.
 */
function headingLevelFromParagraph(outlineLevel, styleBuiltIn) {
  if (typeof outlineLevel === "number" && outlineLevel >= 1 && outlineLevel <= 9) {
    return outlineLevel;
  }
  if (typeof styleBuiltIn === "string") {
    const m = /^Heading([1-9])$/.exec(styleBuiltIn);
    if (m) {
      return Number(m[1]);
    }
  }
  return null;
}

/**
 * @param {string} raw
 * @param {number} maxLen
 */
function normalizeHeadingText(raw, maxLen) {
  const t = typeof raw === "string" ? raw.replace(/\s+/g, " ").trim() : "";
  if (!t) {
    return "(empty heading)";
  }
  if (t.length <= maxLen) {
    return t;
  }
  return `${t.slice(0, Math.max(0, maxLen - 1)).trimEnd()}…`;
}

/**
 * @param {import("office-js").Word.RequestContext} context
 * @param {import("office-js").Word.Paragraph[]} items
 * @param {readonly string[]} props
 */
async function loadParagraphPropsInChunks(context, items, props) {
  for (let i = 0; i < items.length; i += CHUNK) {
    const end = Math.min(i + CHUNK, items.length);
    for (let j = i; j < end; j++) {
      items[j].load(props);
    }
    await context.sync();
  }
}

/**
 * @param {import("office-js").Word.Paragraph[]} items
 */
function scanHeadingParagraphs(items) {
  /** @type {number[]} */
  const headingIndices = [];
  /** @type {(number | null)[]} */
  const levels = [];
  for (let i = 0; i < items.length; i++) {
    const level = headingLevelFromParagraph(items[i].outlineLevel, items[i].styleBuiltIn);
    levels.push(level);
    if (level !== null) {
      headingIndices.push(i);
    }
  }
  return { headingIndices, levels };
}

/**
 * @param {import("office-js").Word.RequestContext} context
 * @param {import("office-js").Word.Paragraph[]} items
 * @param {readonly number[]} indices
 */
async function loadTextsForParagraphIndices(context, items, indices) {
  for (let i = 0; i < indices.length; i += CHUNK) {
    const end = Math.min(i + CHUNK, indices.length);
    for (let k = i; k < end; k++) {
      items[indices[k]].load(TEXT_PROPS);
    }
    await context.sync();
  }
}

/**
 * @param {readonly number[]} indicesToUse
 * @param {readonly (number | null)[]} levels
 * @param {import("office-js").Word.Paragraph[]} items
 */
function buildOutlineEntries(indicesToUse, levels, items) {
  /** @type {{ level: number, paragraph_index: number, text: string }[]} */
  const entries = [];
  for (const idx of indicesToUse) {
    const level = levels[idx];
    if (level === null) {
      continue;
    }
    entries.push({
      level,
      paragraph_index: idx,
      text: normalizeHeadingText(items[idx].text, 320),
    });
  }
  return entries;
}

/**
 * @param {{ entries: { level: number, paragraph_index: number, text: string }[], totalParagraphs: number, truncated: boolean, maxEntries?: number, headingCountTotal: number }} o
 */
function formatOutlineBody(o) {
  const { entries, totalParagraphs, truncated, maxEntries, headingCountTotal } = o;
  /** @type {string[]} */
  const lines = [
    `Document outline (${entries.length} heading(s); ${totalParagraphs} body paragraph(s))`,
  ];
  if (truncated) {
    lines.push(`Note: outline truncated to first ${maxEntries} heading(s) out of ${headingCountTotal}.`);
  }
  lines.push("");
  for (const e of entries) {
    const indent = "  ".repeat(Math.max(0, e.level - 1));
    lines.push(`${indent}- ${e.text}`);
  }
  if (entries.length === 0) {
    lines.push("(No outline headings found — paragraphs need Heading 1–9 styles or outline levels 1–9.)");
  }
  return lines.join("\n");
}

/**
 * @param {import("office-js").Word.RequestContext} context
 * @param {{ maxEntries?: number }} args
 */
async function collectDocumentOutline(context, args) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  const items = paragraphs.items;
  const totalParagraphs = items.length;

  await loadParagraphPropsInChunks(context, items, META_PROPS);
  const { headingIndices, levels } = scanHeadingParagraphs(items);

  const maxEntries = args.maxEntries;
  const truncated =
    typeof maxEntries === "number" &&
    maxEntries > 0 &&
    headingIndices.length > maxEntries;
  const indicesToUse = truncated ? headingIndices.slice(0, maxEntries) : headingIndices;

  await loadTextsForParagraphIndices(context, items, indicesToUse);

  const entries = buildOutlineEntries(indicesToUse, levels, items);
  const body = formatOutlineBody({
    entries,
    totalParagraphs,
    truncated,
    maxEntries,
    headingCountTotal: headingIndices.length,
  });

  return {
    content: [{ type: "text", text: body }],
    details: {
      total_body_paragraphs: totalParagraphs,
      heading_count_total: headingIndices.length,
      heading_count_returned: entries.length,
      truncated,
      max_entries: maxEntries ?? null,
      entries,
    },
  };
}

/** @returns {import("@mariozechner/pi-agent-core").AgentTool} */
export function wordGetDocumentOutlineTool() {
  return {
    name: "word_get_document_outline",
    label: "Read document outline",
    description:
      "Returns a table-of-contents-style outline of the active Word document: heading levels and titles in document order, based on built-in Heading 1–9 styles and/or paragraph outline levels 1–9. Use this to navigate long documents or plan edits without reading the full text.",
    parameters: Type.Object({
      max_entries: Type.Optional(
        Type.Integer({
          minimum: 1,
          description:
            "Optional cap on how many headings to return (first headings in document order). Omit for all headings.",
        }),
      ),
    }),
    execute: async (_id, params) => {
      if (typeof Word === "undefined") {
        throw new Error("Word API is not available outside Microsoft Word.");
      }
      let maxEntries = params?.max_entries;
      if (maxEntries !== undefined && maxEntries !== null) {
        if (typeof maxEntries !== "number" || !Number.isInteger(maxEntries) || maxEntries < 1) {
          throw new Error(
            `word_get_document_outline "max_entries" must be a positive integer, got ${JSON.stringify(maxEntries)}.`,
          );
        }
      } else {
        maxEntries = undefined;
      }
      return Word.run(async (context) =>
        collectDocumentOutline(context, { maxEntries }),
      );
    },
  };
}
