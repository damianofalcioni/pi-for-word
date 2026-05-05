import DOMPurify from "dompurify";
import { marked } from "marked";
import { Type } from "typebox";
import {
  WORD_BODY_SEARCH_OPTIONS_SCHEMA,
  compactWordSearchOptions,
  rangeForBodySearchMatch,
  wordSearchOptionsToSnakeEcho,
} from "./word-tools-search.js";
import { optionalNonNegativeIndex, requireToolString } from "./word-tools-shared.js";

/** Allowlists HTML safe for Word insertion; extends defaults so GFM task lists keep `<input type="checkbox">`. */
const MARKDOWN_HTML_SANITIZE = {
  USE_PROFILES: { html: true },
  ADD_TAGS: ["input"],
  ADD_ATTR: ["checked", "disabled", "type"],
};

/** @type {Readonly<Record<string, "after_selection" | "before_selection" | "replace_selection" | "end_of_document">>} */
const WORD_INSERT_WHERE_ALIASES = {
  after_selection: "after_selection",
  before_selection: "before_selection",
  replace_selection: "replace_selection",
  end_of_document: "end_of_document",
  end: "end_of_document",
  document_end: "end_of_document",
  after: "after_selection",
  before: "before_selection",
  replace: "replace_selection",
};

/**
 * @param {unknown} where
 * @param {string} toolName
 * @returns {"after_selection" | "before_selection" | "replace_selection" | "end_of_document"}
 */
function normalizeWordInsertWhere(where, toolName) {
  if (typeof where !== "string") {
    throw new Error(
      `${toolName} "where" must be a string (e.g. "end_of_document"), got ${typeof where}.`,
    );
  }
  const normalized = WORD_INSERT_WHERE_ALIASES[where] ?? where;
  if (
    normalized !== "after_selection" &&
    normalized !== "before_selection" &&
    normalized !== "replace_selection" &&
    normalized !== "end_of_document"
  ) {
    throw new Error(
      `Invalid ${toolName} "where": ${JSON.stringify(where)}. Use exactly "after_selection", "before_selection", "replace_selection", or "end_of_document" (short aliases: after, before, replace, end).`,
    );
  }
  return normalized;
}

/** Placeholder for literal backslash while unescaping (private-use, unlikely in user text). */
const _UNESC_BACKSLASH_PH = "\uE000";

/**
 * Tool arguments sometimes keep JSON-style escapes as two-character sequences (e.g. backslash + "n")
 * instead of a real newline. Converts those to real characters; use `\\` for a literal backslash.
 * @param {string} s
 * @returns {string}
 */
function unescapeToolStringEscapes(s) {
  return s
    .replace(/\\\\/g, _UNESC_BACKSLASH_PH)
    .replace(/\\n/g, "\n")
    .replace(/\\r/g, "\r")
    .replace(/\\t/g, "\t")
    .replace(new RegExp(_UNESC_BACKSLASH_PH, "g"), "\\");
}

/**
 * @param {string} markdown
 * @returns {string}
 */
function markdownToWordHtml(markdown) {
  const trimmed = markdown.trim();
  if (!trimmed) {
    return "";
  }
  const html = marked(trimmed, { async: false });
  const raw = typeof html === "string" ? html.trim() : "";
  if (!raw) {
    return "";
  }
  return DOMPurify.sanitize(raw, MARKDOWN_HTML_SANITIZE);
}

/**
 * @param {import("office-js").Word.RequestContext} context
 * @param {"after_selection" | "before_selection" | "replace_selection" | "end_of_document"} where
 * @param {string} html
 * @param {{
 *   anchorSearchText: string,
 *   anchorMatchIndex: number,
 *   anchorSearchOptions: import("office-js").Word.SearchOptions | undefined,
 * }} anchor
 */
async function insertMarkdownHtmlAtPlacement(context, where, html, anchor) {
  const { anchorSearchText, anchorMatchIndex, anchorSearchOptions } = anchor;
  if (where === "end_of_document") {
    context.document.body.insertHtml(html, Word.InsertLocation.end);
    return;
  }
  const range =
    anchorSearchText.length > 0
      ? await rangeForBodySearchMatch("word_insert_markdown", context, {
          searchText: anchorSearchText,
          matchIndex: anchorMatchIndex,
          searchOptions: anchorSearchOptions,
        })
      : context.document.getSelection();
  if (where === "replace_selection") {
    range.insertHtml(html, Word.InsertLocation.replace);
  } else if (where === "before_selection") {
    range.insertHtml(html, Word.InsertLocation.before);
  } else {
    range.insertHtml(html, Word.InsertLocation.after);
  }
}

const WORD_INSERT_WHERE_UNION = Type.Union(
  [
    Type.Literal("after_selection"),
    Type.Literal("before_selection"),
    Type.Literal("replace_selection"),
    Type.Literal("end_of_document"),
    Type.Literal("after"),
    Type.Literal("before"),
    Type.Literal("replace"),
    Type.Literal("end"),
    Type.Literal("document_end"),
  ],
  {
    description:
      'Placement: "after_selection" | "before_selection" | "replace_selection" | "end_of_document" (aliases: after, before, replace, end, document_end).',
  },
);

const WORD_INSERT_MARKDOWN_PARAMETERS = Type.Object({
  markdown: Type.String({
    description:
      "Markdown to render and insert. Plain text counts as Markdown; line breaks via real newlines or escapes \\n, \\r, \\t; use \\\\ for a literal backslash.",
  }),
  where: WORD_INSERT_WHERE_UNION,
  anchor_search_text: Type.Optional(
    Type.String({
      description:
        "If non-empty after trim: target this text in the document body (see word_search_text) instead of the current selection. Re-searches on each call. Use the same anchor_search_options as word_search_text’s search_options when both are used.",
    }),
  ),
  anchor_match_index: Type.Optional(
    Type.Integer({
      minimum: 0,
      description:
        "0-based occurrence index when using anchor_search_text (default 0). Ignored without anchor_search_text.",
    }),
  ),
  anchor_search_options: Type.Optional(WORD_BODY_SEARCH_OPTIONS_SCHEMA),
});

/** @returns {import("@mariozechner/pi-agent-core").AgentTool} */
export function wordInsertMarkdownTool() {
  return {
    name: "word_insert_markdown",
    label: "Insert into Word",
    description:
      'Use this for every insert/replace into Word. GitHub-flavored Markdown is converted to HTML, sanitized with DOMPurify, then inserted (plain paragraphs work fine; headings, lists, bold/italic, links, tables, fenced code when needed). Placement: where="after_selection" | "before_selection" | "replace_selection" | "end_of_document" (aliases: after, before, replace, end, document_end). To anchor by a search hit instead of the user selection, set anchor_search_text (same string as word_search_text), optional anchor_match_index, and optional anchor_search_options (Word SearchOptions — https://learn.microsoft.com/en-us/office/dev/add-ins/word/search-option-guidance). Important: when you used word_search_text first, anchor_search_options must mirror that call’s search_options (same flags; copy from word_search_text result details.anchor_search_options if helpful). Omitting anchor_search_options must match omitting search_options on the search—otherwise find and insert can disagree. anchor_* is ignored when where is end_of_document.',
    parameters: WORD_INSERT_MARKDOWN_PARAMETERS,
    execute: async (_id, params) => {
      if (typeof Word === "undefined") {
        throw new Error("Word API is not available outside Microsoft Word.");
      }
      const markdown = unescapeToolStringEscapes(
        requireToolString("word_insert_markdown", "markdown", params.markdown),
      );
      const where = normalizeWordInsertWhere(params.where, "word_insert_markdown");
      const html = markdownToWordHtml(markdown);
      if (!html) {
        return {
          content: [{ type: "text", text: "Nothing inserted (empty markdown)." }],
          details: { where, skipped: true },
        };
      }
      const anchorSearchText =
        typeof params.anchor_search_text === "string" ? params.anchor_search_text.trim() : "";
      const anchorMatchIndex = optionalNonNegativeIndex(
        "word_insert_markdown",
        "anchor_match_index",
        params.anchor_match_index,
        0,
      );
      const anchorSearchOptions = compactWordSearchOptions(
        "word_insert_markdown",
        "anchor_search_options",
        params.anchor_search_options,
      );
      return Word.run(async (context) => {
        await insertMarkdownHtmlAtPlacement(context, where, html, {
          anchorSearchText,
          anchorMatchIndex,
          anchorSearchOptions:
            anchorSearchText.length > 0 ? anchorSearchOptions : undefined,
        });
        await context.sync();
        return {
          content: [{ type: "text", text: "Inserted into Word successfully." }],
          details: {
            where,
            used_search_anchor: where !== "end_of_document" && anchorSearchText.length > 0,
            anchor_search_options_used:
              where !== "end_of_document" && anchorSearchText.length > 0
                ? wordSearchOptionsToSnakeEcho(anchorSearchOptions)
                : undefined,
          },
        };
      });
    },
  };
}
