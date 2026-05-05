import { Type } from "typebox";

/** Maps tool snake_case flags to Word `SearchOptions` (see Microsoft search-option guidance). */
const WORD_SEARCH_OPTION_FIELD_MAP = [
  { snake: "ignore_punct", camel: "ignorePunct" },
  { snake: "ignore_space", camel: "ignoreSpace" },
  { snake: "match_case", camel: "matchCase" },
  { snake: "match_prefix", camel: "matchPrefix" },
  { snake: "match_suffix", camel: "matchSuffix" },
  { snake: "match_whole_word", camel: "matchWholeWord" },
  { snake: "match_wildcards", camel: "matchWildcards" },
];

/** @type {ReadonlySet<string>} */
const WORD_SEARCH_OPTION_SNAKE_KEYS = new Set(
  WORD_SEARCH_OPTION_FIELD_MAP.map(({ snake }) => snake),
);

/**
 * @param {string} toolName
 * @param {string} paramLabel
 * @param {Record<string, unknown>} record
 */
function assertKnownSearchOptionKeys(toolName, paramLabel, record) {
  for (const key of Object.keys(record)) {
    if (!WORD_SEARCH_OPTION_SNAKE_KEYS.has(key)) {
      throw new Error(
        `${toolName} "${paramLabel}" has unknown key ${JSON.stringify(key)}. Allowed: ${[
          ...WORD_SEARCH_OPTION_SNAKE_KEYS,
        ].join(", ")}.`,
      );
    }
  }
}

/**
 * @param {string} toolName
 * @param {string} paramLabel
 * @param {Record<string, unknown>} record
 * @returns {import("office-js").Word.SearchOptions | undefined}
 */
function mapRecordToWordSearchOptions(toolName, paramLabel, record) {
  /** @type {Record<string, boolean>} */
  const out = {};
  for (const { snake, camel } of WORD_SEARCH_OPTION_FIELD_MAP) {
    if (!Object.prototype.hasOwnProperty.call(record, snake)) {
      continue;
    }
    const v = record[snake];
    if (v === undefined) {
      continue;
    }
    if (typeof v !== "boolean") {
      throw new Error(
        `${toolName} "${paramLabel}.${snake}" must be a boolean, got ${typeof v}.`,
      );
    }
    out[camel] = v;
  }
  return Object.keys(out).length > 0 ? /** @type {import("office-js").Word.SearchOptions} */ (out) : undefined;
}

/**
 * @param {string} toolName
 * @param {string} paramLabel e.g. "search_options"
 * @param {unknown} raw
 * @returns {import("office-js").Word.SearchOptions | undefined}
 */
export function compactWordSearchOptions(toolName, paramLabel, raw) {
  if (raw === undefined || raw === null) {
    return undefined;
  }
  if (typeof raw !== "object" || Array.isArray(raw)) {
    throw new Error(
      `${toolName} "${paramLabel}" must be an object with optional boolean flags, got ${typeof raw}.`,
    );
  }
  const record = /** @type {Record<string, unknown>} */ (raw);
  assertKnownSearchOptionKeys(toolName, paramLabel, record);
  return mapRecordToWordSearchOptions(toolName, paramLabel, record);
}

/**
 * @param {import("office-js").Word.SearchOptions | undefined} opts
 * @returns {Record<string, boolean> | undefined}
 */
export function wordSearchOptionsToSnakeEcho(opts) {
  if (!opts) {
    return undefined;
  }
  /** @type {Record<string, boolean>} */
  const echo = {};
  for (const { snake, camel } of WORD_SEARCH_OPTION_FIELD_MAP) {
    if (Object.prototype.hasOwnProperty.call(opts, camel)) {
      echo[snake] = /** @type {boolean} */ (opts[camel]);
    }
  }
  return Object.keys(echo).length > 0 ? echo : undefined;
}

/**
 * Runs a literal body search and loads matching ranges.
 * @param {import("office-js").Word.RequestContext} context
 * @param {string} searchText
 * @param {import("office-js").Word.SearchOptions | undefined} searchOptions
 * @returns {Promise<Word.Range[]>}
 */
export async function loadBodySearchItems(context, searchText, searchOptions) {
  const results =
    searchOptions !== undefined && Object.keys(searchOptions).length > 0
      ? context.document.body.search(searchText, searchOptions)
      : context.document.body.search(searchText);
  results.load("items");
  await context.sync();
  return results.items;
}

/**
 * @typedef {{ searchText: string, matchIndex: number, searchOptions: import("office-js").Word.SearchOptions | undefined }} WordBodySearchSpec
 */

/**
 * @param {string} toolName
 * @param {import("office-js").Word.RequestContext} context
 * @param {WordBodySearchSpec} spec
 * @returns {Promise<{ range: Word.Range, matchCount: number }>}
 */
export async function resolveBodySearchMatch(toolName, context, spec) {
  const { searchText, matchIndex, searchOptions } = spec;
  const items = await loadBodySearchItems(context, searchText, searchOptions);
  const matchCount = items.length;
  if (matchCount === 0) {
    throw new Error(
      `${toolName}: no matches for search text ${JSON.stringify(searchText)}.`,
    );
  }
  if (matchIndex >= matchCount) {
    throw new Error(
      `${toolName}: match index ${matchIndex} is out of range (${matchCount} match${matchCount === 1 ? "" : "es"} found; use 0..${matchCount - 1}).`,
    );
  }
  return { range: items[matchIndex], matchCount };
}

/**
 * @param {string} toolName
 * @param {import("office-js").Word.RequestContext} context
 * @param {WordBodySearchSpec} spec
 * @returns {Promise<Word.Range>}
 */
export async function rangeForBodySearchMatch(toolName, context, spec) {
  const { range } = await resolveBodySearchMatch(toolName, context, spec);
  return range;
}

export const WORD_BODY_SEARCH_OPTIONS_SCHEMA = Type.Object(
  {
    ignore_punct: Type.Optional(Type.Boolean()),
    ignore_space: Type.Optional(Type.Boolean()),
    match_case: Type.Optional(Type.Boolean()),
    match_prefix: Type.Optional(Type.Boolean()),
    match_suffix: Type.Optional(Type.Boolean()),
    match_whole_word: Type.Optional(Type.Boolean()),
    match_wildcards: Type.Optional(Type.Boolean()),
  },
  {
    description:
      "Optional Word SearchOptions (snake_case → ignorePunct, ignoreSpace, matchCase, …). Wildcards & behavior: https://learn.microsoft.com/en-us/office/dev/add-ins/word/search-option-guidance — When you chain word_insert_markdown after word_search_text with the same anchor text, word_insert_markdown’s anchor_search_options must use the same flags as word_search_text’s search_options (or omit both); mismatched flags can target a different occurrence.",
  },
);
