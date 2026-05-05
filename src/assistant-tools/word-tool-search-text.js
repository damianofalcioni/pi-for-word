import { Type } from "typebox";
import {
  WORD_BODY_SEARCH_OPTIONS_SCHEMA,
  compactWordSearchOptions,
  resolveBodySearchMatch,
  wordSearchOptionsToSnakeEcho,
} from "./word-tools-search.js";
import { optionalNonNegativeIndex, requireToolString } from "./word-tools-shared.js";

/**
 * @param {import("office-js").Word.RequestContext} context
 * @param {{ text: string, matchIndex: number, searchOptions: import("office-js").Word.SearchOptions | undefined }} args
 */
async function wordSearchTextSyncResult(context, args) {
  const { text, matchIndex, searchOptions } = args;
  const { range, matchCount } = await resolveBodySearchMatch("word_search_text", context, {
    searchText: text,
    matchIndex,
    searchOptions,
  });
  range.load("text");
  await context.sync();
  const snippet = range.text ?? "";
  const preview =
    snippet.length > 240 ? `${snippet.slice(0, 237).trimEnd()}...` : snippet;
  const summary = `Found ${matchCount} match(es). Showing match index ${matchIndex} of 0..${Math.max(0, matchCount - 1)}.\nPreview: ${preview.length > 0 ? preview : "(empty range)"}`;
  return {
    content: [{ type: "text", text: summary }],
    details: {
      match_count: matchCount,
      match_index: matchIndex,
      preview,
      anchor_search_text: text,
      anchor_match_index: matchIndex,
      anchor_search_options: wordSearchOptionsToSnakeEcho(searchOptions),
    },
  };
}

/** @returns {import("@mariozechner/pi-agent-core").AgentTool} */
export function wordSearchTextTool() {
  return {
    name: "word_search_text",
    label: "Search in Word document",
    description:
      "Find text in the active document body using Word.search with optional SearchOptions (ignore_punct, match_wildcards, etc.). Guidance: https://learn.microsoft.com/en-us/office/dev/add-ins/word/search-option-guidance — Returns match count and a preview of the chosen hit. Before inserting without a selection, pass the same text, match_index, and search_options into word_insert_markdown as anchor_search_text, anchor_match_index, and anchor_search_options. Important: anchor_search_options on insert must use the same Word SearchOptions boolean flags as search_options on this call (you can reuse the tool result details.anchor_search_options); mismatched flags can resolve a different occurrence than the one previewed here.",
    parameters: Type.Object({
      text: Type.String({
        description:
          "Search string (literal or wildcard pattern when match_wildcards is true). Special markers such as ^p, ^t per Word search docs.",
      }),
      match_index: Type.Optional(
        Type.Integer({
          minimum: 0,
          description: "Which match to describe, 0-based. Default 0 (first occurrence).",
        }),
      ),
      search_options: Type.Optional(WORD_BODY_SEARCH_OPTIONS_SCHEMA),
    }),
    execute: async (_id, params) => {
      if (typeof Word === "undefined") {
        throw new Error("Word API is not available outside Microsoft Word.");
      }
      const text = requireToolString("word_search_text", "text", params.text);
      const matchIndex = optionalNonNegativeIndex(
        "word_search_text",
        "match_index",
        params.match_index,
        0,
      );
      const searchOptions = compactWordSearchOptions(
        "word_search_text",
        "search_options",
        params.search_options,
      );
      return Word.run(async (context) =>
        wordSearchTextSyncResult(context, { text, matchIndex, searchOptions }),
      );
    },
  };
}
