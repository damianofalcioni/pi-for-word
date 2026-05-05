import TurndownService from "turndown";
import { gfm } from "turndown-plugin-gfm";
import { Type } from "typebox";

/** @type {import("turndown").default | undefined} */
let wordSelectionTurndown;

/**
 * Converts HTML from Word’s `Range.getHtml()` to GitHub-flavored Markdown (aligned with `word_insert_markdown`).
 * @param {string} html
 * @returns {string}
 */
function wordHtmlToMarkdown(html) {
  const raw = typeof html === "string" ? html.trim() : "";
  if (!raw) {
    return "";
  }
  if (!wordSelectionTurndown) {
    wordSelectionTurndown = new TurndownService({
      headingStyle: "atx",
      bulletListMarker: "-",
      codeBlockStyle: "fenced",
      emDelimiter: "*",
    });
    wordSelectionTurndown.use(gfm);
  }
  return wordSelectionTurndown.turndown(raw).trim();
}

/** @returns {import("@mariozechner/pi-agent-core").AgentTool} */
export function wordGetSelectionTool() {
  return {
    name: "word_get_selection",
    label: "Read Word selection",
    description:
      "Returns the current selection in the active Word document as GitHub-flavored Markdown (headings, lists, emphasis, links, tables, etc., reconstructed from Word's HTML). Use this before editing to see what the user highlighted.",
    parameters: Type.Object(
      {},
      {
        description:
          "No fields — pass an empty object {} only. Do not pass document id, range, or other keys.",
      },
    ),
    execute: async () => {
      if (typeof Word === "undefined") {
        throw new Error("Word API is not available outside Microsoft Word.");
      }
      return Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("text");
        const htmlResult = range.getHtml();
        await context.sync();
        const text = range.text ?? "";
        const html = htmlResult.value ?? "";
        let markdown = wordHtmlToMarkdown(html);
        if (!markdown && text.trim()) {
          markdown = text;
        }
        return {
          content: [
            {
              type: "text",
              text:
                markdown.length > 0
                  ? markdown
                  : "(no selection — empty string)",
            },
          ],
          details: { length: markdown.length },
        };
      });
    },
  };
}
