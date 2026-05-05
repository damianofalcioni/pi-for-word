/**
 * Pi4Word assistant tools for Pi Agent: Word document tools (`Word.run`), JavaScript REPL, and related helpers.
 * Import **`createWordTools`** for **`ChatPanel`** wiring; individual factories are exported for tests or custom sets.
 *
 * @module assistant-tools
 */
import { wordGetDocumentOutlineTool } from "./word-tool-get-document-outline.js";
import { wordGetSelectionTool } from "./word-tool-get-selection.js";
import { wordInsertMarkdownTool } from "./word-tool-insert-markdown.js";
import { wordSearchTextTool } from "./word-tool-search-text.js";

export {
  createPi4WordJavaScriptReplTool,
  pi4wordJavascriptReplParameters,
} from "./javascript-repl-tool.js";
export { wordGetDocumentOutlineTool } from "./word-tool-get-document-outline.js";
export { wordGetSelectionTool } from "./word-tool-get-selection.js";
export { wordInsertMarkdownTool } from "./word-tool-insert-markdown.js";
export { wordSearchTextTool } from "./word-tool-search-text.js";

/**
 * Tools that call the Word JavaScript API (requires host Word + `Word` global).
 * @returns {import("@mariozechner/pi-agent-core").AgentTool[]}
 */
export function createWordTools() {
  return [
    wordGetSelectionTool(),
    wordGetDocumentOutlineTool(),
    wordSearchTextTool(),
    wordInsertMarkdownTool(),
  ];
}
