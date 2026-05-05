import { Type } from "typebox";

/** @type {Readonly<Record<string, "after_selection" | "replace_selection" | "end_of_document">>} */
const WORD_INSERT_WHERE_ALIASES = {
  after_selection: "after_selection",
  replace_selection: "replace_selection",
  end_of_document: "end_of_document",
  end: "end_of_document",
  document_end: "end_of_document",
  after: "after_selection",
  replace: "replace_selection",
};

/**
 * @param {unknown} where
 * @returns {"after_selection" | "replace_selection" | "end_of_document"}
 */
function normalizeWordInsertWhere(where) {
  if (typeof where !== "string") {
    throw new Error(
      `word_insert_text "where" must be a string (e.g. "end_of_document"), got ${typeof where}.`,
    );
  }
  const normalized = WORD_INSERT_WHERE_ALIASES[where] ?? where;
  if (
    normalized !== "after_selection" &&
    normalized !== "replace_selection" &&
    normalized !== "end_of_document"
  ) {
    throw new Error(
      `Invalid word_insert_text "where": ${JSON.stringify(where)}. Use exactly "after_selection", "replace_selection", or "end_of_document" (short aliases: after, replace, end).`,
    );
  }
  return normalized;
}

/**
 * @param {string} field
 * @param {unknown} value
 * @returns {string}
 */
function requireWordInsertTextString(field, value) {
  if (typeof value !== "string") {
    throw new Error(`word_insert_text "${field}" must be a string, got ${typeof value}.`);
  }
  return value;
}

/** @returns {import("@mariozechner/pi-agent-core").AgentTool} */
function wordGetSelectionTool() {
  return {
    name: "word_get_selection",
    label: "Read Word selection",
    description:
      "Returns the text currently selected in the active Word document. Use this before editing to see what the user highlighted.",
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
        await context.sync();
        const text = range.text ?? "";
        return {
          content: [
            {
              type: "text",
              text:
                text.length > 0
                  ? text
                  : "(no selection — empty string)",
            },
          ],
          details: { length: text.length },
        };
      });
    },
  };
}

/** @returns {import("@mariozechner/pi-agent-core").AgentTool} */
function wordInsertTextTool() {
  return {
    name: "word_insert_text",
    label: "Insert text into Word",
    description:
      'Inserts text in Word. Use where="after_selection" (after caret/selection), "replace_selection" (replace highlighted text), or "end_of_document" (append at document end). Shorthand where values end, after, replace, document_end are also accepted.',
    parameters: Type.Object({
      text: Type.String({ description: "Text to insert." }),
      where: Type.Union(
        [
          Type.Literal("after_selection"),
          Type.Literal("replace_selection"),
          Type.Literal("end_of_document"),
          Type.Literal("after"),
          Type.Literal("replace"),
          Type.Literal("end"),
          Type.Literal("document_end"),
        ],
        {
          description:
            'Placement: "after_selection" | "replace_selection" | "end_of_document" (aliases: after, replace, end, document_end).',
        },
      ),
    }),
    execute: async (_id, params) => {
      if (typeof Word === "undefined") {
        throw new Error("Word API is not available outside Microsoft Word.");
      }
      const text = requireWordInsertTextString("text", params.text);
      const where = normalizeWordInsertWhere(params.where);
      return Word.run(async (context) => {
        if (where === "end_of_document") {
          const body = context.document.body;
          body.insertParagraph(text, Word.InsertLocation.end);
        } else {
          const range = context.document.getSelection();
          if (where === "replace_selection") {
            range.insertText(text, Word.InsertLocation.replace);
          } else {
            range.insertText(text, Word.InsertLocation.after);
          }
        }
        await context.sync();
        return {
          content: [{ type: "text", text: "Inserted text successfully." }],
          details: { where },
        };
      });
    },
  };
}

/**
 * Tools that call the Word JavaScript API (requires host Word + `Word` global).
 * @returns {import("@mariozechner/pi-agent-core").AgentTool[]}
 */
export function createWordTools() {
  return [wordGetSelectionTool(), wordInsertTextTool()];
}
