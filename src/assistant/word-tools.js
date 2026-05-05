import { Type } from "typebox";

/** @returns {import("@mariozechner/pi-agent-core").AgentTool} */
function wordGetSelectionTool() {
  return {
    name: "word_get_selection",
    label: "Read Word selection",
    description:
      "Returns the text currently selected in the active Word document. Use this before editing to see what the user highlighted.",
    parameters: Type.Object({}),
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
      "Inserts text into the document: after the current selection, replacing the selection, or at the end of the document.",
    parameters: Type.Object({
      text: Type.String({ description: "Text to insert." }),
      where: Type.Union(
        [
          Type.Literal("after_selection"),
          Type.Literal("replace_selection"),
          Type.Literal("end_of_document"),
        ],
        { description: "Where to put the text." },
      ),
    }),
    execute: async (_id, params) => {
      if (typeof Word === "undefined") {
        throw new Error("Word API is not available outside Microsoft Word.");
      }
      const { text, where } = params;
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
