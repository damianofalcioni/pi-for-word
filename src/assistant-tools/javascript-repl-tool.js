import { createJavaScriptReplTool } from "@mariozechner/pi-web-ui";
import { Type } from "typebox";

const jsReplTitle = Type.String({
  description:
    "Brief title describing what the code snippet tries to achieve in active form, e.g. 'Calculating sum'",
});

/**
 * Same contract as pi-web-ui `javascript_repl`, but accepts `script` as an alias for `code`.
 * Must stay a single JSON `"type": "object"` schema — APIs reject root-level `anyOf`/`oneOf` from `Type.Union`.
 */
export const pi4wordJavascriptReplParameters = Type.Object(
  {
    title: jsReplTitle,
    code: Type.Optional(
      Type.String({
        description: "JavaScript code to execute. Provide this or script (at least one non-empty).",
      }),
    ),
    script: Type.Optional(
      Type.String({
        description: "Same as code. Prefer code when possible.",
      }),
    ),
  },
  {
    description:
      "title (required) plus code and/or script — at least one of code or script must be a non-empty string.",
  },
);

function coerceJavascriptReplTitle(o) {
  const title = o.title;
  if (typeof title !== "string" || title.trim().length === 0) {
    throw new Error('javascript_repl requires a non-empty string "title".');
  }
  return title;
}

/**
 * @param {Record<string, unknown>} o
 * @returns {string}
 */
function pickJavascriptReplCode(o) {
  const codeRaw = o.code;
  if (typeof codeRaw === "string" && codeRaw.length > 0) return codeRaw;
  const scriptRaw = o.script;
  if (typeof scriptRaw === "string" && scriptRaw.length > 0) return scriptRaw;
  return "";
}

/**
 * @param {unknown} args
 * @returns {{ title: string; code: string }}
 */
function normalizeJavascriptReplArgs(args) {
  if (typeof args !== "object" || args === null) {
    throw new Error(`javascript_repl arguments must be an object, got ${typeof args}.`);
  }
  const o = /** @type {Record<string, unknown>} */ (args);
  const title = coerceJavascriptReplTitle(o);
  const code = pickJavascriptReplCode(o);
  if (!code) {
    throw new Error('javascript_repl requires non-empty "code" or "script".');
  }
  return { title, code };
}

/**
 * JavaScript REPL tool with relaxed argument shapes for common model mistakes.
 * @returns {ReturnType<typeof createJavaScriptReplTool>}
 */
export function createPi4WordJavaScriptReplTool() {
  const inner = createJavaScriptReplTool();
  return {
    ...inner,
    parameters: pi4wordJavascriptReplParameters,
    execute: /** @type {typeof inner.execute} */ (function (toolCallId, args, signal) {
      const { title, code } = normalizeJavascriptReplArgs(args);
      return inner.execute.call(this, toolCallId, { title, code }, signal);
    }),
  };
}
