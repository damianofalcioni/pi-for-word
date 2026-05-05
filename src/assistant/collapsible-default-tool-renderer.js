import { html } from "lit";
import { createRef, ref } from "lit/directives/ref.js";
import { Code } from "lucide";
import {
  i18n,
  registerToolRenderer,
  renderCollapsibleHeader,
  renderHeader,
} from "@mariozechner/pi-web-ui";
import { createWordTools } from "./word-tools.js";

/**
 * @param {unknown} params
 * @returns {string}
 */
function paramsToJsonString(params) {
  if (!params) {
    return "";
  }
  try {
    return JSON.stringify(JSON.parse(params), null, 2);
  } catch {
    try {
      return JSON.stringify(params, null, 2);
    } catch {
      return String(params);
    }
  }
}

/**
 * @param {unknown} result
 * @returns {{ text: string, language: "json" | "text" }}
 */
function resultOutputToCodeBlock(result) {
  let outputJson =
    result?.content
      ?.filter((c) => c.type === "text")
      .map((c) => c.text)
      .join("\n") || i18n("(no output)");
  let outputLanguage = "text";
  try {
    const parsed = JSON.parse(outputJson);
    outputJson = JSON.stringify(parsed, null, 2);
    outputLanguage = "json";
  } catch {
    // keep as text
  }
  return { text: outputJson, language: outputLanguage };
}

/**
 * @param {unknown} result
 * @param {boolean} isStreaming
 * @returns {"error" | "complete" | "inprogress"}
 */
function toolRenderState(result, isStreaming) {
  if (result) {
    return result.isError ? "error" : "complete";
  }
  return isStreaming ? "inprogress" : "complete";
}

/**
 * @param {{ toolName: string, bodyRef: ReturnType<typeof createRef>, chevronRef: ReturnType<typeof createRef>, state: "error" | "complete" | "inprogress", paramsJson: string, outputJson: string, outputLanguage: "json" | "text" }} o
 */
function templateResultBlock(o) {
  return html`
    <div>
      ${renderCollapsibleHeader(o.state, Code, o.toolName, o.bodyRef, o.chevronRef, false)}
      <div ${ref(o.bodyRef)} class="max-h-0 overflow-hidden transition-all duration-300 space-y-3">
        ${o.paramsJson
          ? html`<div>
              <div class="text-xs font-medium mb-1 text-muted-foreground">${i18n("Input")}</div>
              <code-block .code=${o.paramsJson} language="json"></code-block>
            </div>`
          : ""}
        <div>
          <div class="text-xs font-medium mb-1 text-muted-foreground">${i18n("Output")}</div>
          <code-block .code=${o.outputJson} language="${o.outputLanguage}"></code-block>
        </div>
      </div>
    </div>
  `;
}

/**
 * @param {{ toolName: string, bodyRef: ReturnType<typeof createRef>, chevronRef: ReturnType<typeof createRef>, state: "error" | "complete" | "inprogress", paramsJson: string }} o
 */
function templateParamsBlock(o) {
  return html`
    <div>
      ${renderCollapsibleHeader(o.state, Code, o.toolName, o.bodyRef, o.chevronRef, false)}
      <div ${ref(o.bodyRef)} class="max-h-0 overflow-hidden transition-all duration-300 space-y-3">
        <div>
          <div class="text-xs font-medium mb-1 text-muted-foreground">${i18n("Input")}</div>
          <code-block .code=${o.paramsJson} language="json"></code-block>
        </div>
      </div>
    </div>
  `;
}

/**
 * @param {string} toolName
 * @param {unknown} result
 * @param {unknown} params
 * @param {boolean} isStreaming
 */
function renderToolCallContent(toolName, result, params, isStreaming) {
  const state = toolRenderState(result, isStreaming);
  const paramsJson = paramsToJsonString(params);
  const bodyRef = createRef();
  const chevronRef = createRef();

  if (result) {
    const { text: outputJson, language: outputLanguage } = resultOutputToCodeBlock(result);
    return {
      content: templateResultBlock({
        toolName,
        bodyRef,
        chevronRef,
        state,
        paramsJson,
        outputJson,
        outputLanguage,
      }),
      isCustom: false,
    };
  }

  if (params) {
    const emptyStreamingParams =
      isStreaming && (!paramsJson || paramsJson === "{}" || paramsJson === "null");
    if (emptyStreamingParams) {
      return {
        content: html`
          <div>
            ${renderHeader(
              state,
              Code,
              `${toolName} — ${i18n("Preparing tool parameters...")}`,
            )}
          </div>
        `,
        isCustom: false,
      };
    }
    return {
      content: templateParamsBlock({
        toolName,
        bodyRef,
        chevronRef,
        state,
        paramsJson,
      }),
      isCustom: false,
    };
  }

  return {
    content: html`
      <div>
        ${renderHeader(state, Code, `${toolName} — ${i18n("Preparing tool...")}`)}
      </div>
    `,
    isCustom: false,
  };
}

/**
 * @param {string} toolName
 */
function createCollapsibleDefaultRenderer(toolName) {
  return {
    /** @param {unknown} params @param {unknown} result @param {boolean} isStreaming */
    render(params, result, isStreaming) {
      return renderToolCallContent(toolName, result, params, isStreaming);
    },
  };
}

let registered = false;

/** Registers collapsible JSON renderers for all Word document tools (same names as `createWordTools`). */
export function registerCollapsibleWordToolRenderers() {
  if (registered) {
    return;
  }
  registered = true;
  for (const tool of createWordTools()) {
    registerToolRenderer(tool.name, createCollapsibleDefaultRenderer(tool.name));
  }
}
