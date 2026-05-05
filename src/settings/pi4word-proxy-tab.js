import { getAppStorage, ProxyTab } from "@mariozechner/pi-web-ui";
import { i18n } from "@mariozechner/mini-lit";
import { Input } from "@mariozechner/mini-lit/dist/Input.js";
import { Label } from "@mariozechner/mini-lit/dist/Label.js";
import { Switch } from "@mariozechner/mini-lit/dist/Switch.js";
import { html } from "lit";

/**
 * CORS proxy (from pi-web-ui {@link ProxyTab}) plus Pi Agent `streamProxy` (URL + Bearer token).
 */
export class Pi4WordProxySettingsTab extends ProxyTab {
  static properties = {
    streamProxyEnabled: { type: Boolean, state: true },
    streamProxyUrl: { type: String, state: true },
    streamProxyToken: { type: String, state: true },
  };

  constructor() {
    super();
    this.streamProxyEnabled = false;
    this.streamProxyUrl = "";
    this.streamProxyToken = "";
  }

  async connectedCallback() {
    await super.connectedCallback();
    try {
      const storage = getAppStorage();
      const e = await storage.settings.get("pi4word.streamProxy.enabled");
      const u = await storage.settings.get("pi4word.streamProxy.url");
      const t = await storage.settings.get("pi4word.streamProxy.token");
      if (e !== null && e !== undefined) {
        this.streamProxyEnabled = Boolean(e);
      }
      if (u !== null && u !== undefined) {
        this.streamProxyUrl = String(u);
      }
      if (t !== null && t !== undefined) {
        this.streamProxyToken = String(t);
      }
    } catch (err) {
      console.error("[pi4word] stream proxy settings load failed:", err);
    }
    this.requestUpdate();
  }

  async saveStreamProxySettings() {
    try {
      const storage = getAppStorage();
      await storage.settings.set("pi4word.streamProxy.enabled", this.streamProxyEnabled);
      await storage.settings.set("pi4word.streamProxy.url", this.streamProxyUrl);
      await storage.settings.set("pi4word.streamProxy.token", this.streamProxyToken);
    } catch (err) {
      console.error("[pi4word] stream proxy settings save failed:", err);
    }
  }

  /** @returns {import("lit").TemplateResult} */
  corsProxySectionTemplate() {
    return html`
      <section class="flex flex-col gap-4">
        <p class="text-sm text-muted-foreground">
          ${i18n(
            "Allows browser-based apps to bypass CORS restrictions when calling LLM providers. Required for Z-AI and Anthropic with OAuth token.",
          )}
        </p>

        <div class="flex items-center justify-between">
          <span class="text-sm font-medium text-foreground">${i18n("Use CORS Proxy")}</span>
          ${Switch({
            checked: this.proxyEnabled,
            onChange: (checked) => {
              this.proxyEnabled = checked;
              void this.saveProxySettings();
            },
          })}
        </div>

        <div class="space-y-2">
          ${Label({ children: i18n("Proxy URL") })}
          ${Input({
            type: "text",
            value: this.proxyUrl,
            disabled: !this.proxyEnabled,
            onInput: (e) => {
              this.proxyUrl = e.target.value;
            },
            onChange: () => this.saveProxySettings(),
          })}
          <p class="text-xs text-muted-foreground">
            ${i18n("Format: The proxy must accept requests as <proxy-url>/?url=<target-url>")}
          </p>
        </div>
      </section>
    `;
  }

  /** @returns {import("lit").TemplateResult} */
  streamProxySectionTemplate() {
    return html`
      <section class="flex flex-col gap-4 border-t border-border pt-6">
        <h3 class="text-sm font-semibold text-foreground">Pi stream proxy</h3>
        <p class="text-sm text-muted-foreground">
          Route LLM requests through a Pi server that implements
          <code class="text-xs">streamProxy</code> (POST <code class="text-xs">/api/stream</code> with Bearer token).
          Separate from the CORS proxy above.
        </p>

        <div class="flex items-center justify-between">
          <span class="text-sm font-medium text-foreground">Use stream proxy</span>
          ${Switch({
            checked: this.streamProxyEnabled,
            onChange: (checked) => {
              this.streamProxyEnabled = checked;
              void this.saveStreamProxySettings();
            },
          })}
        </div>

        <div class="space-y-2">
          ${Label({ children: "Proxy base URL" })}
          ${Input({
            type: "url",
            value: this.streamProxyUrl,
            disabled: !this.streamProxyEnabled,
            placeholder: "https://your-proxy.example.com",
            onInput: (e) => {
              this.streamProxyUrl = e.target.value;
            },
            onChange: () => this.saveStreamProxySettings(),
          })}
        </div>

        <div class="space-y-2">
          ${Label({ children: "Proxy Bearer token" })}
          ${Input({
            type: "password",
            value: this.streamProxyToken,
            disabled: !this.streamProxyEnabled,
            autocomplete: "off",
            onInput: (e) => {
              this.streamProxyToken = e.target.value;
            },
            onChange: () => this.saveStreamProxySettings(),
          })}
        </div>
      </section>
    `;
  }

  render() {
    return html`
      <div class="flex flex-col gap-8">
        ${this.corsProxySectionTemplate()}
        ${this.streamProxySectionTemplate()}
      </div>
    `;
  }
}

customElements.define("pi4word-proxy-settings-tab", Pi4WordProxySettingsTab);
