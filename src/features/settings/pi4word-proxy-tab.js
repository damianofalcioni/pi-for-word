import { getAppStorage, SettingsTab } from "@mariozechner/pi-web-ui";
import { Input } from "@mariozechner/mini-lit/dist/Input.js";
import { Label } from "@mariozechner/mini-lit/dist/Label.js";
import { Switch } from "@mariozechner/mini-lit/dist/Switch.js";
import { html } from "lit";

/**
 * Settings tab: Pi Agent `streamProxy` (server URL + Bearer token), separate from CORS proxy in ProxyTab.
 */
export class Pi4WordProxyTab extends SettingsTab {
  static properties = {
    enabled: { type: Boolean, state: true },
    url: { type: String, state: true },
    token: { type: String, state: true },
  };

  constructor() {
    super();
    this.enabled = false;
    this.url = "";
    this.token = "";
  }

  async connectedCallback() {
    super.connectedCallback();
    try {
      const storage = getAppStorage();
      const e = await storage.settings.get("pi4word.streamProxy.enabled");
      const u = await storage.settings.get("pi4word.streamProxy.url");
      const t = await storage.settings.get("pi4word.streamProxy.token");
      if (e !== null && e !== undefined) {
        this.enabled = Boolean(e);
      }
      if (u !== null && u !== undefined) {
        this.url = String(u);
      }
      if (t !== null && t !== undefined) {
        this.token = String(t);
      }
    } catch (err) {
      console.error("[pi4word] Pi4WordProxyTab load failed:", err);
    }
    this.requestUpdate();
  }

  async save() {
    try {
      const storage = getAppStorage();
      await storage.settings.set("pi4word.streamProxy.enabled", this.enabled);
      await storage.settings.set("pi4word.streamProxy.url", this.url);
      await storage.settings.set("pi4word.streamProxy.token", this.token);
    } catch (err) {
      console.error("[pi4word] Pi4WordProxyTab save failed:", err);
    }
  }

  getTabName() {
    return "Pi4Word proxy";
  }

  render() {
    return html`
      <div class="flex flex-col gap-4">
        <p class="text-sm text-muted-foreground">
          Route LLM requests through a Pi server that implements
          <code class="text-xs">streamProxy</code> (POST <code class="text-xs">/api/stream</code> with Bearer
          token). This is separate from the CORS proxy in the Proxy tab.
        </p>

        <div class="flex items-center justify-between">
          <span class="text-sm font-medium text-foreground">Use stream proxy</span>
          ${Switch({
            checked: this.enabled,
            onChange: (checked) => {
              this.enabled = checked;
              void this.save();
            },
          })}
        </div>

        <div class="space-y-2">
          ${Label({ children: "Proxy base URL" })}
          ${Input({
            type: "url",
            value: this.url,
            disabled: !this.enabled,
            placeholder: "https://your-proxy.example.com",
            onInput: (e) => {
              this.url = e.target.value;
            },
            onChange: () => this.save(),
          })}
        </div>

        <div class="space-y-2">
          ${Label({ children: "Proxy Bearer token" })}
          ${Input({
            type: "password",
            value: this.token,
            disabled: !this.enabled,
            autocomplete: "off",
            onInput: (e) => {
              this.token = e.target.value;
            },
            onChange: () => this.save(),
          })}
        </div>
      </div>
    `;
  }
}

customElements.define("pi4word-proxy-tab", Pi4WordProxyTab);
