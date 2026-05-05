import { ProvidersModelsTab } from "@mariozechner/pi-web-ui";

/**
 * Same as pi-web-ui {@link ProvidersModelsTab}; shorter sidebar label for Pi4Word settings.
 */
export class Pi4WordProvidersTab extends ProvidersModelsTab {
  getTabName() {
    return "Providers";
  }
}

customElements.define("pi4word-providers-tab", Pi4WordProvidersTab);
