/** @param {{ host?: unknown } | null | undefined} info */
export function computeWordHost(info) {
  return Boolean(
    info &&
      typeof Office !== "undefined" &&
      typeof Office.HostType !== "undefined" &&
      info.host === Office.HostType.Word,
  );
}

/**
 * Bootstraps Office.onReady (or falls back when Office.js is missing) and a timeout fallback.
 * @param {(info: { host?: unknown } | null | undefined) => void} startApp
 * @param {{ getStarted: () => boolean }} started
 */
export function scheduleOfficeBoot(startApp, started) {
  const runOfficeReady = () => {
    if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
      Office.onReady((officeInfo) => startApp(officeInfo));
    } else {
      startApp(null);
    }
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", runOfficeReady, { once: true });
  } else {
    runOfficeReady();
  }

  window.setTimeout(() => {
    if (!started.getStarted()) {
      console.warn(
        "[pi4word] Office.onReady did not finish; showing UI without host context (check Office.js / tracking prevention).",
      );
      startApp(null);
    }
  }, 2500);
}
