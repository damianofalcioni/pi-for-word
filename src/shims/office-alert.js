/**
 * Office.js task pane hosts block native `window.alert`.
 * pi-web-ui uses `alert()` for attachment and paste errors.
 */
window.alert = function officeAlertShim(message) {
  console.warn("[alert]", String(message));
};
