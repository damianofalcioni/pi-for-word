/**
 * Task pane composition root: Office bootstrap, IndexedDB (via settings), ChatPanel, Word agent,
 * session autosave, Settings/Sessions wiring.
 *
 * Side effects: `initializeTaskPane()` schedules Office host boot, storage init/migration, DOM shell,
 * pi-web-ui dialogs, and agent subscribe for session persistence.
 *
 * @module task-pane
 */
export { initializeTaskPane } from "./task-pane.init.js";
