/**
 * Shared argument validation for Word assistant tools.
 */

/**
 * @param {string} toolName
 * @param {string} field
 * @param {unknown} value
 * @returns {string}
 */
export function requireToolString(toolName, field, value) {
  if (typeof value !== "string") {
    throw new Error(`${toolName} "${field}" must be a string, got ${typeof value}.`);
  }
  return value;
}

/**
 * @param {string} toolName
 * @param {string} field
 * @param {unknown} value
 * @param {number} defaultIndex
 * @returns {number}
 */
export function optionalNonNegativeIndex(toolName, field, value, defaultIndex) {
  if (value === undefined || value === null) {
    return defaultIndex;
  }
  if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
    throw new Error(
      `${toolName} "${field}" must be a non-negative integer, got ${JSON.stringify(value)}.`,
    );
  }
  return value;
}
