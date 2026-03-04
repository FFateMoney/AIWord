/**
 * storageService.js
 *
 * Thin wrapper around localStorage.
 * All keys are namespaced with a version prefix so future schema changes
 * can be handled cleanly (bump VERSION and clearAll old keys).
 *
 * Key names exposed via STORAGE_KEYS:
 *   provider, apikey, baseurl, model  — AI config
 *   ai_view, doc_name                 — last edited document
 */

const PREFIX = 'webaiword:v1:'

/** Well-known storage key names (without prefix). */
export const STORAGE_KEYS = {
  PROVIDER: 'provider',
  API_KEY:  'apikey',
  BASE_URL: 'baseurl',
  MODEL:    'model',
  AI_VIEW:  'ai_view',
  DOC_NAME: 'doc_name',
}

/**
 * Read a value from localStorage.
 * @param {string} key  - one of STORAGE_KEYS values
 * @returns {string|null}
 */
export function storageGet(key) {
  try {
    return localStorage.getItem(PREFIX + key)
  } catch {
    return null
  }
}

/**
 * Write a string value to localStorage.
 * Returns true on success, false if quota was exceeded (caller may degrade gracefully).
 * @param {string} key
 * @param {string} value
 * @returns {boolean}
 */
export function storageSet(key, value) {
  try {
    localStorage.setItem(PREFIX + key, value)
    return true
  } catch {
    return false
  }
}

/**
 * Remove a single key from localStorage.
 * @param {string} key
 */
export function storageRemove(key) {
  try {
    localStorage.removeItem(PREFIX + key)
  } catch { /* ignore */ }
}

/**
 * Remove ALL keys that belong to this app (prefix-matched).
 */
export function storageClearAll() {
  try {
    const toRemove = []
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i)
      if (k && k.startsWith(PREFIX)) toRemove.push(k)
    }
    toRemove.forEach(k => localStorage.removeItem(k))
  } catch { /* ignore */ }
}
