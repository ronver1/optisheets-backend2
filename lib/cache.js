/**
 * Simple in-memory cache helper.
 * Note: on Vercel serverless each function instance has its own memory,
 * so this is best used for per-invocation memoization or short-lived caching.
 * For cross-request caching, store data in Supabase or an external KV store.
 */

const store = new Map();

/**
 * @param {string} key
 * @param {any} value
 * @param {number} ttlMs - time-to-live in milliseconds (default: 60s)
 */
function set(key, value, ttlMs = 60_000) {
  const expiresAt = Date.now() + ttlMs;
  store.set(key, { value, expiresAt });
}

/**
 * @param {string} key
 * @returns {any|null} cached value or null if missing/expired
 */
function get(key) {
  const entry = store.get(key);
  if (!entry) return null;
  if (Date.now() > entry.expiresAt) {
    store.delete(key);
    return null;
  }
  return entry.value;
}

/**
 * @param {string} key
 * @param {() => Promise<any>} fetchFn - async function to populate cache on miss
 * @param {number} ttlMs
 */
async function getOrFetch(key, fetchFn, ttlMs = 60_000) {
  const cached = get(key);
  if (cached !== null) return cached;
  const value = await fetchFn();
  set(key, value, ttlMs);
  return value;
}

function del(key) {
  store.delete(key);
}

function flush() {
  store.clear();
}

module.exports = { set, get, getOrFetch, del, flush };
