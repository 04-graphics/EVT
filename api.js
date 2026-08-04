/*******************************************************
 * EventPass Lite — api.js
 * Single helper that every page uses to talk to the
 * Apps Script JSON API. Replaces all google.script.run calls.
 *******************************************************/

// 1) Paste your deployed Apps Script Web App URL here (ends in /exec).
//    You can override it per-link with a "?api=..." query param (used
//    e.g. when opening scanner.html from staff.html) — the override is
//    remembered in localStorage so you don't have to pass it on every page.
const DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycbzAnoslatPJvcAo4ooKYI5MlCMwnYjru3ZpnE3PZ3Es7VlJeugFhccFhlEMImD4n3wH/exec';

(function () {
  const qs = new URLSearchParams(location.search);
  const fromQuery = qs.get('api');
  if (fromQuery) localStorage.setItem('ep_api_url', fromQuery);
})();

function getApiUrl() {
  return localStorage.getItem('ep_api_url') || DEFAULT_API_URL;
}

/**
 * Calls the Apps Script API.
 * @param {string} action - action name (e.g. "getGuests")
 * @param {object} data - extra payload fields
 * @returns {Promise<object>} parsed JSON response
 *
 * IMPORTANT: Content-Type is deliberately "text/plain" (not
 * "application/json"). Apps Script can't handle a CORS
 * preflight (OPTIONS) request, and application/json would
 * trigger one. text/plain keeps this a "simple request" that
 * skips preflight entirely, while the body is still valid JSON
 * that doPost() parses on the other end.
 */
async function api(action, data = {}) {
  const API_URL = getApiUrl();

  if (!API_URL || API_URL === 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE') {
    return { success: false, error: 'Apps Script API URL is not configured. Set DEFAULT_API_URL in js/api.js.' };
  }

  try {
    const res = await fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, ...data })
    });

    if (!res.ok) {
      return { success: false, error: 'Server returned HTTP ' + res.status };
    }

    return await res.json();
  } catch (err) {
    console.error('API error:', err);
    return { success: false, error: 'Could not reach the backend. Check your connection.' };
  }
}

/** Convenience wrapper for reads — same call, just unwraps { success, data }. */
async function apiGet(action, data = {}) {
  const res = await api(action, data);
  return res && res.success ? res.data : null;
}
