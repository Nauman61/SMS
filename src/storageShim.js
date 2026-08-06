// This app was originally built for Claude's in-chat "artifact" environment,
// which provides a built-in `window.storage` key-value API. That API does not
// exist on a normal website, so this file polyfills the same interface using
// the browser's localStorage — this lets App.jsx work completely unchanged.
//
// IMPORTANT LIMITATION: localStorage is per-browser, per-device. Data saved
// on one computer will NOT appear on another computer/phone, and will be
// lost if someone clears their browser's site data. This is fine for a
// single front-desk computer, but if you need multiple computers to see the
// same live data, you need a real backend/database instead of this shim.

const STORAGE_KEY = "hss_school_data_v1";

function readAll() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    console.error("Storage read error", e);
    return {};
  }
}

function writeAll(obj) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
  } catch (e) {
    console.error("Storage write error", e);
  }
}

window.storage = {
  async get(key) {
    const all = readAll();
    if (!(key in all)) return null;
    return { key, value: all[key], shared: false };
  },
  async set(key, value) {
    const all = readAll();
    all[key] = value;
    writeAll(all);
    return { key, value, shared: false };
  },
  async delete(key) {
    const all = readAll();
    if (!(key in all)) return null;
    delete all[key];
    writeAll(all);
    return { key, deleted: true, shared: false };
  },
  async list(prefix) {
    const all = readAll();
    const keys = Object.keys(all).filter((k) => !prefix || k.startsWith(prefix));
    return { keys, prefix, shared: false };
  },
};
