// This app was originally built for Claude's in-chat "artifact" environment,
// which provides a built-in `window.storage` key-value API. This file
// polyfills that same interface so App.jsx works completely unchanged —
// but it now has TWO modes:
//
//  1. CLOUD MODE (recommended): if VITE_SUPABASE_URL and
//     VITE_SUPABASE_ANON_KEY are set (see .env.example), every read/write
//     goes to a real shared Postgres database via Supabase. Every device
//     that opens the app — admin's phone, staff's phone, the office PC —
//     sees the SAME live data. This is what you want for real school use
//     across multiple people/devices.
//
//  2. LOCAL MODE (fallback): if those aren't set, it falls back to the
//     browser's localStorage like before — works standalone with zero
//     setup, but data stays on one device only.
//
// The app automatically uses whichever mode is configured. No changes to
// App.jsx are needed either way.

import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;
const CLOUD_ENABLED = Boolean(supabaseUrl && supabaseAnonKey);

let supabase = null;
if (CLOUD_ENABLED) {
  supabase = createClient(supabaseUrl, supabaseAnonKey);
} else {
  console.warn(
    "[HSS School] No Supabase credentials found (VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY). " +
    "Falling back to local browser storage — data will NOT sync across devices. " +
    "See README.md → 'Attach a shared database' to enable cloud sync."
  );
}

// Exposed so the app UI can show whether it's running in cloud or local mode.
window.__HSS_STORAGE_MODE__ = CLOUD_ENABLED ? "cloud" : "local";

// ---------- Local (localStorage) fallback ----------
const LOCAL_KEY = "hss_school_data_v1";
function readAllLocal() {
  try {
    const raw = localStorage.getItem(LOCAL_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    console.error("Local storage read error", e);
    return {};
  }
}
function writeAllLocal(obj) {
  try {
    localStorage.setItem(LOCAL_KEY, JSON.stringify(obj));
  } catch (e) {
    console.error("Local storage write error", e);
  }
}

const localStorageBackend = {
  async get(key) {
    const all = readAllLocal();
    if (!(key in all)) return null;
    return { key, value: all[key], shared: false };
  },
  async set(key, value) {
    const all = readAllLocal();
    all[key] = value;
    writeAllLocal(all);
    return { key, value, shared: false };
  },
  async delete(key) {
    const all = readAllLocal();
    if (!(key in all)) return null;
    delete all[key];
    writeAllLocal(all);
    return { key, deleted: true, shared: false };
  },
  async list(prefix) {
    const all = readAllLocal();
    const keys = Object.keys(all).filter((k) => !prefix || k.startsWith(prefix));
    return { keys, prefix, shared: false };
  },
};

// ---------- Cloud (Supabase) backend ----------
// Expects a table created via supabase/schema.sql:
//   kv_store(key text primary key, value text not null, updated_at timestamptz)
const cloudBackend = {
  async get(key) {
    const { data, error } = await supabase.from("kv_store").select("value").eq("key", key).maybeSingle();
    if (error) {
      console.error("Supabase get error", error);
      return null;
    }
    if (!data) return null;
    return { key, value: data.value, shared: true };
  },
  async set(key, value) {
    const { error } = await supabase.from("kv_store").upsert({ key, value, updated_at: new Date().toISOString() });
    if (error) {
      console.error("Supabase set error", error);
      return null;
    }
    return { key, value, shared: true };
  },
  async delete(key) {
    const { error } = await supabase.from("kv_store").delete().eq("key", key);
    if (error) {
      console.error("Supabase delete error", error);
      return null;
    }
    return { key, deleted: true, shared: true };
  },
  async list(prefix) {
    let query = supabase.from("kv_store").select("key");
    if (prefix) query = query.like("key", `${prefix}%`);
    const { data, error } = await query;
    if (error) {
      console.error("Supabase list error", error);
      return { keys: [], prefix, shared: true };
    }
    return { keys: data.map((row) => row.key), prefix, shared: true };
  },
};

window.storage = CLOUD_ENABLED ? cloudBackend : localStorageBackend;
