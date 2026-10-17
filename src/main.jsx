import React from "react";
import ReactDOM from "react-dom/client";
import "./storageShim.js";
import "./index.css";
import App from "./App.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

// Register the service worker so the app is installable on phones (Add to
// Home Screen / Install App) and reopens even with a spotty connection.
//
// A previously-installed device can otherwise keep serving an old cached
// build indefinitely, silently running stale code against the local
// fallback instead of the live Supabase data. To prevent that: check for
// a newer service worker on every load, and once a new one is ready,
// reload automatically so the device is always running the latest build.
if ("serviceWorker" in navigator) {
  let hasReloaded = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (hasReloaded) return;
    hasReloaded = true;
    window.location.reload();
  });

  window.addEventListener("load", () => {
    navigator.serviceWorker.register("/sw.js").then((registration) => {
      registration.update().catch(() => {});
    }).catch((err) => {
      console.warn("Service worker registration failed:", err);
    });
  });
}
