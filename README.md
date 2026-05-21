PS Toolbar — Svelte 5 scaffold

Purpose: lightweight, extensible toolbar component that reads a single JSON feed and renders notification icons with optional counters.

Quick start:
1. pnpm install
2. pnpm run dev

Notes:
- In development the component reads /notifications.json (included under public/).
- In production set VITE_NOTIF_FEED_URL to the live JSON endpoint before building.
- JSON shape (array): [{ id, icon, href, title, count }].

Integration with PowerSchool: build (npm run build) and integrate the built assets into the PowerSchool extension or load the bundle via your app bar hook.
