# Copilot Instructions for ps-toolbar

## Build, Test, and Lint Commands

- **Install dependencies:**
  - `pnpm install`
- **Start development server:**
  - `pnpm run dev`
- **Build for production:**
  - `pnpm run build`
- **Run all tests:**
  - `pnpm run test`
- **Run a single test file:**
  - `pnpm exec vitest run path/to/testfile.ts`
- **Watch tests:**
  - `pnpm run test:watch`

## High-Level Architecture

- **Svelte 5 custom element**: The main toolbar is a Svelte custom element (`ps-toolbar`) rendered via `App.svelte` and `Toolbar.svelte`.
- **Notification feed**: Reads a JSON feed (default: `/notifications.json`) with items: `{ id, icon, href, title, count }`.
- **Portal-aware**: The toolbar auto-selects the feed URL based on the `portal` prop (`admin`, `teachers`, `guardian`).
- **Icons**: SVG icons are mapped in `src/lib/icons.ts` and referenced by name in the feed.
- **Admin UI**: `ps-toolbar-admin` custom element provides an admin interface for managing toolbar icons.
- **Build system**: Uses Vite (see `vite.config.ts`). Output is placed in `dist/WEB_ROOT/ps-toolbar/`.
- **Integration**: Built assets are intended for PowerSchool extension or direct app bar hook integration.

## Key Conventions

- **Props normalization**: The toolbar normalizes portal/user type props (`portal`, `usertype`, `userType`, `user-type`).
- **State management**: Uses `$state` and `$derived` for reactivity in Svelte components.
- **Icon registration**: Add new icons to `iconSvgMap` in `src/lib/icons.ts` and to `availableIcons` in `IconForm.svelte`.
- **Testing**: Uses Vitest and @testing-library/svelte. Mock fetches in tests as shown in `Toolbar.test.ts`.
- **Environment**: In production, set `VITE_NOTIF_FEED_URL` for the live JSON endpoint.

---

If you want to configure MCP servers (e.g., Playwright for E2E testing), let me know!

---

This file summarizes build/test commands, architecture, and conventions for Copilot. Would you like to adjust anything or add coverage for other areas?