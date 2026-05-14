# MC2 Lab Website

React + TypeScript website for the Center for Metaverse and Computational Creativity (MC2), HKUST and HKUST(GZ).

Author: **Dongyijie Primo Pan**

## Project Structure

- `src/router/routes.tsx`: canonical app routes and compatibility redirects.
- `src/pages/`: route-level React pages.
- `src/components/`: shared UI plus page-specific React components.
- `src/components/news/`, `src/components/people/`, `src/components/project/`: migrated content views.
- `src/data/`: typed content sources and generated JSON data used by React pages.
- `src/data/recentPublications.json`: editable recent publication entries grouped by year.
- `src/i18n/`: shared navigation and language labels.
- `src/hooks/`: reusable browser and interaction hooks.
- `src/styles/`: application CSS and migrated page styling.
- `public/images/`: static image assets served by the app.
- `public/legacy/`: transitional legacy CSS only; no raw legacy HTML pages should be added.
- `content/submissions/`: publication source data.
- `tools/update-publications.ts`: publication data generator.

## Run and Build

```bash
npm install
npm start
```

Production build:

```bash
npm run build
```

Refresh publications and build:

```bash
npm run update:all
```

## Routes

Canonical routes are React routes. Old `.html` URLs are kept only as redirects for external compatibility.

English:

- `/`: Home
- `/people`: People
- `/publication`: Publications
- `/project`: Projects
- `/news`: News list
- `/news/:slug`: News detail pages
- `/leader`: Director page

Chinese:

- `/zh`: Home
- `/zh/people`: People
- `/zh/publication`: Publications
- `/zh/project`: Projects
- `/zh/news`: News list
- `/zh/news/:slug`: News detail pages
- `/zh/leader`: Director page

## Content Guidelines

- Build user-facing pages as React components, not iframe wrappers or standalone HTML files.
- Keep English and Chinese content synchronized in structure and key claims.
- Place static assets in `public/images/` and reference them with `/images/...`.
- Add new recent publication entries in `src/data/recentPublications.json`; `src/data/recentPublications.ts` is only the typed bridge used by React.
- Keep `public/legacy/` limited to transitional CSS that is still needed by migrated pages.
- Do not add new `public/legacy/**/*.html` files.
- Preserve old `.html` route aliases as redirects when a public URL has already existed.

## Visual QA

Before merging UI changes:

- `npm run build` passes.
- Desktop and mobile layouts match the existing visual behavior.
- Navigation works in both languages.
- News list cards open local React detail pages.
- Browser back/forward returns to the expected page.
- No raw legacy HTML or iframe page is introduced.
