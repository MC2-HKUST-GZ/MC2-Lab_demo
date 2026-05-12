# MC2 Lab Website (React + Legacy Pages)

Official website project for the Center for Metaverse and Computational Creativity (MC2), HKUST and HKUST(GZ).

Author: **Dongyijie Primo Pan**

## 1. Project Structure

This project uses a React router shell with legacy HTML content rendered in iframes.

- `src/router/routes.tsx`: canonical route definitions and compatibility redirects.
- `src/pages/*.tsx`: React page wrappers that point to legacy files.
- `src/components/LegacyFramePage.tsx`: shared iframe container for legacy pages.
- `public/legacy/`: English legacy pages, shared banner scripts/styles.
- `public/legacy/zh/`: Chinese legacy pages and localized assets.
- `images/` and `public/images/`: website images and news cover assets.
- `tools/update-publications.ts`: publication content update script.

## 2. Run and Build

```bash
npm install
npm start
```

Production build:

```bash
npm run build
```

Update publications and build:

```bash
npm run update:all
```

## 3. Page and Route Responsibilities

Canonical routes are used for navigation and browser history. Old `.html` routes are kept as redirects for compatibility.

English pages:

- `/`: Home (`public/legacy/index.html`)
- `/people`: People (`public/legacy/people.html`)
- `/publication`: Publications (`src/pages/PublicationPage.tsx`)
- `/project`: Projects (`public/legacy/project.html`)
- `/news`: News list (`public/legacy/news.html`)
- `/news/nature-spotlight`: Nature spotlight news detail (`public/legacy/news-mc2-nature-spotlight.html`)
- `/news/japantimes-ai-love`: Japan Times news detail (`public/legacy/news-japantimes-ai-love.html`)
- `/leader`: Director page (`src/pages/LeaderPage.tsx`)

Chinese pages:

- `/zh`: Home (`public/legacy/zh/index.html`)
- `/zh/people`: People (`public/legacy/zh/people.html`)
- `/zh/publication`: Publications (`src/pages/PublicationPage.tsx`)
- `/zh/project`: Projects (`public/legacy/zh/project.html`)
- `/zh/news`: News list (`public/legacy/zh/news.html`)
- `/zh/news/nature-spotlight`: Nature spotlight detail (`public/legacy/zh/news-mc2-nature-spotlight.html`)
- `/zh/news/japantimes-ai-love`: Japan Times detail (`public/legacy/zh/news-japantimes-ai-love.html`)
- `/zh/leader`: Director page (`src/pages/LeaderPage.tsx`)

## 4. News Maintenance Rules

When adding or updating news content, follow these rules:

1. Add the new card in both news list pages:
- `public/legacy/news.html`
- `public/legacy/zh/news.html`

2. Add detail pages in both languages:
- `public/legacy/news-<slug>.html`
- `public/legacy/zh/news-<slug>.html`

3. Register canonical routes in `src/router/routes.tsx`:
- `/news/<slug>`
- `/zh/news/<slug>`

4. Add `.html` alias redirects to canonical routes for backward compatibility.

5. Use local detail pages as link targets (never link list cards directly to third-party sources).

6. Keep publication date aligned with the original source publication date.

7. Include explicit source attribution and research/paper links in each detail page.

8. Keep English and Chinese content synchronized in structure and key claims.

9. If a new legacy detail filename is introduced, update route maps in:
- `public/legacy/banner-nav.js`
- `public/legacy/zh/banner-nav.js`

## 5. Navigation and History Rules

To keep browser back/forward behavior stable:

- Use canonical React routes (no direct iframe-only internal navigation for core pages).
- Keep top-level navigation links targeting app routes (for example `/news`, `/zh/news`).
- Keep `.html` URLs as redirects only, not as primary user-facing routes.
- Validate: Home -> News list -> News detail -> browser back should return to News list.

## 6. Visual and Content Consistency Rules

- Keep MC squared formatting as `MC<sup>2</sup>` in HTML when superscript rendering is required.
- Keep date format consistent per language:
  - English: `13 Feb 2026`
  - Chinese: `2026年2月13日`
- Reuse existing card and article styles unless a full visual refresh is requested.
- Ensure all new assets are available in `public/images/` and use stable file names.

## 7. Pre-merge Checklist

- `npm run build` passes.
- New routes open correctly in both languages.
- News list links open local detail pages, not external source pages.
- Browser back from news detail returns to news list.
- Source links and paper links are valid.
