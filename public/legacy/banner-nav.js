const APP_ROUTE_SET = new Set([
  '/',
  '/people',
  '/publication',
  '/project',
  '/news',
  '/leader',
  '/news/ai-as-catalyst-workshop',
  '/news/japantimes-ai-love',
  '/news/nature-spotlight',
  '/zh',
  '/zh/people',
  '/zh/publication',
  '/zh/project',
  '/zh/news',
  '/zh/news/ai-as-catalyst-workshop',
  '/zh/leader',
  '/zh/news/japantimes-ai-love',
  '/zh/news/nature-spotlight'
]);

const LEGACY_FILE_ROUTE_MAP = {
  'index.html': '',
  'people.html': 'people',
  'publication.html': 'publication',
  'project.html': 'project',
  'news.html': 'news',
  'leader.html': 'leader',
  'news-ai-as-catalyst-workshop.html': 'news/ai-as-catalyst-workshop',
  'news-japantimes-ai-love.html': 'news/japantimes-ai-love',
  'news-mc2-nature-spotlight.html': 'news/nature-spotlight'
};

const HTML_ALIAS_MAP = {
  '/index.html': '/',
  '/people.html': '/people',
  '/publication.html': '/publication',
  '/project.html': '/project',
  '/news.html': '/news',
  '/leader.html': '/leader',
  '/news/ai-as-catalyst-workshop.html': '/news/ai-as-catalyst-workshop',
  '/news/japantimes-ai-love.html': '/news/japantimes-ai-love',
  '/news/nature-spotlight.html': '/news/nature-spotlight',
  '/zh/index.html': '/zh',
  '/zh/people.html': '/zh/people',
  '/zh/publication.html': '/zh/publication',
  '/zh/project.html': '/zh/project',
  '/zh/news.html': '/zh/news',
  '/zh/leader.html': '/zh/leader',
  '/zh/news/ai-as-catalyst-workshop.html': '/zh/news/ai-as-catalyst-workshop',
  '/zh/news/japantimes-ai-love.html': '/zh/news/japantimes-ai-love',
  '/zh/news/nature-spotlight.html': '/zh/news/nature-spotlight'
};

function routeFromLegacyFile(filename, isZh) {
  const slug = LEGACY_FILE_ROUTE_MAP[filename];
  if (typeof slug === 'undefined') return null;
  const prefix = isZh ? '/zh' : '';
  if (!slug) {
    return prefix || '/';
  }
  return `${prefix}/${slug}`;
}

function toCanonicalRoute(pathname) {
  const normalizedPath = pathname && pathname.length > 1 && pathname.endsWith('/')
    ? pathname.slice(0, -1)
    : pathname || '/';

  if (HTML_ALIAS_MAP[normalizedPath]) {
    return HTML_ALIAS_MAP[normalizedPath];
  }

  if (APP_ROUTE_SET.has(normalizedPath)) {
    return normalizedPath;
  }

  const parts = normalizedPath.split('/').filter(Boolean);
  if (parts[0] !== 'legacy') return null;

  const isZh = parts[1] === 'zh';
  const filename = parts[isZh ? 2 : 1] || 'index.html';
  return routeFromLegacyFile(filename, isZh);
}

function toCanonicalHref(rawHref) {
  if (!rawHref) return null;
  if (rawHref.startsWith('#')) return null;
  if (rawHref.startsWith('mailto:') || rawHref.startsWith('tel:') || rawHref.startsWith('javascript:')) {
    return null;
  }

  let url;
  try {
    url = new URL(rawHref, window.location.href);
  } catch (error) {
    return null;
  }

  if (url.origin !== window.location.origin) return null;
  const canonicalRoute = toCanonicalRoute(url.pathname);
  if (!canonicalRoute || !APP_ROUTE_SET.has(canonicalRoute)) return null;
  return `${canonicalRoute}${url.search}${url.hash}`;
}

function toAlternateLocaleRoute(route) {
  if (route === '/zh') return '/';
  if (route.startsWith('/zh/')) return route.slice(3);
  if (route === '/') return '/zh';
  return `/zh${route}`;
}

function normalizeLocalLinksToTopRoute() {
  document.querySelectorAll('a[href]').forEach(function (anchor) {
    const canonicalHref = toCanonicalHref(anchor.getAttribute('href'));
    if (!canonicalHref) return;

    anchor.setAttribute('href', canonicalHref);
    if (anchor.getAttribute('target') !== '_blank') {
      anchor.setAttribute('target', '_top');
      anchor.setAttribute('rel', 'noopener');
    }
  });
}

function isModifiedClick(event) {
  return event.metaKey || event.ctrlKey || event.shiftKey || event.altKey || event.button !== 0;
}

function normalizePath(pathname) {
  if (!pathname) return '/';
  if (pathname.length > 1 && pathname.endsWith('/')) {
    return pathname.slice(0, -1);
  }
  return pathname;
}

function attachRouteTransitionNavigation(anchor, header, closeMenu) {
  anchor.addEventListener('click', function (event) {
    if (event.defaultPrevented || isModifiedClick(event)) return;
    if (anchor.getAttribute('download')) return;

    const target = (anchor.getAttribute('target') || '').toLowerCase();
    if (target === '_blank') return;

    const canonicalHref = toCanonicalHref(anchor.getAttribute('href'));
    if (!canonicalHref) return;

    const currentCanonicalRoute = toCanonicalRoute(window.location.pathname) || normalizePath(window.location.pathname);
    const currentCanonicalHref = `${currentCanonicalRoute}${window.location.search || ''}${window.location.hash || ''}`;
    if (currentCanonicalHref === canonicalHref) {
      closeMenu();
      return;
    }

    event.preventDefault();
    closeMenu();
    if (header.classList.contains('is-route-leaving')) return;

    if (window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
      window.top.location.assign(canonicalHref);
      return;
    }

    const leaveMs = window.innerWidth <= 980 ? 70 : 85;
    header.classList.add('is-route-leaving');
    window.setTimeout(function () {
      window.top.location.assign(canonicalHref);
    }, leaveMs);
  });
}

function initBannerNav() {
  const header = document.querySelector('.major-nav');
  const nav = document.querySelector('.major-nav .nav');

  if (!header || !nav) return;
  if (header.querySelector('.banner-controls')) return;

  normalizeLocalLinksToTopRoute();

  const canonicalRoute = toCanonicalRoute(window.location.pathname) || '/';
  const isZh = canonicalRoute === '/zh' || canonicalRoute.startsWith('/zh/');
  const query = window.location.search || '';
  const hash = window.location.hash || '';
  const alternateRoute = toAlternateLocaleRoute(canonicalRoute);

  const controls = document.createElement('div');
  controls.className = 'banner-controls';

  let lang = null;
  if (APP_ROUTE_SET.has(alternateRoute)) {
    lang = document.createElement('a');
    lang.className = 'banner-lang';
    lang.textContent = isZh ? 'EN' : '中文';
    lang.href = `${alternateRoute}${query}${hash}`;
    lang.setAttribute('target', '_top');
    lang.setAttribute('rel', 'noopener');
  }

  const menuButton = document.createElement('button');
  menuButton.type = 'button';
  menuButton.className = 'banner-menu-toggle';
  menuButton.id = 'bannerMenuToggle';
  menuButton.textContent = isZh ? '菜单' : 'Menu';
  menuButton.setAttribute('aria-expanded', 'false');
  menuButton.setAttribute('aria-label', isZh ? '打开导航菜单' : 'Open navigation menu');

  if (lang) {
    controls.appendChild(lang);
  }
  controls.appendChild(menuButton);
  header.appendChild(controls);

  const closeMenu = function () {
    header.classList.remove('is-menu-open');
    menuButton.setAttribute('aria-expanded', 'false');
  };

  menuButton.addEventListener('click', function () {
    const opening = !header.classList.contains('is-menu-open');
    header.classList.toggle('is-menu-open', opening);
    menuButton.setAttribute('aria-expanded', String(opening));
  });

  document.addEventListener('click', function (event) {
    if (!header.contains(event.target)) {
      closeMenu();
    }
  });

  nav.querySelectorAll('a').forEach(function (anchor) {
    attachRouteTransitionNavigation(anchor, header, closeMenu);
  });
  if (lang) {
    attachRouteTransitionNavigation(lang, header, closeMenu);
  }

  document.querySelectorAll('.major-nav .nav-logo a[href]').forEach(function (anchor) {
    attachRouteTransitionNavigation(anchor, header, closeMenu);
  });

  window.addEventListener('resize', function () {
    if (window.innerWidth > 980) {
      closeMenu();
    }
  });

  // Auto-hide header on downward scroll; reveal on upward scroll or top-edge hover.
  let lastScrollY = window.scrollY || 0;
  const revealZone = 78;
  const hideAfter = 96;
  const delta = 8;

  const showHeader = function () {
    header.classList.remove('is-scroll-hidden');
  };

  const hideHeader = function () {
    if (!header.classList.contains('is-menu-open')) {
      header.classList.add('is-scroll-hidden');
    }
  };

  window.addEventListener('scroll', function () {
    const currentY = window.scrollY || 0;
    const direction = currentY - lastScrollY;

    if (header.classList.contains('is-menu-open')) {
      showHeader();
      lastScrollY = currentY;
      return;
    }

    if (currentY <= revealZone) {
      showHeader();
    } else if (direction > delta && currentY > hideAfter) {
      hideHeader();
    } else if (direction < -delta) {
      showHeader();
    }

    lastScrollY = currentY;
  }, { passive: true });

  document.addEventListener('mousemove', function (event) {
    if (event.clientY <= revealZone) {
      showHeader();
    }
  });

  document.addEventListener('touchstart', function (event) {
    if (!event.touches || event.touches.length === 0) return;
    if (event.touches[0].clientY <= revealZone) {
      showHeader();
    }
  }, { passive: true });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initBannerNav);
} else {
  initBannerNav();
}
