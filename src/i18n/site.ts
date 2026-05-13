import type { Locale, RouteKey } from '../types/common';

export type SiteTextKey = 'menuLabel' | 'menuAria' | 'languageLabel';

export interface SiteNavItem {
  href: string;
  label: string;
  active: boolean;
  route: RouteKey;
}

const routePaths: Record<Locale, Record<RouteKey, string>> = {
  en: {
    home: '/',
    people: '/people',
    publication: '/publication',
    project: '/project',
    news: '/news',
    leader: '/leader'
  },
  'zh-CN': {
    home: '/zh',
    people: '/zh/people',
    publication: '/zh/publication',
    project: '/zh/project',
    news: '/zh/news',
    leader: '/zh/leader'
  }
};

const routeLabels: Record<Locale, Record<RouteKey, string>> = {
  en: {
    home: 'Home',
    people: 'People',
    publication: 'Publication',
    project: 'Project',
    news: 'News',
    leader: 'Director'
  },
  'zh-CN': {
    home: '首页',
    people: '成员',
    publication: '论文',
    project: '项目',
    news: '新闻',
    leader: '负责人'
  }
};

const siteText: Record<Locale, Record<SiteTextKey, string>> = {
  en: {
    menuLabel: 'Menu',
    menuAria: 'Open navigation menu',
    languageLabel: '中文'
  },
  'zh-CN': {
    menuLabel: '菜单',
    menuAria: '打开导航菜单',
    languageLabel: 'EN'
  }
};

const navRouteOrder: RouteKey[] = ['home', 'people', 'publication', 'project', 'news', 'leader'];

export function isZhLocale(locale: Locale): boolean {
  return locale === 'zh-CN';
}

export function getRoutePath(route: RouteKey, locale: Locale): string {
  return routePaths[locale][route];
}

export function getSiteText(locale: Locale, key: SiteTextKey): string {
  return siteText[locale][key];
}

export function getAlternateLocale(locale: Locale): Locale {
  return isZhLocale(locale) ? 'en' : 'zh-CN';
}

export function getAlternateRoutePath(route: RouteKey, locale: Locale): string {
  return getRoutePath(route, getAlternateLocale(locale));
}

export function getSiteNavItems(locale: Locale, activeRoute: RouteKey): SiteNavItem[] {
  return navRouteOrder.map((route) => ({
    route,
    href: getRoutePath(route, locale),
    label: routeLabels[locale][route],
    active: route === activeRoute
  }));
}

