import React from 'react';
import NewsArticleContent from '../components/news/NewsArticleContent';
import SitePageShell from '../components/SitePageShell';
import { newsPageContent } from '../data/legacyPageContent';
import { legacyPageStylesheets } from '../data/legacyStylesheets';
import { useDocumentTitle } from '../hooks/useDocumentTitle';
import { usePageStylesheets } from '../hooks/usePageStylesheets';
import type { Locale } from '../types/common';
import '../styles/migrated-legacy-pages.css';

interface NewsNatureSpotlightPageProps {
  locale?: Locale;
}

const NEWS_STYLESHEETS = legacyPageStylesheets('/legacy/news.css');

export default function NewsNatureSpotlightPage({ locale = 'en' }: NewsNatureSpotlightPageProps): JSX.Element {
  useDocumentTitle('MC2 | HKUST(GZ), HKUST');
  usePageStylesheets(NEWS_STYLESHEETS);

  const article = newsPageContent[locale].articles['nature-spotlight'];
  return (
    <SitePageShell className='news-page-shell' ariaLabel={locale === 'zh-CN' ? 'Nature专题聚焦' : 'Nature Spotlight'} locale={locale} activeRoute='news'>
      <NewsArticleContent article={article} />
    </SitePageShell>
  );
}
