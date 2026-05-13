import React from 'react';
import NewsListContent from '../components/news/NewsListContent';
import SitePageShell from '../components/SitePageShell';
import { newsPageContent } from '../data/legacyPageContent';
import { legacyPageStylesheets } from '../data/legacyStylesheets';
import { useDocumentTitle } from '../hooks/useDocumentTitle';
import { usePageStylesheets } from '../hooks/usePageStylesheets';
import type { Locale } from '../types/common';
import '../styles/migrated-legacy-pages.css';

interface NewsPageProps {
  locale?: Locale;
}

const NEWS_STYLESHEETS = legacyPageStylesheets('/legacy/news.css');

export default function NewsPage({ locale = 'en' }: NewsPageProps): JSX.Element {
  useDocumentTitle('MC2 | HKUST(GZ), HKUST');
  usePageStylesheets(NEWS_STYLESHEETS);

  const content = newsPageContent[locale].list;
  return (
    <SitePageShell className='news-page-shell' ariaLabel={locale === 'zh-CN' ? '新闻' : 'News'} locale={locale} activeRoute='news'>
      <NewsListContent content={content} />
    </SitePageShell>
  );
}
