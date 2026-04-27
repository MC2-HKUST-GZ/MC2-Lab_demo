import React from 'react';
import LegacyFramePage from '../components/LegacyFramePage';

interface NewsNatureSpotlightPageProps {
  locale?: 'en' | 'zh-CN';
}

export default function NewsNatureSpotlightPage({ locale = 'en' }: NewsNatureSpotlightPageProps): JSX.Element {
  const pagePath = locale === 'zh-CN'
    ? '/legacy/zh/news-mc2-nature-spotlight.html'
    : '/legacy/news-mc2-nature-spotlight.html';
  const title = locale === 'zh-CN'
    ? 'MC² Nature专题聚焦'
    : 'MC² Nature Spotlight';

  return <LegacyFramePage title={title} pagePath={pagePath} />;
}
