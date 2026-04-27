import React from 'react';
import LegacyFramePage from '../components/LegacyFramePage';

interface NewsAIAsCatalystWorkshopPageProps {
  locale?: 'en' | 'zh-CN';
}

export default function NewsAIAsCatalystWorkshopPage({ locale = 'en' }: NewsAIAsCatalystWorkshopPageProps): JSX.Element {
  const pagePath = locale === 'zh-CN'
    ? '/legacy/zh/news-ai-as-catalyst-workshop.html'
    : '/legacy/news-ai-as-catalyst-workshop.html';

  const title = locale === 'zh-CN'
    ? 'AI as Catalyst 工作坊'
    : 'AI as Catalyst Workshop';

  return <LegacyFramePage title={title} pagePath={pagePath} />;
}
