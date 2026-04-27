import React from 'react';
import LegacyFramePage from '../components/LegacyFramePage';

interface NewsJapanTimesAILovePageProps {
  locale?: 'en' | 'zh-CN';
}

export default function NewsJapanTimesAILovePage({ locale = 'en' }: NewsJapanTimesAILovePageProps): JSX.Element {
  const pagePath = locale === 'zh-CN'
    ? '/legacy/zh/news-japantimes-ai-love.html'
    : '/legacy/news-japantimes-ai-love.html';

  const title = locale === 'zh-CN'
    ? 'Japan Times与AI情感研究'
    : 'Japan Times on AI Attachment';

  return <LegacyFramePage title={title} pagePath={pagePath} />;
}
