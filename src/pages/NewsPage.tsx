import React from 'react';
import LegacyFramePage from '../components/LegacyFramePage';
import { useNewsModalQueryEffect  } from '../hooks/useNewsModalQueryEffect';

interface NewsPageProps {
  locale?: 'en' | 'zh-CN' ;
}

export default function NewsPage({ locale = 'en' }: NewsPageProps): JSX.Element {
  const modalQuery = useNewsModalQueryEffect();
  const basePath = locale === 'zh-CN' ? '/legacy/zh/news.html' : '/legacy/news.html';
  return <LegacyFramePage title="News" pagePath={`${basePath}${modalQuery}`} />;
}
