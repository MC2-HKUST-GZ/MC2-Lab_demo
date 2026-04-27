import React from 'react';
import LegacyFramePage from '../components/LegacyFramePage';

interface HomePageProps {
  locale?: 'en' | 'zh-CN';
}

export default function HomePage({ locale = 'en' }: HomePageProps): JSX.Element {
  const pagePath = locale === 'zh-CN' ? '/legacy/zh/index.html' : '/legacy/index.html';
  return <LegacyFramePage title="Home" pagePath={pagePath} />;
}
