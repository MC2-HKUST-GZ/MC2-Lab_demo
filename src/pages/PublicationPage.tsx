import React from 'react';
import LegacyFramePage from '../components/LegacyFramePage';

interface PublicationPageProps {
  locale?: 'en' | 'zh-CN';
}

export default function PublicationPage({ locale = 'en' }: PublicationPageProps): JSX.Element {
  const pagePath =
    locale === 'zh-CN' ? '/legacy/zh/publication.html' : '/legacy/publication.html';
  return <LegacyFramePage title="Publication" pagePath={pagePath} />;
}
