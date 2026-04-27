import React from 'react';
import LegacyFramePage from '../components/LegacyFramePage';

interface PeoplePageProps {
  locale?: 'en' | 'zh-CN';
}

export default function PeoplePage({ locale = 'en' }: PeoplePageProps): JSX.Element {
  const pagePath = locale === 'zh-CN' ? '/legacy/zh/people.html' : '/legacy/people.html';
  return <LegacyFramePage title="People" pagePath={pagePath} />;
}
