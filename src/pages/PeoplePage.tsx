import React from 'react';
import PeopleContent from '../components/people/PeopleContent';
import SitePageShell from '../components/SitePageShell';
import { peoplePageContent } from '../data/legacyPageContent';
import { legacyPageStylesheets } from '../data/legacyStylesheets';
import { useDocumentTitle } from '../hooks/useDocumentTitle';
import { usePageStylesheets } from '../hooks/usePageStylesheets';
import type { Locale } from '../types/common';
import '../styles/migrated-legacy-pages.css';

interface PeoplePageProps {
  locale?: Locale;
}

const PEOPLE_STYLESHEETS = legacyPageStylesheets('/legacy/people.css');

export default function PeoplePage({ locale = 'en' }: PeoplePageProps): JSX.Element {
  useDocumentTitle('MC2 | HKUST(GZ), HKUST');
  usePageStylesheets(PEOPLE_STYLESHEETS);

  const content = peoplePageContent[locale];
  return (
    <SitePageShell className='people-page-shell' ariaLabel={locale === 'zh-CN' ? '成员' : 'People'} locale={locale} activeRoute='people'>
      <PeopleContent content={content} />
    </SitePageShell>
  );
}
