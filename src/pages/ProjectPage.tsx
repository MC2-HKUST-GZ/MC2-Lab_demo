import React from 'react';
import ProjectContent from '../components/project/ProjectContent';
import SitePageShell from '../components/SitePageShell';
import { projectPageContent } from '../data/legacyPageContent';
import { legacyPageStylesheets } from '../data/legacyStylesheets';
import { useBodyClass } from '../hooks/useBodyClass';
import { useDocumentTitle } from '../hooks/useDocumentTitle';
import { usePageStylesheets } from '../hooks/usePageStylesheets';
import type { Locale } from '../types/common';
import '../styles/migrated-legacy-pages.css';

interface ProjectPageProps {
  locale?: Locale;
}

const PROJECT_STYLESHEETS: Record<Locale, string[]> = {
  en: legacyPageStylesheets('/legacy/project.css'),
  'zh-CN': legacyPageStylesheets('/legacy/zh/project.css')
};

export default function ProjectPage({ locale = 'en' }: ProjectPageProps): JSX.Element {
  useDocumentTitle('MC2 | HKUST(GZ), HKUST');
  useBodyClass('publication-on');
  usePageStylesheets(PROJECT_STYLESHEETS[locale]);

  const content = projectPageContent[locale];
  return (
    <SitePageShell className='project-page-shell' ariaLabel={locale === 'zh-CN' ? '项目' : 'Project'} locale={locale} activeRoute='project'>
      <ProjectContent content={content} />
    </SitePageShell>
  );
}
