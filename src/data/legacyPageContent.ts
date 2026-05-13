import newsContent from './newsContent.generated.json';
import peopleContent from './peopleContent.generated.json';
import projectContent from './projectContent.generated.json';
import type { LocalizedContent, NewsPageContent, PeoplePageContent, ProjectPageContent } from '../types/legacyPages';

export const peoplePageContent = peopleContent as LocalizedContent<PeoplePageContent>;
export const projectPageContent = projectContent as LocalizedContent<ProjectPageContent>;
export const newsPageContent = newsContent as LocalizedContent<NewsPageContent>;
