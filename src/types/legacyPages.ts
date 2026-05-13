import type { Locale } from './common';

export type MemberRole = 'all' | 'chair' | 'postdoc' | 'admin' | 'phd' | 'mphil' | 'ra' | 'leader' | 'other';

export interface MemberProfile {
  key: string;
  name: string;
  role: Exclude<MemberRole, 'all'>;
  positionTitle: string;
  cohortTerm: string;
  cohortYear: number | null;
  photo: string;
  researchDirection: string;
  website: string;
  drivePhotoSource?: string;
}

export interface MemberRoleChip {
  role: MemberRole;
  label: string;
}

export interface AlumniEntry {
  name: string;
  detail: string;
}

export interface AlumniGroup {
  heading: string;
  entries: AlumniEntry[];
}

export interface PeoplePageContent {
  pageHeading: string;
  currentLabel: string;
  alumniLabel: string;
  searchPlaceholder: string;
  searchAria: string;
  roleFilterAria: string;
  emptyText: string;
  fallbackResearch: string;
  roleChips: MemberRoleChip[];
  members: MemberProfile[];
  alumniGroups: AlumniGroup[];
}

export interface ProjectItem {
  id: string;
  title: string;
  people: string;
  image: string;
  imageAlt: string;
  tags: string[];
  descriptions: string[];
  href: string;
}

export interface ProjectPageContent {
  pageHeading: string;
  projects: ProjectItem[];
}

export interface NewsListItem {
  id: string;
  kind: 'link' | 'modal';
  href: string;
  title: string;
  titleFontSize: string;
  image: string;
  imageAlt: string;
  meta: string;
  descriptions: string[];
  source: string;
}

export interface NewsYearSection {
  id: string;
  label: string;
  items: NewsListItem[];
}

export interface NewsModalContent {
  poster: string;
  posterAlt: string;
  srcSet: string;
  sizes: string;
  submitHref: string;
  submitAria: string;
  submitText: string;
  helperText: string;
}

export interface NewsListContent {
  pageHeading: string;
  years: NewsYearSection[];
  modal: NewsModalContent;
}

export type NewsArticleBlock =
  | { type: 'paragraph'; text: string }
  | { type: 'heading'; text: string }
  | { type: 'list'; items: string[] };

export interface NewsArticleLink {
  label: string;
  href: string;
}

export interface NewsArticleContent {
  pageHeading: string;
  title: string;
  meta: string;
  image: string;
  imageAlt: string;
  blocks: NewsArticleBlock[];
  links: NewsArticleLink[];
  backLabel: string;
  backHref: string;
}

export interface NewsPageContent {
  list: NewsListContent;
  articles: Record<'japantimes-ai-love' | 'nature-spotlight', NewsArticleContent>;
}

export type LocalizedContent<T> = Record<Locale, T>;
