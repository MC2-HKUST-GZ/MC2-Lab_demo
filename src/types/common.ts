export type Locale = 'en' | 'zh-CN';

export type RouteKey =
  | 'home'
  | 'people'
  | 'publication'
  | 'project'
  | 'news'
  | 'leader';

export interface PublicationSubmission {
  title: string;
  conference: string;
  photo: string;
  representative_work: string;
  link: string;
}

export interface PublicationRecord {
  id: string;
  title: string;
  conference: string;
  photo: string;
  representativeWork: string;
  link: string;
}

export interface LeaderProfile {
  slug: 'pan-hui';
  name: string;
  photo: string;
  title: { en: string; zh: string };
  summary: { en: string; zh: string };
  researchFocus: { en: string[]; zh: string[] };
  honors: { en: string[]; zh: string[] };
  highlights: { en: string[]; zh: string[] };
  links: { homepage?: string; scholar?: string; email?: string };
}
