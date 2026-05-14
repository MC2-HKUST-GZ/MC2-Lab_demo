import recentPublicationDataJson from './recentPublications.json';

export interface RecentPublicationEntry {
  title: string;
  authors: string;
  venue: string;
}

export const recentPublicationData: Record<string, RecentPublicationEntry[]> = recentPublicationDataJson;
