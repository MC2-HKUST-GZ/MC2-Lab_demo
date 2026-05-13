export type HomePanel = 'about' | 'contact' | null;

export interface HomeAboutCopy {
  intro: string;
  leaderPrefix: string;
  leaderName: string;
  leaderHref: string;
  leaderSuffix: string;
  closing: string;
}

export interface HomeContactLink {
  label: string;
  href: string;
}

export interface HomeResearchFocusItem {
  title: string;
  body: string;
}
