import { useLayoutEffect } from 'react';

export function usePageStylesheets(hrefs: string[]): void {
  useLayoutEffect(() => {
    const links = hrefs.map((href) => {
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = href;
      link.dataset.mc2PageStylesheet = href;
      document.head.appendChild(link);
      return link;
    });

    return () => {
      links.forEach((link) => link.remove());
    };
  }, [hrefs]);
}
