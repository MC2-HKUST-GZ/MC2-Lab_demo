import React from 'react';
import { useLegacyCursorEffects } from '../hooks/useLegacyCursorEffects';
import { useLegacyMenuEffects } from '../hooks/useLegacyMenuEffects';

interface LegacyFramePageProps {
  pagePath: string;
  title: string;
}

export default function LegacyFramePage({ pagePath, title }: LegacyFramePageProps): JSX.Element {
  useLegacyCursorEffects();
  useLegacyMenuEffects();

  return (
    <main className="legacy-page-shell" aria-label={title}>
      <iframe
        className="legacy-frame"
        title={title}
        src={pagePath}
        loading="eager"
      />
    </main>
  );
}
