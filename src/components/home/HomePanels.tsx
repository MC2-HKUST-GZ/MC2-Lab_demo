import React from 'react';
import HomeAboutPanel from './HomeAboutPanel';
import HomeContactPanel from './HomeContactPanel';

interface HomePanelsProps {
  onClosePanel: () => void;
}

export default function HomePanels({ onClosePanel }: HomePanelsProps): JSX.Element {
  return (
    <>
      <HomeAboutPanel onClosePanel={onClosePanel} />
      <HomeContactPanel onClosePanel={onClosePanel} />
    </>
  );
}
