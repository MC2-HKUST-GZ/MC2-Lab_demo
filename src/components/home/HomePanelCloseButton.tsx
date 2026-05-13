import React from 'react';
import { runOnEnterOrSpace } from './homeKeyboard';

interface HomePanelCloseButtonProps {
  className: string;
  onClose: () => void;
}

export default function HomePanelCloseButton({ className, onClose }: HomePanelCloseButtonProps): JSX.Element {
  return (
    <div
      className={className}
      role='button'
      tabIndex={0}
      onClick={onClose}
      onKeyDown={(event) => runOnEnterOrSpace(event, onClose)}
    >
      Close
    </div>
  );
}
