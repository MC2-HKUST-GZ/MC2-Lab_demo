import { useEffect, type RefObject } from 'react';

export function useCloseOnOutsideInteraction<T extends HTMLElement>(
  targetRef: RefObject<T>,
  enabled: boolean,
  onClose: () => void
): void {
  useEffect(() => {
    if (!enabled) return undefined;

    const closeOnOutsideInteraction = (event: MouseEvent | TouchEvent) => {
      const target = event.target;
      if (target instanceof Node && !targetRef.current?.contains(target)) {
        onClose();
      }
    };

    window.addEventListener('click', closeOnOutsideInteraction);
    window.addEventListener('touchstart', closeOnOutsideInteraction);
    return () => {
      window.removeEventListener('click', closeOnOutsideInteraction);
      window.removeEventListener('touchstart', closeOnOutsideInteraction);
    };
  }, [enabled, onClose, targetRef]);
}

