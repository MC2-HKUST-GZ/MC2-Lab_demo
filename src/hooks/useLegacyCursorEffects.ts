import { useEffect } from 'react';

export function useLegacyCursorEffects(): void {
  useEffect(() => {
    // Legacy cursor effects are handled inside legacy pages loaded in iframe.
  }, []);
}
