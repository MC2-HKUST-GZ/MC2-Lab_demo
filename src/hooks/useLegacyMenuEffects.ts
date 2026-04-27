import { useEffect } from 'react';

export function useLegacyMenuEffects(): void {
  useEffect(() => {
    // Legacy menu interactions are handled inside legacy pages loaded in iframe.
  }, []);
}
