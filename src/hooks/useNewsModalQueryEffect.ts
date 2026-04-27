import { useMemo } from 'react';
import { useLocation } from 'react-router-dom';

export function useNewsModalQueryEffect(): string {
  const { search } = useLocation();

  return useMemo(() => {
    if (!search) return '';
    const params = new URLSearchParams(search);
    if (params.get('action') === 'showModal') {
      return `?${params.toString()}`;
    }
    return '';
  }, [search]);
}
