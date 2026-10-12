import { useCallback, useEffect, useRef } from 'react';
import type { DisplayList } from '@betteroffice/docx/layout/render';

/**
 * The newest display list with pages, waiting for one: right after an edit the
 * list can have none until the edit's layout lands, and Print or a list-card
 * capture drawn then would draw nothing.
 */
export function usePagedList(list: DisplayList | null): () => Promise<DisplayList> {
  const latest = useRef(list);
  latest.current = list;
  const waiters = useRef(new Set<(list: DisplayList) => void>());
  useEffect(() => {
    if (!list?.pages.length) return;
    for (const resume of waiters.current) resume(list);
    waiters.current.clear();
  }, [list]);
  return useCallback(() => {
    const now = latest.current;
    if (now?.pages.length) return Promise.resolve(now);
    return new Promise<DisplayList>((resume) => waiters.current.add(resume));
  }, []);
}
