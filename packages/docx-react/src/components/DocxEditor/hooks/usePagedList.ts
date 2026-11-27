import { useCallback, useEffect, useRef } from 'react';
import type { DisplayList } from '@betteroffice/docx/layout/render';

/** Every page laid out and built (a first frame builds only those in view, Q815 C). */
const whole = (list: DisplayList | null): list is DisplayList =>
  Boolean(list?.pages.length) && !list!.pages.some((page) => page.unbuilt);

/**
 * The newest display list with every page, waiting for one: right after an
 * edit the list can have none until the edit's layout lands, and while an
 * editor opens the pages out of view are still being built; Print or a
 * list-card capture drawn then would draw pages missing.
 */
export function usePagedList(list: DisplayList | null): () => Promise<DisplayList> {
  const latest = useRef(list);
  latest.current = list;
  const waiters = useRef(new Set<(list: DisplayList) => void>());
  useEffect(() => {
    if (!whole(list)) return;
    for (const resume of waiters.current) resume(list);
    waiters.current.clear();
  }, [list]);
  return useCallback(() => {
    const now = latest.current;
    if (whole(now)) return Promise.resolve(now);
    return new Promise<DisplayList>((resume) => waiters.current.add(resume));
  }, []);
}
