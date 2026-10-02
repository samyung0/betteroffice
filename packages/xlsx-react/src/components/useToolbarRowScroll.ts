import { type RefObject, useEffect, useState } from 'react';

/** Open dropdowns and menus, which render inside the row. */
const POPUP = '[role="dialog"], [role="menu"], [role="listbox"]';

/**
 * The single-row toolbar's scrolling part (as docx-react's). A vertical wheel
 * scrolls it sideways, so a mouse without a horizontal wheel reaches the
 * clipped groups; the result says which edges have more to scroll.
 */
export function useToolbarRowScroll(ref: RefObject<HTMLElement | null>, enabled: boolean) {
  const [edges, setEdges] = useState({ start: false, end: false });

  useEffect(() => {
    const row = ref.current;
    if (!enabled || !row) return;

    const mark = () => {
      const max = row.scrollWidth - row.clientWidth;
      const start = row.scrollLeft > 1;
      const end = row.scrollLeft < max - 1;
      setEdges((current) =>
        current.start === start && current.end === end ? current : { start, end }
      );
    };

    const onWheel = (event: WheelEvent) => {
      if (
        event.defaultPrevented ||
        !event.cancelable ||
        event.ctrlKey ||
        event.metaKey ||
        event.shiftKey ||
        event.deltaX !== 0 ||
        !event.deltaY
      )
        return;
      const rect = row.getBoundingClientRect();
      if (
        (event.target instanceof Element && event.target.closest(POPUP)) ||
        event.clientY < rect.top ||
        event.clientY > rect.bottom
      )
        return;
      const max = row.scrollWidth - row.clientWidth;
      if (max <= 0) return;
      let delta = event.deltaY;
      if (event.deltaMode === WheelEvent.DOM_DELTA_LINE)
        delta *= Number.parseFloat(getComputedStyle(row).lineHeight) || 16;
      else if (event.deltaMode === WheelEvent.DOM_DELTA_PAGE) delta *= row.clientWidth;
      const next = Math.max(0, Math.min(max, row.scrollLeft + delta));
      if (next === row.scrollLeft) return;
      event.preventDefault();
      row.scrollLeft = next;
    };

    const resize = new ResizeObserver(mark);
    const watch = () => {
      resize.disconnect();
      resize.observe(row);
      for (const child of Array.from(row.children)) resize.observe(child);
      mark();
    };
    const children = new MutationObserver(watch);
    children.observe(row, { childList: true });
    watch();
    row.addEventListener('scroll', mark, { passive: true });
    row.addEventListener('wheel', onWheel, { passive: false });
    return () => {
      children.disconnect();
      resize.disconnect();
      row.removeEventListener('scroll', mark);
      row.removeEventListener('wheel', onWheel);
    };
  }, [ref, enabled]);

  return edges;
}
