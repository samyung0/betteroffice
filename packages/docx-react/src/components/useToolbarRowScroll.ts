import { type RefObject, useEffect } from 'react';

/**
 * The single-row toolbar's scrolling part. A vertical wheel scrolls it sideways,
 * so a mouse without a horizontal wheel reaches the clipped groups, and
 * `data-scroll-start` / `data-scroll-end` mark the edges that have more to
 * scroll, for the edge fades in every browser.
 */
export function useToolbarRowScroll(ref: RefObject<HTMLElement | null>, enabled: boolean) {
  useEffect(() => {
    const row = ref.current;
    if (!enabled || !row) return;

    const mark = () => {
      const max = row.scrollWidth - row.clientWidth;
      row.toggleAttribute('data-scroll-start', row.scrollLeft > 1);
      row.toggleAttribute('data-scroll-end', row.scrollLeft < max - 1);
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
      const max = row.scrollWidth - row.clientWidth;
      if (max <= 0) return;
      let delta = event.deltaY;
      if (event.deltaMode === WheelEvent.DOM_DELTA_LINE)
        delta *= Number.parseFloat(getComputedStyle(row).lineHeight) || 16;
      else if (event.deltaMode === WheelEvent.DOM_DELTA_PAGE) delta *= row.clientWidth;
      const next = Math.max(0, Math.min(max, row.scrollLeft + delta));
      if (next === row.scrollLeft) return;
      // Only wheel input this row consumes; the rest scrolls the page as usual.
      event.preventDefault();
      row.scrollLeft = next;
    };

    // Groups come and go (image and table controls) and change width, so the
    // children are watched as well as the row.
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
}
