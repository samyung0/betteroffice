export type LayoutUpdateOrigin = 'local' | 'remote';

export function mergeLayoutUpdateOrigin(
  current: LayoutUpdateOrigin | null,
  next: LayoutUpdateOrigin
): LayoutUpdateOrigin {
  return current === 'local' || next === 'local' ? 'local' : 'remote';
}

export interface ViewportAnchorSnapshot {
  viewportOffset: number;
  scrollTopSnapshot: number;
}

export interface ScrollRestoreTicket<T> {
  generation: number;
  value: T;
}

export class PendingScrollRestoreController<T> {
  private generation = 0;
  private pending: ScrollRestoreTicket<T> | null = null;

  capture(value: T): ScrollRestoreTicket<T> {
    const ticket = { generation: ++this.generation, value };
    this.pending = ticket;
    return ticket;
  }

  cancel(): void {
    this.generation += 1;
    this.pending = null;
  }

  peek(): ScrollRestoreTicket<T> | null {
    return this.pending;
  }

  take(): ScrollRestoreTicket<T> | null {
    const ticket = this.pending;
    this.pending = null;
    return ticket && this.isCurrent(ticket) ? ticket : null;
  }

  isCurrent(ticket: ScrollRestoreTicket<T>): boolean {
    return ticket.generation === this.generation;
  }

  run(ticket: ScrollRestoreTicket<T>, restore: () => void): boolean {
    if (!this.isCurrent(ticket)) return false;
    restore();
    return true;
  }
}

export function computeViewportAnchoredScrollTop(
  anchor: ViewportAnchorSnapshot,
  nextTargetTop: number | null,
  maxScrollTop: number
): number {
  const requested =
    nextTargetTop == null ? anchor.scrollTopSnapshot : nextTargetTop - anchor.viewportOffset;
  return Math.min(Math.max(0, requested), Math.max(0, maxScrollTop));
}

export function shouldScrollCaretIntoView(
  layoutUpdateOrigin: LayoutUpdateOrigin,
  selectionChanged: boolean
): boolean {
  return layoutUpdateOrigin === 'local' || selectionChanged;
}

/**
 * The scrollTop that keeps the document point at the pane's top in place when
 * the pages' scale goes from `from` to `to`. `columnTop` is the first page's top
 * in the scroll content and `gap` the space between pages, neither of which
 * scales; `heights` are the page heights at 100%.
 */
export function rescaleScrollTop(
  scrollTop: number,
  columnTop: number,
  heights: readonly number[],
  gap: number,
  from: number,
  to: number
): number {
  if (scrollTop <= columnTop) return scrollTop;
  let top = columnTop;
  let next = columnTop;
  for (const height of heights) {
    const into = scrollTop - top;
    if (into <= height * from) return next + (into / from) * to;
    if (into <= height * from + gap) return next + height * to + (into - height * from);
    top += height * from + gap;
    next += height * to + gap;
  }
  return next + (scrollTop - top);
}
