import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { YrsSession, YrsTocLayout } from '@betteroffice/docx/yrs';

/** Twips per CSS pixel at 96 dpi. */
const TWIPS_PER_PX = 15;

/**
 * What a table of contents takes from `layout`: each listed heading's page as
 * its page shows the number (a heading that has no fragment, hidden text, takes
 * the page of the paragraph before it), and Word's right tab at the text width
 * less 10 twips of the page `paraId` starts on, else the first page.
 */
export function tocLayout(
  session: YrsSession,
  layout: Layout,
  paraId: string | null,
  emptyText: string
): YrsTocLayout {
  const firstPage = new Map<string, Layout['pages'][number]>();
  for (const page of layout.pages)
    for (const fragment of page.fragments)
      if (fragment.kind === 'paragraph' && typeof fragment.blockId === 'string')
        if (!firstPage.has(fragment.blockId)) firstPage.set(fragment.blockId, page);
  const label = (page: Layout['pages'][number]) => page.pageLabel ?? String(page.number);
  const listed = new Set(session.tableOfContentsHeadings().map((heading) => heading.paraId));
  const pages: Record<string, string> = {};
  let previous = layout.pages[0];
  for (const { paraId: id } of session.paragraphs('body')) {
    previous = firstPage.get(id) ?? previous;
    if (listed.has(id) && previous) pages[id] = label(previous);
  }
  const page = (paraId && firstPage.get(paraId)) || layout.pages[0];
  const width = page
    ? page.size.w - page.margins.left - page.margins.right - (page.margins.gutter ?? 0)
    : 0;
  return { pages, tabTwips: Math.max(0, Math.round(width * TWIPS_PER_PX) - 10), emptyText };
}

/** Whether two tables of contents would show the same page numbers. */
export function samePages(a: YrsTocLayout, b: YrsTocLayout): boolean {
  const keys = Object.keys(a.pages);
  return (
    keys.length === Object.keys(b.pages).length && keys.every((id) => a.pages[id] === b.pages[id])
  );
}
