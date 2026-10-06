/**
 * Selection and copy over a read-only viewer's text layer (the positioned a11y
 * mirror, made selectable by `CanvasPagesView selectableText`). Every glyph
 * cluster is its own absolutely positioned run, so the browser's own word
 * selection and copied text would break at every run: copy, double/triple
 * click and select-all are rebuilt here from the mirror's runs and doc ranges.
 */

import { type RefObject, useEffect } from 'react';
import {
  buildMirrorPage,
  MIRROR_TEXT_PAGE_CLASS,
  type DisplayPage,
} from '@betteroffice/docx/layout/render';

const PAGE = '.layout-page-mirror';
const RUN = '.layout-run-text';
const BLOCK = '[data-block-id]';
const CELL = '[role="cell"], [role="columnheader"]';
/**
 * Footnotes and endnotes are stories of their own, as in Word: a selection
 * highlights and copies only the story it began in (its anchor).
 */
export type TextStory = 'body' | 'footnotes' | 'endnotes';

interface Piece {
  run: HTMLElement;
  text: string;
  /** The selection starts inside this run. */
  cut: boolean;
  tabsBefore: number;
  breaksBefore: number;
  tabsAfter: number;
  breaksAfter: number;
  listMarker: boolean;
  blockId: string | null;
  cell: HTMLElement | null;
}

/**
 * Plain text of `range` over the pages under `root`, as the editor copies it:
 * paragraphs and line breaks as newlines, tabs as tabs, fields as their shown
 * text, list numbers and bullets as shown followed by a tab; table cells are
 * tab-separated and rows end in a newline. Tabs and line breaks come from the
 * runs' `tabsBefore`, `breaksBefore`, `tabsAfter` and `breaksAfter`, not from
 * gaps in their doc positions, which hidden text and content-control edges
 * leave too. A page away from the viewport holds only its plain-text mirror,
 * so its runs are read from a detached positioned mirror at the same text
 * offsets.
 */
export function textLayerText(
  range: Range,
  root: HTMLElement,
  pages: readonly DisplayPage[],
  /** The story copied: the one the selection began in (its anchor). */
  story: TextStory = storyOf(range.startContainer)
): string {
  const pieces: Piece[] = [];
  for (const pageEl of root.querySelectorAll<HTMLElement>(PAGE)) {
    if (!range.intersectsNode(pageEl)) continue;
    const local = pageEl.ownerDocument.createRange();
    local.selectNodeContents(pageEl);
    const whole =
      local.compareBoundaryPoints(Range.START_TO_START, range) >= 0 &&
      local.compareBoundaryPoints(Range.END_TO_END, range) <= 0;
    if (!whole) {
      if (local.compareBoundaryPoints(Range.START_TO_START, range) < 0) {
        local.setStart(range.startContainer, range.startOffset);
      }
      if (local.compareBoundaryPoints(Range.END_TO_END, range) > 0) {
        local.setEnd(range.endContainer, range.endOffset);
      }
    }
    if (!pageEl.classList.contains(MIRROR_TEXT_PAGE_CLASS)) {
      collect(pieces, pageEl, whole ? null : local, story);
      continue;
    }
    const page = pages.find((page) => String(page.pageIndex) === pageEl.dataset.pageIndex);
    if (!page) throw new Error(`No display page ${pageEl.dataset.pageIndex} for the text layer`);
    const full = buildMirrorPage(page, { document: pageEl.ownerDocument });
    const fullRange = whole
      ? null
      : rangeAtOffsets(
          full,
          textOffset(pageEl, local.startContainer, local.startOffset),
          textOffset(pageEl, local.endContainer, local.endOffset)
        );
    collect(pieces, full, fullRange, story);
  }
  return joined(pieces);
}

function joined(pieces: Piece[]): string {
  let text = '';
  pieces.forEach((piece, index) => {
    if (index > 0) text += separator(pieces[index - 1], piece);
    else if (!piece.cut) text += '\t'.repeat(piece.tabsBefore);
    text += piece.text;
  });
  return text;
}

/** The story's runs under `el` in reading order, with their paragraph and table cell. */
function collect(
  pieces: Piece[],
  el: Element,
  range: Range | null,
  story: TextStory,
  block: HTMLElement | null = null,
  cell: HTMLElement | null = null,
  current: TextStory = 'body'
): void {
  for (let child = el.firstElementChild; child; child = child.nextElementSibling) {
    if (!(child instanceof HTMLElement) || leftOut(child)) continue;
    const childStory = noteStory(child) ?? current;
    // other notes are skipped whole; body containers hold the notes
    if (childStory !== 'body' && childStory !== story) continue;
    if (child.classList.contains('layout-run-text')) {
      if (childStory !== story || (range && !range.intersectsNode(child))) continue;
      const node = child.firstChild;
      const listMarker = child.classList.contains('layout-list-marker');
      const text = selectedText(child, range).replace(/\u00ad/g, '');
      pieces.push({
        run: child,
        text: listMarker ? markerText(text, child.style.fontFamily) : text,
        cut: !!range && (range.startContainer === child || range.startContainer === node),
        tabsBefore: Number(child.getAttribute('data-tabs-before') ?? 0),
        breaksBefore: Number(child.getAttribute('data-breaks-before') ?? 0),
        tabsAfter: Number(child.getAttribute('data-tabs-after') ?? 0),
        breaksAfter: Number(child.getAttribute('data-breaks-after') ?? 0),
        listMarker,
        blockId: block?.getAttribute('data-block-id') ?? null,
        cell,
      });
      continue;
    }
    const role = child.getAttribute('role');
    collect(
      pieces,
      child,
      range,
      story,
      child.hasAttribute('data-block-id') ? child : block,
      role === 'cell' || role === 'columnheader' ? child : cell,
      childStory
    );
  }
}

/**
 * Unicode for the private-use (U+F0xx) characters Symbol and Wingdings list
 * markers are drawn with, after the fonts' published Unicode tables; any other
 * symbol-font marker copies as a bullet.
 */
const SYMBOL_FONT_MARKERS: Record<string, Record<number, string>> = {
  symbol: { 0xb7: '\u2022' },
  wingdings: {
    0x6e: '\u25a0', // ■
    0x71: '\u2751', // ❑
    0x76: '\u2756', // ❖
    0xa7: '\u25aa', // ▪
    0xd8: '\u27a2', // ➢
    0xe0: '\u2192', // →
    0xe8: '\u2794', // ➔
    0xfb: '\u2717', // ✗
    0xfc: '\u2713', // ✓
    0xfe: '\u2611', // ☑
  },
};

/** A list marker as plain text: symbol-font characters become their Unicode. */
export function markerText(text: string, fontFamily: string): string {
  const family = /wingdings/i.test(fontFamily)
    ? 'wingdings'
    : /symbol/i.test(fontFamily)
      ? 'symbol'
      : '';
  return Array.from(text, (char) => {
    const code = char.codePointAt(0) ?? 0;
    const symbolic = code >= 0xf000 && code <= 0xf0ff;
    if (!symbolic && !family) return char;
    if (!symbolic && code > 0xff) return char;
    return SYMBOL_FONT_MARKERS[family]?.[code & 0xff] ?? '\u2022';
  }).join('');
}

/** Not selectable on the text layer (see editor.css), so never copied. */
function leftOut(el: HTMLElement): boolean {
  const { classList } = el;
  if (el.getAttribute('aria-hidden') === 'true') return !classList.contains('layout-list-marker');
  return (
    classList.contains('layout-page-header') ||
    classList.contains('layout-page-footer') ||
    classList.contains('layout-run-chart') ||
    el.getAttribute('data-repeated-header') === 'true' ||
    // a tab's leader dots: the tab itself is the next run's tabsBefore
    el.hasAttribute('data-tab-leader')
  );
}

/** The note story `el` starts, if it is a note or a page's note area. */
function noteStory(el: Element): TextStory | null {
  if (el.classList.contains('layout-note')) {
    return el.getAttribute('role') === 'doc-endnote' ? 'endnotes' : 'footnotes';
  }
  if (el.classList.contains('layout-page-notes')) {
    return el.getAttribute('data-note-kind') === 'endnote' ? 'endnotes' : 'footnotes';
  }
  return null;
}

export function storyOf(node: Node | null): TextStory {
  const el = node instanceof Element ? node : node?.parentElement;
  const note = el?.closest('.layout-note, .layout-page-notes');
  return (note && noteStory(note)) ?? 'body';
}

function selectedText(run: HTMLElement, range: Range | null): string {
  const node = run.firstChild;
  if (!(node instanceof Text)) return '';
  if (!range) return node.data;
  if (range.startContainer === run && range.startOffset > 0) return '';
  if (range.endContainer === run && range.endOffset === 0) return '';
  const start = range.startContainer === node ? range.startOffset : 0;
  const end = range.endContainer === node ? range.endOffset : node.data.length;
  return node.data.slice(start, end);
}

/**
 * Newlines between two blocks: one between paragraphs, and one for each break
 * between two parts of a paragraph a page or column break splits, whose ids
 * (`id#n`) count the blocks before them.
 */
function paragraphBreaks(previous: string | null, next: string | null): number {
  const part = (id: string | null): [string | null, number] => {
    const match = id?.match(/^(.*)#(\d+)$/);
    return match ? [match[1]!, Number(match[2])] : [id, 0];
  };
  const [[from, at], [to, until]] = [part(previous), part(next)];
  return from === to && until > at ? until - at : 1;
}

function separator(previous: Piece, next: Piece): string {
  const tabs = '\t'.repeat(next.tabsBefore);
  // what ends the previous run's line: its trailing tabs, then its breaks
  const breaks = '\t'.repeat(previous.tabsAfter) + '\n'.repeat(previous.breaksAfter);
  if (previous.cell !== next.cell) {
    const sameRow =
      previous.cell && next.cell && previous.cell.parentElement === next.cell.parentElement;
    return breaks + (sameRow ? '\t' : '\n') + tabs;
  }
  // A paragraph split across pages keeps its block id on both.
  if (previous.blockId !== next.blockId) {
    return `${breaks}${'\n'.repeat(paragraphBreaks(previous.blockId, next.blockId))}${tabs}`;
  }
  if (previous.listMarker) return `\t${tabs}`;
  return breaks + '\n'.repeat(next.breaksBefore) + tabs;
}

/** Length of the text before a boundary point, counted from `root`. */
function textOffset(root: Node, node: Node, offset: number): number {
  const range = root.ownerDocument!.createRange();
  range.setStart(root, 0);
  range.setEnd(node, offset);
  return range.toString().length;
}

function rangeAtOffsets(root: HTMLElement, start: number, end: number): Range {
  const range = root.ownerDocument.createRange();
  range.selectNodeContents(root);
  const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let seen = 0;
  let startSet = false;
  for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
    const length = node.data.length;
    if (!startSet && start <= seen + length) {
      range.setStart(node, start - seen);
      startSet = true;
    }
    if (end <= seen + length) {
      range.setEnd(node, end - seen);
      break;
    }
    seen += length;
  }
  return range;
}

/**
 * Selects the word (`'word'`) or the paragraph part on this page
 * (`'paragraph'`) around `offset` in `run`, as a double or triple click does
 * in running text.
 */
export function selectAround(run: HTMLElement, offset: number, unit: 'word' | 'paragraph'): void {
  const selection = run.ownerDocument.getSelection();
  const block = run.closest<HTMLElement>(BLOCK);
  if (!selection || !block) return;
  const story = storyOf(block);
  const pieces: Piece[] = [];
  collect(pieces, block, null, story, block, block.closest<HTMLElement>(CELL), story);
  // The paragraph's text with its tabs and line breaks, so words end there.
  const runs: Array<{ node: Text; at: number }> = [];
  let text = '';
  pieces.forEach((piece, index) => {
    if (index > 0) text += separator(pieces[index - 1], piece);
    const node = piece.run.firstChild;
    if (!(node instanceof Text)) return;
    runs.push({ node, at: text.length });
    text += node.data;
  });
  if (runs.length === 0) return;
  let start = 0;
  let end = text.length;
  if (unit === 'word') {
    const at = (runs.find((entry) => entry.node.parentNode === run)?.at ?? 0) + offset;
    const lang = block.closest<HTMLElement>('[lang]')?.lang || undefined;
    const segment = new Intl.Segmenter(lang, { granularity: 'word' })
      .segment(text)
      .containing(Math.min(at, text.length - 1));
    if (!segment) return;
    start = segment.index;
    end = segment.index + segment.segment.length;
  }
  const first = [...runs].reverse().find((entry) => entry.at <= start) ?? runs[0];
  const last = [...runs].reverse().find((entry) => entry.at < end) ?? runs[0];
  selection.setBaseAndExtent(
    first.node,
    Math.min(start - first.at, first.node.data.length),
    last.node,
    Math.min(end - last.at, last.node.data.length)
  );
}

/** Character offset under a point inside a run, where the browser can tell. */
function offsetAtPoint(run: HTMLElement, x: number, y: number): number {
  const doc = run.ownerDocument as Document & {
    caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
  };
  const position = doc.caretPositionFromPoint?.(x, y);
  if (position) return position.offsetNode.parentNode === run ? position.offset : 0;
  const range = doc.caretRangeFromPoint?.(x, y);
  return range?.startContainer.parentNode === run ? range.startOffset : 0;
}


/**
 * Select all in the viewer, as Word does: the pages' body, or every footnote
 * (endnote) when the selection is in one.
 */
export function selectStory(host: HTMLElement): void {
  const selection = host.ownerDocument.getSelection();
  if (!selection) return;
  const anchor = selection.anchorNode;
  const story = anchor && host.contains(anchor) ? storyOf(anchor) : 'body';
  if (story === 'body') {
    selection.selectAllChildren(host);
    return;
  }
  const role = story === 'endnotes' ? 'doc-endnote' : 'doc-footnote';
  const notes = host.querySelectorAll(`.layout-note[role="${role}"]`);
  const first = notes[0];
  const last = notes[notes.length - 1];
  if (first && last) selection.setBaseAndExtent(first, 0, last, last.childNodes.length);
}

/**
 * Copy, select-all and double/triple-click selection for the text layer of
 * the pages under `hostRef`. Select-all acts once the last pointer press
 * landed in the viewer (in a frame the viewer fills, the body has focus).
 */
export function useTextLayer(
  hostRef: RefObject<HTMLElement | null>,
  pages: readonly DisplayPage[]
): void {
  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const doc = host.ownerDocument;
    let active = false;
    const onDocumentDown = (event: MouseEvent) => {
      active = host.contains(event.target as Node);
    };
    const runAt = (event: MouseEvent): HTMLElement | null => {
      if (event.button !== 0 || event.detail < 2) return null;
      const run = (event.target as Element).closest?.<HTMLElement>(RUN);
      return run && !leftOut(run) ? run : null;
    };
    const onDown = (event: MouseEvent) => {
      const pressed = (event.target as Element).closest?.(RUN);
      if (event.button === 0 && !event.shiftKey) {
        // A press on empty page area (the canvas, which starts no selection)
        // clears the selection, as a PDF viewer does.
        if (!pressed) doc.getSelection()?.removeAllRanges();
        // A press on text of the other story starts afresh there (editor.css
        // leaves the story a selection did not start in unselectable
        // meanwhile).
        const story = host.getAttribute('data-text-selection');
        if (pressed && story && story !== storyOf(pressed)) {
          host.removeAttribute('data-text-selection');
        }
      }
      // The browser would select one run; the word or paragraph is selected
      // on click, after the browser's own release handling (which collapses a
      // selection clicked inside).
      if (runAt(event)) event.preventDefault();
    };
    const onClick = (event: MouseEvent) => {
      // A left click on a mirror link does nothing, as on the canvas it covers.
      if ((event.target as Element).closest?.(`${PAGE} a`)) event.preventDefault();
      const run = runAt(event);
      if (!run) return;
      const offset = offsetAtPoint(run, event.clientX, event.clientY);
      selectAround(run, offset, event.detail === 2 ? 'word' : 'paragraph');
    };
    const onKey = (event: KeyboardEvent) => {
      if (!active || event.key.toLowerCase() !== 'a' || event.altKey || event.shiftKey) return;
      if (!(event.metaKey || event.ctrlKey)) return;
      const target = event.target as Node;
      if (target !== doc.body && !host.contains(target)) return;
      event.preventDefault();
      selectStory(host);
    };
    const onCopy = (event: ClipboardEvent) => {
      const selection = doc.getSelection();
      if (!selection || selection.isCollapsed || selection.rangeCount === 0) return;
      // Firefox splits a selection around unselectable content (the canvas).
      const range = selection.getRangeAt(0).cloneRange();
      const last = selection.getRangeAt(selection.rangeCount - 1);
      range.setEnd(last.endContainer, last.endOffset);
      if (!range.intersectsNode(host)) return;
      event.preventDefault();
      const story = storyOf(selection.anchorNode);
      event.clipboardData?.setData('text/plain', textLayerText(range, host, pages, story));
    };
    // The layer is painted only while it holds a selection, which keeps to
    // the story it started in (see editor.css).
    const onSelection = () => {
      const selection = doc.getSelection();
      const story =
        selection &&
        !selection.isCollapsed &&
        selection.rangeCount > 0 &&
        selection.getRangeAt(0).intersectsNode(host)
          ? storyOf(selection.anchorNode)
          : null;
      if (host.getAttribute('data-text-selection') === story) return;
      if (story) host.setAttribute('data-text-selection', story);
      else host.removeAttribute('data-text-selection');
    };
    doc.addEventListener('selectionchange', onSelection);
    doc.addEventListener('mousedown', onDocumentDown, true);
    host.addEventListener('mousedown', onDown);
    host.addEventListener('click', onClick);
    doc.addEventListener('keydown', onKey);
    doc.addEventListener('copy', onCopy);
    return () => {
      doc.removeEventListener('selectionchange', onSelection);
      doc.removeEventListener('mousedown', onDocumentDown, true);
      host.removeEventListener('mousedown', onDown);
      host.removeEventListener('click', onClick);
      doc.removeEventListener('keydown', onKey);
      doc.removeEventListener('copy', onCopy);
    };
  }, [hostRef, pages]);
}
