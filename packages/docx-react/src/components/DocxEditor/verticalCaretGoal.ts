import type { DisplayListQueries, DisplayListVisualLine } from '@betteroffice/docx/layout/render';

export class VerticalCaretGoal {
  private goalX: number | undefined;

  current(): number | undefined {
    return this.goalX;
  }

  retain(goalX: number): void {
    this.goalX = goalX;
  }

  reset(): void {
    this.goalX = undefined;
  }
}

export interface VerticalParagraph {
  paraId: string;
  length: number;
}

export interface VerticalParagraphLoc {
  story: string;
  paraId: string;
  offset: number;
}

export function paragraphVerticalMove(
  paragraphs: readonly VerticalParagraph[],
  head: VerticalParagraphLoc,
  direction: 'up' | 'down'
): VerticalParagraphLoc {
  const index = paragraphs.findIndex((paragraph) => paragraph.paraId === head.paraId);
  if (index < 0) return head;
  const target = paragraphs[index + (direction === 'up' ? -1 : 1)];
  return target
    ? {
        story: head.story,
        paraId: target.paraId,
        offset: Math.min(head.offset, target.length),
      }
    : head;
}

export interface PageMove {
  position: number;
  goalX: number;
  /** Client px between the caret's line and the target line, for the view to follow. */
  scrollBy: number;
}

/**
 * Where PageUp/PageDown from body display `position` lands (#1538): the visual line whose centre
 * lies nearest one viewport (`distance`, client px) above or below the caret line's centre, hit
 * at `goalX`. With no line further that way the caret goes to the start or end of its line.
 * `clientY` projects a page-local y to client px, or null for a page with no geometry.
 */
export function pageMoveTarget(
  queries: Pick<
    DisplayListQueries,
    'visualLines' | 'visualLineAtPosition' | 'caretRect' | 'hitTestRegions' | 'pageSize'
  >,
  clientY: (pageIndex: number, y: number) => number | null,
  position: number,
  direction: 'up' | 'down',
  distance: number,
  goalX: number | undefined
): PageMove | null {
  const line = queries.visualLineAtPosition(position);
  const centre = (candidate: DisplayListVisualLine) =>
    clientY(candidate.pageIndex, candidate.y + candidate.height / 2);
  const from = line ? centre(line) : null;
  if (!line || from === null) return null;
  const x = goalX ?? queries.caretRect(position)?.x ?? line.x;
  const want = direction === 'down' ? from + distance : from - distance;
  let target: DisplayListVisualLine | null = null;
  let targetCentre = from;
  for (const candidate of queries.visualLines()) {
    const y = centre(candidate);
    if (y === null || (direction === 'down' ? y <= from : y >= from)) continue;
    if (!target || Math.abs(y - want) < Math.abs(targetCentre - want)) {
      target = candidate;
      targetCentre = y;
    }
  }
  const edgeX =
    direction === 'down' ? (queries.pageSize(line.pageIndex)?.width ?? line.x + line.width) : 0;
  const hitLine = target ?? line;
  const hit = queries.hitTestRegions(
    hitLine.pageIndex,
    target ? x : edgeX,
    hitLine.y + hitLine.height / 2
  );
  if (hit?.region !== 'body' || hit.pos === null) return null;
  return { position: hit.pos, goalX: x, scrollBy: targetCentre - from };
}
