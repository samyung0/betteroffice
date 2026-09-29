import { useMemo } from 'react';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { TrackedChangesResult } from '@betteroffice/docx/layout/render';
import type { ReactSidebarItem } from '../plugin-api/types';
import { ReviewAllActions } from '../components/sidebar/ReviewAllActions';

export const REVIEW_ALL_ITEM_ID = 'review-all';

export interface ReviewAllItemsOptions {
  session: YrsSession | null | undefined;
  /** The listed changes; each new result also re-reads the changes fields keep. */
  trackedChanges: TrackedChangesResult;
  readOnly: boolean;
  /** Called once Accept all or Reject all resolved the session's changes. */
  onResolved: (session: YrsSession) => void;
}

/**
 * The sidebar row holding "Accept all" and "Reject all", pinned above the
 * change list while the document can change and holds changes: listed ones
 * or those fields keep in their code or result, which only these resolve.
 */
export function useReviewAllItems({
  session,
  trackedChanges,
  readOnly,
  onResolved,
}: ReviewAllItemsOptions): ReactSidebarItem[] {
  // eslint-disable-next-line react-hooks/exhaustive-deps -- a new result re-reads them
  const fieldChanges = useMemo(() => session?.hasFieldChanges() ?? false, [session, trackedChanges]);
  const shown = !readOnly && (trackedChanges.entries.length > 0 || fieldChanges);
  return useMemo(() => {
    if (!shown || !session) return [];
    const resolveAll = (accept: boolean): void => {
      if (accept) session.acceptChange({ all: true });
      else session.rejectChange({ all: true });
      onResolved(session);
    };
    return [
      {
        id: REVIEW_ALL_ITEM_ID,
        anchorPos: 0,
        fixedY: 0,
        priority: -1,
        estimatedHeight: 36,
        render: ({ measureRef }) => (
          <ReviewAllActions
            measureRef={measureRef}
            onAcceptAll={() => resolveAll(true)}
            onRejectAll={() => resolveAll(false)}
          />
        ),
      },
    ];
  }, [shown, session, onResolved]);
}
