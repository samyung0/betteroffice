import { useMemo } from 'react';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { TrackedChangesResult } from '@betteroffice/docx/layout/render';
import type { ReactSidebarItem } from '../plugin-api/types';
import { ReviewAllActions } from '../components/sidebar/ReviewAllActions';

export interface ReviewAllItemsOptions {
  /** The session Accept all or Reject all resolves, read on click. */
  getSession: () => YrsSession | null;
  /** The listed changes, and whether fields keep changes the list leaves out. */
  trackedChanges: TrackedChangesResult;
  readOnly: boolean;
  /** Called once Accept all or Reject all resolved the session's changes. */
  onResolved: (session: YrsSession) => void;
}

/**
 * The sidebar row holding "Accept all" and "Reject all", pinned above the
 * change list while the document can change and holds changes: listed ones
 * or those fields keep in their code or result, which only these resolve.
 * With only the latter, a line says the changes are inside fields.
 */
export function useReviewAllItems({
  getSession,
  trackedChanges,
  readOnly,
  onResolved,
}: ReviewAllItemsOptions): ReactSidebarItem[] {
  const listed = trackedChanges.entries.length > 0;
  const { fieldChanges } = trackedChanges;
  const shown = !readOnly && (listed || fieldChanges);
  const fieldsOnly = !listed && fieldChanges;
  return useMemo(() => {
    if (!shown) return [];
    const resolveAll = (accept: boolean): void => {
      const session = getSession();
      if (!session) return;
      // Its own undo step, apart from typing just before it.
      session.addUndoBoundary();
      if (accept) session.acceptChange({ all: true });
      else session.rejectChange({ all: true });
      onResolved(session);
    };
    return [
      {
        id: 'review-all',
        anchorPos: 0,
        fixedY: 0,
        priority: -1,
        estimatedHeight: fieldsOnly ? 56 : 36,
        render: ({ measureRef }) => (
          <ReviewAllActions
            measureRef={measureRef}
            fieldsOnly={fieldsOnly}
            onAcceptAll={() => resolveAll(true)}
            onRejectAll={() => resolveAll(false)}
          />
        ),
      },
    ];
  }, [shown, fieldsOnly, getSession, onResolved]);
}
