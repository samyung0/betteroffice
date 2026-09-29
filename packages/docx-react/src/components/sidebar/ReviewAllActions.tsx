import type { CSSProperties } from 'react';
import { MaterialSymbol } from '../ui/Icons';
import { ICON_BUTTON_STYLE } from './cardUtils';
import { useTranslation } from '../../i18n';

const BUTTON_STYLE: CSSProperties = {
  ...ICON_BUTTON_STYLE,
  alignItems: 'center',
  gap: 4,
  borderRadius: 16,
  padding: '4px 10px 4px 6px',
  fontSize: 13,
  whiteSpace: 'nowrap',
};

export interface ReviewAllActionsProps {
  onAcceptAll: () => void;
  onRejectAll: () => void;
  /** Whether every change sits inside a field, where the page cannot mark it. */
  fieldsOnly?: boolean;
  measureRef?: (el: HTMLDivElement | null) => void;
}

/** "Accept all" and "Reject all" above the change list, styled as a change card's buttons. */
export function ReviewAllActions({
  onAcceptAll,
  onRejectAll,
  fieldsOnly = false,
  measureRef,
}: ReviewAllActionsProps) {
  const { t } = useTranslation();
  return (
    <div
      ref={measureRef}
      className="docx-review-all-actions"
      // The editor keeps focus, so Ctrl+Z still undoes the resolve.
      onMouseDown={(e) => {
        e.preventDefault();
        e.stopPropagation();
      }}
    >
      {fieldsOnly && (
        <div
          style={{
            fontSize: 12,
            lineHeight: '16px',
            color: 'var(--doc-text-muted)',
            textAlign: 'right',
            marginBottom: 4,
          }}
        >
          {t('trackedChanges.inFieldsOnly')}
        </div>
      )}
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 4 }}>
        <button type="button" onClick={onAcceptAll} style={BUTTON_STYLE}>
          <MaterialSymbol name="check" size={20} />
          {t('trackedChanges.acceptAll')}
        </button>
        <button type="button" onClick={onRejectAll} style={BUTTON_STYLE}>
          <MaterialSymbol name="close" size={20} />
          {t('trackedChanges.rejectAll')}
        </button>
      </div>
    </div>
  );
}
