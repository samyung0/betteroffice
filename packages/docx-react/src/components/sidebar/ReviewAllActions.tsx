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
};

export interface ReviewAllActionsProps {
  onAcceptAll: () => void;
  onRejectAll: () => void;
  measureRef?: (el: HTMLDivElement | null) => void;
}

/** "Accept all" and "Reject all" above the change list, styled as a change card's buttons. */
export function ReviewAllActions({ onAcceptAll, onRejectAll, measureRef }: ReviewAllActionsProps) {
  const { t } = useTranslation();
  return (
    <div
      ref={measureRef}
      className="docx-review-all-actions"
      onMouseDown={(e) => e.stopPropagation()}
      style={{ display: 'flex', justifyContent: 'flex-end', gap: 4 }}
    >
      <button type="button" onClick={onAcceptAll} style={BUTTON_STYLE}>
        <MaterialSymbol name="check" size={20} />
        {t('trackedChanges.acceptAll')}
      </button>
      <button type="button" onClick={onRejectAll} style={BUTTON_STYLE}>
        <MaterialSymbol name="close" size={20} />
        {t('trackedChanges.rejectAll')}
      </button>
    </div>
  );
}
