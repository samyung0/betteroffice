/**
 * TableBorderColorPicker - Wrapper around ColorPicker for table border colors.
 *
 * Translates ColorPicker's ColorValue output to the TableAction format
 * expected by the toolbar's table action handler.
 */

import { useCallback } from 'react';
import type { ColorValue } from '@betteroffice/docx/types/document';
import type { Theme } from '@betteroffice/docx/types/document';
import type { TableAction } from './TableToolbar';
import { ColorPicker, type ColorPaletteColor } from './ColorPicker';
import { useTranslation } from '../../i18n';

export interface TableBorderColorPickerProps {
  onAction: (action: TableAction) => void;
  disabled?: boolean;
  theme?: Theme | null;
  /** Current border color (RGB hex without #) */
  value?: string;
  /** The host's colours, in place of the theme's (`ColorPicker` `palette`). */
  palette?: readonly ColorPaletteColor[];
}

export function TableBorderColorPicker({
  onAction,
  disabled = false,
  theme,
  value,
  palette,
}: TableBorderColorPickerProps) {
  const { t } = useTranslation();
  const handleChange = useCallback(
    (color: ColorValue | string) => {
      if (typeof color === 'string') {
        onAction({ type: 'borderColor', color: color.replace(/^#/, '') });
      } else if (color.rgb) {
        onAction({ type: 'borderColor', color: color.rgb.replace(/^#/, '') });
      } else if (color.auto) {
        onAction({ type: 'borderColor', color: '000000' });
      }
    },
    [onAction]
  );

  return (
    <ColorPicker
      mode="border"
      value={value}
      onChange={handleChange}
      theme={theme}
      disabled={disabled}
      palette={palette}
      title={t('table.borderColor')}
    />
  );
}

export default TableBorderColorPicker;
