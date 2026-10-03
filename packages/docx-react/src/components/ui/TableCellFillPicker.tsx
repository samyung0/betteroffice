/**
 * TableCellFillPicker - Wrapper around ColorPicker for table cell fill/shading.
 *
 * Translates ColorPicker's output to the TableAction format
 * expected by the toolbar's table action handler.
 */

import { useCallback } from 'react';
import type { ColorValue, Theme } from '@betteroffice/docx/types/document';
import type { TableAction } from './TableToolbar';
import { ColorPicker, type ColorPaletteColor } from './ColorPicker';
import { useTranslation } from '../../i18n';

export interface TableCellFillPickerProps {
  onAction: (action: TableAction) => void;
  disabled?: boolean;
  theme?: Theme | null;
  /** Current fill color (RGB hex without #) */
  value?: string;
  /** The host's colours, in place of the theme's (`ColorPicker` `palette`). */
  palette?: readonly ColorPaletteColor[];
}

export function TableCellFillPicker({
  onAction,
  disabled = false,
  theme,
  value,
  palette,
}: TableCellFillPickerProps) {
  const { t } = useTranslation();
  const handleChange = useCallback(
    (color: ColorValue | string) => {
      // highlight mode emits hex strings or 'none'
      if (typeof color === 'string') {
        if (color === 'none') {
          onAction({ type: 'cellFillColor', color: null });
        } else {
          onAction({ type: 'cellFillColor', color: color.replace(/^#/, '') });
        }
      }
    },
    [onAction]
  );

  return (
    <ColorPicker
      mode="highlight"
      value={value}
      onChange={handleChange}
      theme={theme}
      disabled={disabled}
      palette={palette}
      title={t('table.cellFillColor')}
      icon="format_color_fill"
      autoLabel={t('colorPicker.noColor')}
    />
  );
}

export default TableCellFillPicker;
