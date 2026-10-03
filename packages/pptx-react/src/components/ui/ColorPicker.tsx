import { useEffect, useRef } from 'react';
import { useTranslation } from '../../i18n';
import { ToolbarIcon } from './ToolbarIcon';
import type { ToolbarIconName } from './ToolbarIcon';
import { ToolbarDropdown, toolbarColors, toolbarFont } from './ToolbarPrimitives';

/** Eight hues by five shades, as Capy's document colours. */
const PALETTE = [
  '#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#eeeeee', '#ffffff',
  '#991b1b', '#9a3412', '#a16207', '#3f6212', '#0f766e', '#1d4ed8', '#6d28d9', '#9d174d',
  '#dc2626', '#ea580c', '#ca8a04', '#65a30d', '#0d9488', '#2563eb', '#7c3aed', '#db2777',
  '#fca5a5', '#fdba74', '#fde047', '#bef264', '#5eead4', '#93c5fd', '#c4b5fd', '#f9a8d4',
  '#fee2e2', '#ffedd5', '#fef9c3', '#ecfccb', '#ccfbf1', '#dbeafe', '#ede9fe', '#fce7f3',
] as const;

export interface ColorPickerProps {
  value: string;
  label: string;
  onChange?: (value: string) => void;
  onClear?: () => void;
  icon?: ToolbarIconName;
  none?: boolean;
  clearLabel?: string;
  disabled?: boolean;
  testId?: string;
}

/** A colour button whose popover holds the palette, a clear item and a custom colour. */
export function ColorPicker({
  value,
  label,
  onChange,
  onClear,
  icon = 'textColor',
  none = false,
  clearLabel,
  disabled = false,
  testId,
}: ColorPickerProps) {
  return (
    <ToolbarDropdown
      title={label}
      disabled={disabled}
      menuWidth={256}
      testId={testId}
      style={{ position: 'relative' }}
      trigger={
        <>
          <ToolbarIcon name={icon} size={20} />
          <span
            aria-hidden="true"
            style={{
              position: 'absolute',
              left: 6,
              right: 6,
              bottom: 2,
              height: 4,
              borderRadius: 999,
              background: none ? 'transparent' : value,
              border: `1px solid var(--pptx-swatch-border, ${toolbarColors.border})`,
              boxSizing: 'border-box',
            }}
          />
        </>
      }
    >
      {(close) => (
        <ColorPalette
          value={none ? undefined : value}
          label={label}
          clearLabel={clearLabel}
          testId={testId}
          onChange={(color) => {
            onChange?.(color);
            close();
          }}
          onClear={
            onClear &&
            (() => {
              onClear();
              close();
            })
          }
        />
      )}
    </ToolbarDropdown>
  );
}

function ColorPalette({
  value,
  label,
  clearLabel,
  testId,
  onChange,
  onClear,
}: {
  value?: string;
  label: string;
  clearLabel?: string;
  testId?: string;
  onChange: (color: string) => void;
  onClear?: () => void;
}) {
  const { t } = useTranslation();
  const custom = useRef<HTMLInputElement>(null);
  const current = value?.toLowerCase();
  // A colour input's input event fires while dragging; only its change commits.
  useEffect(() => {
    const input = custom.current;
    if (!input) return;
    const commit = () => onChange(input.value);
    input.addEventListener('change', commit);
    return () => input.removeEventListener('change', commit);
  }, [onChange]);
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, padding: '2px 6px 4px' }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <span style={{ color: toolbarColors.muted, font: `600 12px ${toolbarFont}` }}>{label}</span>
        {onClear && clearLabel ? (
          <button
            type="button"
            aria-label={clearLabel}
            onMouseDown={(event) => event.preventDefault()}
            onClick={onClear}
            style={{
              appearance: 'none',
              height: 28,
              padding: '0 8px',
              border: 0,
              borderRadius: 'var(--pptx-button-radius, 4px)',
              background: 'transparent',
              color: toolbarColors.muted,
              cursor: 'pointer',
              font: `500 12px ${toolbarFont}`,
            }}
          >
            {clearLabel}
          </button>
        ) : null}
      </div>
      <div role="grid" aria-label={label} style={{ display: 'grid', gridTemplateColumns: 'repeat(8, 24px)', gap: 6, justifyContent: 'space-between' }}>
        {PALETTE.map((color) => {
          const selected = current === color;
          return (
            <button
              key={color}
              type="button"
              role="gridcell"
              aria-label={color}
              aria-selected={selected}
              title={color}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => onChange(color)}
              style={{
                appearance: 'none',
                display: 'grid',
                placeItems: 'center',
                width: 24,
                height: 24,
                padding: 0,
                border: `1px solid var(--pptx-swatch-border, ${toolbarColors.border})`,
                borderRadius: 999,
                background: color,
                boxShadow: selected
                  ? `0 0 0 1px ${toolbarColors.surface}, 0 0 0 3px var(--pptx-accent, #1a73e8)`
                  : 'none',
                color: brightness(color) > 150 ? '#000000' : '#ffffff',
                cursor: 'pointer',
              }}
            >
              {selected ? <ToolbarIcon name="check" size={14} /> : null}
            </button>
          );
        })}
      </div>
      <label
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: 12,
          paddingTop: 8,
          borderTop: `1px solid ${toolbarColors.divider}`,
          color: toolbarColors.muted,
          font: `600 12px ${toolbarFont}`,
        }}
      >
        {t('toolbar.customColor')}
        <span style={{ display: 'flex', alignItems: 'center', gap: 8, color: toolbarColors.text, font: '400 12px ui-monospace, monospace' }}>
          {current ?? ''}
          <input
            ref={custom}
            data-testid={testId && `${testId}-custom`}
            type="color"
            aria-label={`${label}: ${t('toolbar.customColor')}`}
            defaultValue={current ?? '#000000'}
            style={{
              width: 36,
              height: 28,
              padding: 2,
              border: `1px solid ${toolbarColors.border}`,
              borderRadius: 'var(--pptx-button-radius, 4px)',
              background: 'transparent',
              cursor: 'pointer',
            }}
          />
        </span>
      </label>
    </div>
  );
}

function brightness(hex: string): number {
  const red = Number.parseInt(hex.slice(1, 3), 16);
  const green = Number.parseInt(hex.slice(3, 5), 16);
  const blue = Number.parseInt(hex.slice(5, 7), 16);
  return (red * 299 + green * 587 + blue * 114) / 1000;
}
