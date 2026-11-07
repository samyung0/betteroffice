import { useEffect, useRef, useState } from 'react';
import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import { useTranslation } from '../../i18n';
import { ToolbarIcon } from './ToolbarIcon';
import { ToolbarDropdown, menuColors, menuFontFamily } from './ToolbarPrimitives';

export interface PaletteColor {
  nameKey: TranslationKey;
  value: string;
}

/** Eight columns: neutrals, then dark, strong, soft and pale hues. */
export const PALETTE: readonly PaletteColor[] = [
  { nameKey: 'toolbar.palette.colors.black', value: '#000000' },
  { nameKey: 'toolbar.palette.colors.charcoal', value: '#434343' },
  { nameKey: 'toolbar.palette.colors.darkGray', value: '#666666' },
  { nameKey: 'toolbar.palette.colors.gray', value: '#999999' },
  { nameKey: 'toolbar.palette.colors.silver', value: '#b7b7b7' },
  { nameKey: 'toolbar.palette.colors.lightGray', value: '#cccccc' },
  { nameKey: 'toolbar.palette.colors.cloud', value: '#eeeeee' },
  { nameKey: 'toolbar.palette.colors.white', value: '#ffffff' },
  { nameKey: 'toolbar.palette.colors.darkRed', value: '#991b1b' },
  { nameKey: 'toolbar.palette.colors.darkOrange', value: '#9a3412' },
  { nameKey: 'toolbar.palette.colors.darkYellow', value: '#a16207' },
  { nameKey: 'toolbar.palette.colors.darkGreen', value: '#3f6212' },
  { nameKey: 'toolbar.palette.colors.darkTeal', value: '#0f766e' },
  { nameKey: 'toolbar.palette.colors.darkBlue', value: '#1d4ed8' },
  { nameKey: 'toolbar.palette.colors.darkPurple', value: '#6d28d9' },
  { nameKey: 'toolbar.palette.colors.darkPink', value: '#9d174d' },
  { nameKey: 'toolbar.palette.colors.red', value: '#dc2626' },
  { nameKey: 'toolbar.palette.colors.orange', value: '#ea580c' },
  { nameKey: 'toolbar.palette.colors.yellow', value: '#ca8a04' },
  { nameKey: 'toolbar.palette.colors.green', value: '#65a30d' },
  { nameKey: 'toolbar.palette.colors.teal', value: '#0d9488' },
  { nameKey: 'toolbar.palette.colors.blue', value: '#2563eb' },
  { nameKey: 'toolbar.palette.colors.purple', value: '#7c3aed' },
  { nameKey: 'toolbar.palette.colors.pink', value: '#db2777' },
  { nameKey: 'toolbar.palette.colors.softRed', value: '#fca5a5' },
  { nameKey: 'toolbar.palette.colors.softOrange', value: '#fdba74' },
  { nameKey: 'toolbar.palette.colors.softYellow', value: '#fde047' },
  { nameKey: 'toolbar.palette.colors.softGreen', value: '#bef264' },
  { nameKey: 'toolbar.palette.colors.softTeal', value: '#5eead4' },
  { nameKey: 'toolbar.palette.colors.softBlue', value: '#93c5fd' },
  { nameKey: 'toolbar.palette.colors.softPurple', value: '#c4b5fd' },
  { nameKey: 'toolbar.palette.colors.softPink', value: '#f9a8d4' },
  { nameKey: 'toolbar.palette.colors.paleRed', value: '#fee2e2' },
  { nameKey: 'toolbar.palette.colors.paleOrange', value: '#ffedd5' },
  { nameKey: 'toolbar.palette.colors.paleYellow', value: '#fef9c3' },
  { nameKey: 'toolbar.palette.colors.paleGreen', value: '#ecfccb' },
  { nameKey: 'toolbar.palette.colors.paleTeal', value: '#ccfbf1' },
  { nameKey: 'toolbar.palette.colors.paleBlue', value: '#dbeafe' },
  { nameKey: 'toolbar.palette.colors.palePurple', value: '#ede9fe' },
  { nameKey: 'toolbar.palette.colors.palePink', value: '#fce7f3' },
];

export interface ColorPickerProps {
  mode: 'text' | 'fill';
  value: string;
  label: string;
  onChange?: (value: string) => void;
  /** Removes the colour, so the cell takes the default again. */
  onClear?: () => void;
  disabled?: boolean;
}

/** A colour button whose dropdown holds the palette, a reset and a custom colour. */
export function ColorPicker({
  mode,
  value,
  label,
  onChange,
  onClear,
  disabled = false,
}: ColorPickerProps) {
  return (
    <ToolbarDropdown
      title={label}
      disabled={disabled}
      menuWidth={256}
      style={{ position: 'relative' }}
      trigger={
        <>
          <ToolbarIcon name={mode === 'text' ? 'textColor' : 'fillColor'} size={20} />
          <span
            aria-hidden="true"
            data-color-bar
            style={{
              position: 'absolute',
              left: 5,
              right: 5,
              bottom: 2,
              height: 3,
              borderRadius: 2,
              background: value,
            }}
          />
        </>
      }
    >
      {(close) => (
        <ColorPalette
          value={value}
          onPick={(color) => {
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

const HEX = /^#[\da-f]{6}$/i;

function isBright(hex: string) {
  const [r, g, b] = [1, 3, 5].map((index) => Number.parseInt(hex.slice(index, index + 2), 16));
  return r * 0.299 + g * 0.587 + b * 0.114 > 160;
}

/** The palette's captions: its heading, Default and Custom color. */
const captionStyle = {
  color: menuColors.secondary,
  font: `600 var(--office-menu-label-font-size, 12px) / 16px ${menuFontFamily}`,
} as const;

export function ColorPalette({
  value,
  onPick,
  onClear,
  heading = true,
}: {
  value: string;
  onPick: (color: string) => void;
  onClear?: () => void;
  /** False leaves out the "Document colors" heading, for a menu that names its section. */
  heading?: boolean;
}) {
  const { t } = useTranslation();
  const current = value.toLowerCase();
  const customRef = useRef<HTMLInputElement>(null);
  const [clearHovered, setClearHovered] = useState(false);
  const pickRef = useRef(onPick);
  pickRef.current = onPick;

  // The native picker fires `input` while dragging; only its `change` commits.
  useEffect(() => {
    const input = customRef.current;
    if (!input) return;
    const commit = () => pickRef.current(input.value);
    input.addEventListener('change', commit);
    return () => input.removeEventListener('change', commit);
  }, []);

  return (
    <div
      style={{
        display: 'flex',
        flexDirection: 'column',
        gap: 8,
        padding: '0 6px',
      }}
    >
      {heading && (
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <span style={captionStyle}>{t('toolbar.palette.documentColors')}</span>
          {onClear && (
            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onMouseEnter={() => setClearHovered(true)}
              onMouseLeave={() => setClearHovered(false)}
              onClick={onClear}
              style={{
                appearance: 'none',
                display: 'inline-flex',
                alignItems: 'center',
                gap: 4,
                height: 28,
                padding: '0 8px',
                border: 0,
                borderRadius: 'var(--office-menu-button-radius, var(--xlsx-button-radius, 4px))',
                cursor: 'pointer',
                ...captionStyle,
                fontWeight: 500,
                background: clearHovered ? menuColors.buttonHover : 'transparent',
                color: clearHovered ? menuColors.text : menuColors.secondary,
              }}
            >
              <ToolbarIcon name="refresh" size={14} />
              {t('toolbar.palette.default')}
            </button>
          )}
        </div>
      )}
      <div
        role="grid"
        aria-label={t('toolbar.palette.documentColors')}
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(8, 24px)',
          gap: 6,
        }}
      >
        {PALETTE.map((color) => {
          const selected = current === color.value;
          const name = t(color.nameKey);
          return (
            <button
              key={color.value}
              type="button"
              role="gridcell"
              aria-label={name}
              aria-selected={selected}
              title={`${name} (${color.value})`}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => onPick(color.value)}
              style={{
                appearance: 'none',
                display: 'grid',
                placeItems: 'center',
                width: 24,
                height: 24,
                padding: 0,
                border: `1px solid ${menuColors.swatchBorder}`,
                borderRadius: '50%',
                background: color.value,
                color: isBright(color.value) ? '#000000' : '#ffffff',
                cursor: 'pointer',
                boxShadow: selected
                  ? `0 0 0 1px ${menuColors.bg}, 0 0 0 3px ${menuColors.ring}`
                  : undefined,
              }}
            >
              {selected && <ToolbarIcon name="check" size={14} />}
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
          marginTop: 8,
          paddingTop: 8,
          borderTop: `1px solid ${menuColors.separator}`,
          ...captionStyle,
        }}
      >
        {t('toolbar.palette.custom')}
        <input
          ref={customRef}
          type="color"
          defaultValue={HEX.test(value) ? current : '#000000'}
          aria-label={t('toolbar.palette.chooseCustom')}
          style={{
            width: 36,
            height: 28,
            padding: 2,
            border: `1px solid ${menuColors.border}`,
            borderRadius: 'var(--office-menu-button-radius, var(--xlsx-button-radius, 4px))',
            background: 'transparent',
            cursor: 'pointer',
          }}
        />
      </label>
    </div>
  );
}
