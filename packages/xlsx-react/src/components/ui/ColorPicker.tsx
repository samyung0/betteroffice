import { useEffect, useRef, useState } from 'react';
import { useTranslation } from '../../i18n';
import { ToolbarIcon } from './ToolbarIcon';
import { ToolbarDropdown, menuColors, menuFontFamily } from './ToolbarPrimitives';

export interface PaletteColor {
  name: string;
  value: string;
}

/** Eight columns: neutrals, then dark, strong, soft and pale hues. */
export const PALETTE: readonly PaletteColor[] = [
  { name: 'Black', value: '#000000' },
  { name: 'Charcoal', value: '#434343' },
  { name: 'Dark gray', value: '#666666' },
  { name: 'Gray', value: '#999999' },
  { name: 'Silver', value: '#b7b7b7' },
  { name: 'Light gray', value: '#cccccc' },
  { name: 'Cloud', value: '#eeeeee' },
  { name: 'White', value: '#ffffff' },
  { name: 'Dark red', value: '#991b1b' },
  { name: 'Dark orange', value: '#9a3412' },
  { name: 'Dark yellow', value: '#a16207' },
  { name: 'Dark green', value: '#3f6212' },
  { name: 'Dark teal', value: '#0f766e' },
  { name: 'Dark blue', value: '#1d4ed8' },
  { name: 'Dark purple', value: '#6d28d9' },
  { name: 'Dark pink', value: '#9d174d' },
  { name: 'Red', value: '#dc2626' },
  { name: 'Orange', value: '#ea580c' },
  { name: 'Yellow', value: '#ca8a04' },
  { name: 'Green', value: '#65a30d' },
  { name: 'Teal', value: '#0d9488' },
  { name: 'Blue', value: '#2563eb' },
  { name: 'Purple', value: '#7c3aed' },
  { name: 'Pink', value: '#db2777' },
  { name: 'Soft red', value: '#fca5a5' },
  { name: 'Soft orange', value: '#fdba74' },
  { name: 'Soft yellow', value: '#fde047' },
  { name: 'Soft green', value: '#bef264' },
  { name: 'Soft teal', value: '#5eead4' },
  { name: 'Soft blue', value: '#93c5fd' },
  { name: 'Soft purple', value: '#c4b5fd' },
  { name: 'Soft pink', value: '#f9a8d4' },
  { name: 'Pale red', value: '#fee2e2' },
  { name: 'Pale orange', value: '#ffedd5' },
  { name: 'Pale yellow', value: '#fef9c3' },
  { name: 'Pale green', value: '#ecfccb' },
  { name: 'Pale teal', value: '#ccfbf1' },
  { name: 'Pale blue', value: '#dbeafe' },
  { name: 'Pale purple', value: '#ede9fe' },
  { name: 'Pale pink', value: '#fce7f3' },
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
  const [draft, setDraft] = useState(HEX.test(value) ? current : '');
  const [clearHovered, setClearHovered] = useState(false);
  const pickRef = useRef(onPick);
  pickRef.current = onPick;

  // The native picker fires `input` while dragging; only its `change` commits.
  useEffect(() => {
    const input = customRef.current;
    if (!input) return;
    const update = () => setDraft(input.value);
    const commit = () => pickRef.current(input.value);
    input.addEventListener('input', update);
    input.addEventListener('change', commit);
    return () => {
      input.removeEventListener('input', update);
      input.removeEventListener('change', commit);
    };
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
          return (
            <button
              key={color.value}
              type="button"
              role="gridcell"
              aria-label={color.name}
              aria-selected={selected}
              title={`${color.name} (${color.value})`}
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
        <span
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: 8,
            color: menuColors.text,
            font: '400 12px ui-monospace, SFMono-Regular, Menlo, monospace',
          }}
        >
          {draft}
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
        </span>
      </label>
    </div>
  );
}
