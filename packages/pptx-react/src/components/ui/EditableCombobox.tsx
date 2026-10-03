import { useCallback, useEffect, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { ToolbarIcon } from './ToolbarIcon';
import { menuRowStyle, popoverStyle, toolbarColors, toolbarFont } from './ToolbarPrimitives';

export interface ComboboxOption {
  value: string;
  label: string;
}

export interface EditableComboboxProps {
  value: string;
  options: readonly ComboboxOption[];
  label: string;
  onCommit?: (value: string) => void;
  disabled?: boolean;
  width?: number;
  inputStyle?: CSSProperties;
  testId?: string;
  /** Without the chevron the box opens its list on focus only. */
  chevron?: boolean;
}

export function EditableCombobox({
  value,
  options,
  label,
  onCommit,
  disabled = false,
  width = 72,
  inputStyle,
  testId,
  chevron = true,
}: EditableComboboxProps) {
  const [draft, setDraft] = useState(value);
  const [open, setOpen] = useState(false);
  const [position, setPosition] = useState({ top: 0, left: 0, width });
  const rootRef = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => setDraft(value), [value]);

  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open || !rootRef.current) return;
    const rect = rootRef.current.getBoundingClientRect();
    setPosition({
      top: rect.bottom + 4,
      left: rect.left,
      width: Math.max(rect.width, width),
    });
  }, [open, width]);

  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: MouseEvent) => {
      const target = event.target as Node;
      if (!rootRef.current?.contains(target) && !menuRef.current?.contains(target)) close();
    };
    const onKeyDown = (event: globalThis.KeyboardEvent) => {
      if (event.key === 'Escape') {
        setDraft(value);
        close();
      }
    };
    document.addEventListener('mousedown', onPointerDown);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('mousedown', onPointerDown);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [open, close, value]);

  const commit = useCallback(() => {
    if (!disabled && onCommit && draft.trim()) onCommit(draft.trim());
    setDraft(value);
  }, [disabled, onCommit, draft, value]);

  return (
    <div
      ref={rootRef}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        width,
        height: 28,
        border: `1px solid ${toolbarColors.border}`,
        borderRadius: 'var(--pptx-input-radius, 4px)',
        background: toolbarColors.surface,
        opacity: disabled ? 0.48 : 1,
        boxSizing: 'border-box',
        flex: '0 0 auto',
      }}
    >
      <input
        ref={inputRef}
        data-testid={testId}
        role="combobox"
        aria-label={label}
        aria-autocomplete="list"
        aria-expanded={open}
        disabled={disabled}
        value={draft}
        onFocus={(event) => {
          if (disabled) return;
          setOpen(true);
          event.currentTarget.select();
        }}
        onChange={(event) => setDraft(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            commit();
            close();
            event.currentTarget.blur();
          } else if (event.key === 'ArrowDown') {
            setOpen(true);
            event.preventDefault();
          }
        }}
        onBlur={() => {
          requestAnimationFrame(() => {
            if (!menuRef.current?.contains(document.activeElement)) commit();
          });
        }}
        style={{
          appearance: 'none',
          width: chevron ? 'calc(100% - 22px)' : '100%',
          height: 26,
          padding: '0 2px 0 7px',
          border: 0,
          outline: 0,
          background: 'transparent',
          color: toolbarColors.text,
          font: `500 13px ${toolbarFont}`,
          boxSizing: 'border-box',
          ...inputStyle,
        }}
      />
      {chevron && (
        <button
          type="button"
          disabled={disabled}
          tabIndex={-1}
          aria-label={label}
          title={label}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            if (disabled) return;
            setOpen((current) => !current);
            inputRef.current?.focus();
          }}
          style={{
            appearance: 'none',
            display: 'grid',
            placeItems: 'center',
            width: 21,
            height: 26,
            padding: 0,
            border: 0,
            background: 'transparent',
            color: toolbarColors.muted,
            cursor: disabled ? 'default' : 'pointer',
          }}
        >
          <ToolbarIcon name="chevronDown" size={14} />
        </button>
      )}
      {open && !disabled && (
        <div
          ref={menuRef}
          role="listbox"
          aria-label={label}
          style={{
            position: 'fixed',
            top: position.top,
            left: position.left,
            zIndex: 10000,
            width: position.width,
            maxHeight: 260,
            overflowY: 'auto',
            ...popoverStyle,
          }}
        >
          {options.map((option) => (
            <ComboboxOption
              key={option.value}
              label={option.label}
              selected={option.value === value}
              centered={inputStyle?.textAlign === 'center'}
              onPick={() => {
                setDraft(option.label);
                onCommit?.(option.value);
                close();
              }}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function ComboboxOption({
  label,
  selected,
  centered,
  onPick,
}: {
  label: string;
  selected: boolean;
  centered: boolean;
  onPick: () => void;
}) {
  const [hovered, setHovered] = useState(false);
  return (
    <button
      type="button"
      role="option"
      aria-selected={selected}
      onMouseDown={(event) => event.preventDefault()}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      onClick={onPick}
      style={{ ...menuRowStyle(hovered, false, selected), justifyContent: centered ? 'center' : 'flex-start' }}
    >
      {label}
    </button>
  );
}
