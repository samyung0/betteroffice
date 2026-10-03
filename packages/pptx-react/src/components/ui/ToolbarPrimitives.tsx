import { useCallback, useEffect, useRef, useState } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { ToolbarIcon } from './ToolbarIcon';

/** Hosts restyle the chrome through these `--pptx-*` variables. */
export const toolbarColors = {
  text: 'var(--pptx-text, #3c4043)',
  muted: 'var(--pptx-text-muted, #5f6368)',
  disabled: 'var(--pptx-text-disabled, #9aa0a6)',
  hover: 'var(--pptx-hover, #e2e7ef)',
  active: 'var(--pptx-pressed, #d3e3fd)',
  activeText: 'var(--pptx-pressed-text, #3c4043)',
  border: 'var(--pptx-border, #c7cacf)',
  divider: 'var(--pptx-divider, #c7cacf)',
  surface: 'var(--pptx-surface, #ffffff)',
  rail: 'var(--pptx-rail, #edf2fa)',
  menuShadow: 'var(--pptx-menu-shadow, 0 4px 16px rgba(60, 64, 67, 0.24))',
};

export const toolbarFont = 'var(--pptx-font, ui-sans-serif, system-ui, sans-serif)';

/** Every dropdown's panel; hosts match their own popovers through `--pptx-menu-*`. */
export const popoverStyle: CSSProperties = {
  padding: 'var(--pptx-menu-padding, 6px)',
  border: `1px solid ${toolbarColors.border}`,
  borderRadius: 'var(--pptx-menu-radius, 8px)',
  background: toolbarColors.surface,
  boxShadow: toolbarColors.menuShadow,
  boxSizing: 'border-box',
};

/** A row in a dropdown: an item, a font, a size or a zoom level. */
export function menuRowStyle(hovered: boolean, disabled = false, selected = false): CSSProperties {
  return {
    appearance: 'none',
    display: 'flex',
    alignItems: 'center',
    gap: 'var(--pptx-menu-item-gap, 10px)',
    width: '100%',
    minHeight: 'var(--pptx-menu-item-height, 32px)',
    padding: '0 var(--pptx-menu-item-padding-x, 9px)',
    border: 0,
    borderRadius: 'var(--pptx-menu-item-radius, 4px)',
    background: selected
      ? 'var(--pptx-selected-bg, #d3e3fd)'
      : hovered && !disabled
        ? `var(--pptx-menu-hover, ${toolbarColors.hover})`
        : 'transparent',
    color: disabled
      ? toolbarColors.disabled
      : selected
        ? 'var(--pptx-selected-text, #3c4043)'
        : toolbarColors.text,
    cursor: disabled ? 'default' : 'pointer',
    opacity: disabled ? 0.48 : 1,
    font: `var(--pptx-menu-font-weight, 400) var(--pptx-menu-font-size, 13px) ${toolbarFont}`,
    textAlign: 'left',
    boxSizing: 'border-box',
  };
}

const baseButtonStyle: CSSProperties = {
  appearance: 'none',
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  gap: 2,
  minWidth: 'var(--pptx-button-size, 28px)',
  height: 'var(--pptx-button-size, 28px)',
  padding: '0 5px',
  border: 0,
  borderRadius: 'var(--pptx-button-radius, 4px)',
  color: toolbarColors.text,
  font: `var(--pptx-control-font-weight, 500) var(--pptx-control-font-size, 13px) ${toolbarFont}`,
  lineHeight: 1,
  whiteSpace: 'nowrap',
  boxSizing: 'border-box',
};

function interactiveButtonStyle(
  disabled: boolean,
  active: boolean,
  hovered: boolean,
  style?: CSSProperties
): CSSProperties {
  return {
    ...baseButtonStyle,
    background: active
      ? toolbarColors.active
      : hovered && !disabled
        ? toolbarColors.hover
        : 'transparent',
    color: disabled
      ? toolbarColors.disabled
      : active
        ? toolbarColors.activeText
        : toolbarColors.text,
    cursor: disabled ? 'default' : 'pointer',
    opacity: disabled ? 0.48 : 1,
    ...style,
  };
}

export interface ToolbarButtonProps {
  active?: boolean;
  disabled?: boolean;
  title: string;
  onClick?: () => void;
  children: ReactNode;
  style?: CSSProperties;
  testId?: string;
  ariaExpanded?: boolean;
}

export function ToolbarButton({
  active = false,
  disabled = false,
  title,
  onClick,
  children,
  style,
  testId,
  ariaExpanded,
}: ToolbarButtonProps) {
  const [hovered, setHovered] = useState(false);
  return (
    <button
      type="button"
      data-testid={testId}
      disabled={disabled}
      aria-label={title}
      aria-pressed={active || undefined}
      aria-expanded={ariaExpanded}
      title={title}
      onMouseDown={(event) => event.preventDefault()}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      onClick={disabled ? undefined : onClick}
      style={interactiveButtonStyle(disabled, active, hovered, style)}
    >
      {children}
    </button>
  );
}

export function ToolbarGroup({
  label,
  children,
  style,
}: {
  label: string;
  children: ReactNode;
  style?: CSSProperties;
}) {
  return (
    <div
      role="group"
      aria-label={label}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        gap: 1,
        flex: '0 0 auto',
        ...style,
      }}
    >
      {children}
    </div>
  );
}

export function ToolbarSeparator({ style }: { style?: CSSProperties }) {
  return (
    <div
      role="separator"
      style={{
        width: 1,
        height: 24,
        margin: '0 5px',
        background: toolbarColors.divider,
        flex: '0 0 auto',
        ...style,
      }}
    />
  );
}

export interface ToolbarDropdownProps {
  title: string;
  trigger: ReactNode;
  children: (close: () => void) => ReactNode;
  disabled?: boolean;
  active?: boolean;
  menuWidth?: number;
  style?: CSSProperties;
  testId?: string;
}

export function ToolbarDropdown({
  title,
  trigger,
  children,
  disabled = false,
  active = false,
  menuWidth = 220,
  style,
  testId,
}: ToolbarDropdownProps) {
  const [open, setOpen] = useState(false);
  const [hovered, setHovered] = useState(false);
  const [position, setPosition] = useState({ top: 0, left: 0 });
  const triggerRef = useRef<HTMLButtonElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open || !triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    const left = Math.min(rect.left, Math.max(8, window.innerWidth - menuWidth - 8));
    setPosition({ top: rect.bottom + 4, left });
  }, [open, menuWidth]);

  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: MouseEvent) => {
      const target = event.target as Node;
      if (!triggerRef.current?.contains(target) && !menuRef.current?.contains(target)) close();
    };
    const onKeyDown = (event: globalThis.KeyboardEvent) => {
      if (event.key === 'Escape') close();
    };
    const onScroll = () => close();
    document.addEventListener('mousedown', onPointerDown);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('scroll', onScroll, true);
    return () => {
      document.removeEventListener('mousedown', onPointerDown);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('scroll', onScroll, true);
    };
  }, [open, close]);

  return (
    <div style={{ position: 'relative', display: 'inline-flex', flex: '0 0 auto' }}>
      <button
        ref={triggerRef}
        type="button"
        data-testid={testId}
        disabled={disabled}
        aria-label={title}
        aria-haspopup="menu"
        aria-expanded={open}
        title={title}
        onMouseDown={(event) => event.preventDefault()}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        onClick={() => !disabled && setOpen((value) => !value)}
        style={interactiveButtonStyle(disabled, active || open, hovered, style)}
      >
        {trigger}
      </button>
      {open && (
        <div
          ref={menuRef}
          role="menu"
          aria-label={title}
          onMouseDown={(event) => event.preventDefault()}
          style={{
            position: 'fixed',
            top: position.top,
            left: position.left,
            zIndex: 10000,
            width: menuWidth,
            maxHeight: 'min(440px, calc(100vh - 16px))',
            overflowY: 'auto',
            ...popoverStyle,
          }}
        >
          {children(close)}
        </div>
      )}
    </div>
  );
}

export interface ToolbarMenuItemProps {
  label: string;
  icon?: ReactNode;
  selected?: boolean;
  disabled?: boolean;
  onClick?: () => void;
  close?: () => void;
}

export function ToolbarMenuItem({
  label,
  icon,
  selected = false,
  disabled = false,
  onClick,
  close,
}: ToolbarMenuItemProps) {
  const [hovered, setHovered] = useState(false);
  return (
    <button
      type="button"
      role="menuitem"
      disabled={disabled}
      aria-label={label}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      onClick={() => {
        if (disabled) return;
        onClick?.();
        close?.();
      }}
      style={menuRowStyle(hovered, disabled)}
    >
      {icon !== undefined && (
        <span
          style={{
            display: 'inline-grid',
            placeItems: 'center',
            width: 16,
            flex: '0 0 auto',
          }}
        >
          {icon}
        </span>
      )}
      <span style={{ flex: 1 }}>{label}</span>
      {selected && <ToolbarIcon name="check" size={16} />}
    </button>
  );
}

/** A dropdown section's heading, as a host's popover groups. */
export function ToolbarMenuLabel({ children }: { children: ReactNode }) {
  return (
    <div
      style={{
        padding: 'var(--pptx-menu-label-padding, 4px 8px 6px)',
        color: toolbarColors.muted,
        font: `var(--pptx-menu-label-font-weight, 500) var(--pptx-menu-label-font-size, 12px) / 16px ${toolbarFont}`,
      }}
    >
      {children}
    </div>
  );
}

export function ToolbarMenuSeparator() {
  return (
    <div
      role="separator"
      style={{ height: 1, margin: '5px 2px', background: toolbarColors.border }}
    />
  );
}
