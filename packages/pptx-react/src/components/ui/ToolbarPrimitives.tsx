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
};

export const toolbarFont = 'var(--pptx-font, ui-sans-serif, system-ui, sans-serif)';

/**
 * The popover look a host shares across its editors (`--office-menu-*`, as
 * Capy's note toolbar draws its popovers); the fallbacks are this editor's own.
 */
export const menuColors = {
  bg: `var(--office-menu-bg, ${toolbarColors.surface})`,
  text: `var(--office-menu-text, ${toolbarColors.text})`,
  muted: `var(--office-menu-muted, ${toolbarColors.muted})`,
  secondary: `var(--office-menu-secondary, ${toolbarColors.muted})`,
  border: `var(--office-menu-border, ${toolbarColors.border})`,
  separator: `var(--office-menu-separator, ${toolbarColors.divider})`,
  hover: `var(--office-menu-hover, ${toolbarColors.hover})`,
  selectedBg: `var(--office-menu-selected-bg, ${toolbarColors.active})`,
  selectedText: `var(--office-menu-selected-text, ${toolbarColors.activeText})`,
  buttonHover: `var(--office-menu-button-hover, ${toolbarColors.hover})`,
  pressedBg: `var(--office-menu-pressed-bg, ${toolbarColors.active})`,
  pressedText: `var(--office-menu-pressed-text, ${toolbarColors.activeText})`,
  swatchBorder: `var(--office-menu-swatch-border, ${toolbarColors.border})`,
  ring: 'var(--office-menu-ring, var(--pptx-accent, #1a73e8))',
};

export const menuFontFamily = `var(--office-menu-font-family, ${toolbarFont})`;

/** Every dropdown's panel. */
export const popoverStyle: CSSProperties = {
  padding: 'var(--office-menu-padding, 6px)',
  border: `1px solid ${menuColors.border}`,
  borderRadius: 'var(--office-menu-radius, 8px)',
  background: menuColors.bg,
  color: menuColors.text,
  boxShadow: 'var(--office-menu-shadow, 0 4px 16px rgba(60, 64, 67, 0.24))',
  boxSizing: 'border-box',
};

/** A row in a dropdown: an item, a font, a size or a zoom level. */
export function menuRowStyle(hovered: boolean, disabled = false, selected = false): CSSProperties {
  return {
    appearance: 'none',
    display: 'flex',
    alignItems: 'center',
    gap: 'var(--office-menu-item-gap, 10px)',
    width: '100%',
    minHeight: 'var(--office-menu-item-height, 32px)',
    padding: '0 var(--office-menu-item-padding-x, 9px)',
    border: 0,
    borderRadius: 'var(--office-menu-item-radius, 4px)',
    background: selected ? menuColors.selectedBg : hovered && !disabled ? menuColors.hover : 'transparent',
    color: disabled ? toolbarColors.disabled : selected ? menuColors.selectedText : menuColors.text,
    cursor: disabled ? 'default' : 'pointer',
    opacity: disabled ? 0.48 : 1,
    font: `var(--office-menu-font-weight, 400) var(--office-menu-font-size, 13px) / var(--office-menu-line-height, normal) ${menuFontFamily}`,
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

/** A button in a popover (shapes, alignment), as the host's popover buttons. */
const menuButtonStyle: CSSProperties = {
  minWidth: 'var(--office-menu-button-size, var(--pptx-button-size, 28px))',
  height: 'var(--office-menu-button-size, var(--pptx-button-size, 28px))',
  borderRadius: 'var(--office-menu-button-radius, var(--pptx-button-radius, 4px))',
};

function interactiveButtonStyle(
  disabled: boolean,
  active: boolean,
  hovered: boolean,
  style?: CSSProperties,
  inMenu = false
): CSSProperties {
  const colors = inMenu
    ? { active: menuColors.pressedBg, activeText: menuColors.pressedText, hover: menuColors.buttonHover }
    : toolbarColors;
  return {
    ...baseButtonStyle,
    ...(inMenu ? menuButtonStyle : {}),
    background: active ? colors.active : hovered && !disabled ? colors.hover : 'transparent',
    color: disabled ? toolbarColors.disabled : active ? colors.activeText : toolbarColors.text,
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
  /** Inside a popover: the host's popover button size and colours. */
  inMenu?: boolean;
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
  inMenu = false,
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
      style={interactiveButtonStyle(disabled, active, hovered, style, inMenu)}
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
  /** Overrides on the open panel, e.g. a button popover's tighter padding. */
  menuStyle?: CSSProperties;
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
  menuStyle,
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
            ...menuStyle,
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
/** A dropdown section's heading; the panel's first one sits closer to the top. */
export function ToolbarMenuLabel({ children, first = false }: { children: ReactNode; first?: boolean }) {
  return (
    <div
      style={{
        padding: 'var(--office-menu-label-padding, 12px 8px 6px)',
        ...(first ? { paddingTop: 4 } : {}),
        color: menuColors.muted,
        font: `var(--office-menu-label-font-weight, 500) var(--office-menu-label-font-size, 12px) / var(--office-menu-label-line-height, 16px) ${menuFontFamily}`,
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
      style={{
        height: 1,
        margin: '4px var(--office-menu-item-padding-x, 8px)',
        background: menuColors.separator,
      }}
    />
  );
}
