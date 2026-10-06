import { createContext, useContext } from 'react';
import type { ComponentType, CSSProperties, ReactNode } from 'react';

/** Every name `ToolbarIcon` draws. */
export const TOOLBAR_ICON_NAMES = [
  'save',
  'image',
  'insertImage',
  'undo',
  'redo',
  'newSlide',
  'select',
  'textBox',
  'shape',
  'bringToFront',
  'sendToBack',
  'bringForward',
  'sendBackward',
  'fillColor',
  'borderColor',
  'borderWidth',
  'bold',
  'italic',
  'underline',
  'textColor',
  'highlight',
  'alignLeft',
  'alignCenter',
  'alignRight',
  'alignJustify',
  'alignTop',
  'alignMiddle',
  'alignBottom',
  'lineSpacing',
  'bulletedList',
  'numberedList',
  'indentDecrease',
  'indentIncrease',
  'clearFormatting',
  'more',
  'chevronDown',
  'remove',
  'add',
  'check',
  'notes',
] as const;

export type ToolbarIconName = (typeof TOOLBAR_ICON_NAMES)[number];

export interface IconProps {
  size?: number;
  style?: CSSProperties;
}

/** The show's and the notes window's controls, which they otherwise draw themselves. */
export const DRAWN_ICON_NAMES = [
  'presentationExit',
  'presentationPrevious',
  'presentationNext',
  'presentationFullscreen',
  'presentationMore',
  'presenterPause',
  'presenterResume',
  'presenterReset',
  'presenterNotesSize',
] as const;

export type DrawnIconName = (typeof DRAWN_ICON_NAMES)[number];

/** A host's replacement for every toolbar icon and drawn glyph. */
export type IconSet = Record<ToolbarIconName | DrawnIconName, ComponentType<IconProps>>;

/**
 * Provided by `PptxEditor`'s `icons` prop, or by a viewer around
 * `PresentationOverlay`; null keeps the built-in drawings.
 */
export const IconSetContext = createContext<IconSet | null>(null);

/** The host set's `name` when there is one, else the built-in drawing. */
export function DrawnIcon({
  name,
  size,
  children,
}: {
  name: DrawnIconName;
  size: number;
  children: ReactNode;
}) {
  const HostIcon = useContext(IconSetContext)?.[name];
  return HostIcon ? <HostIcon size={size} /> : children;
}

export interface ToolbarIconProps {
  name: ToolbarIconName;
  size?: number;
  style?: CSSProperties;
}

export function ToolbarIcon({ name, size = 20, style }: ToolbarIconProps) {
  const HostIcon = useContext(IconSetContext)?.[name];
  if (HostIcon) return <HostIcon size={size} style={style} />;
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
      style={{ display: 'inline-flex', flexShrink: 0, ...style }}
    >
      {name === 'save' && <path d="M12 3v11m0 0 4-4m-4 4-4-4M5 19h14" />}
      {name === 'image' && (
        <>
          <rect x="3" y="5" width="18" height="14" rx="2" />
          <circle cx="8.5" cy="9.5" r="1.25" />
          <path d="m6 17 4.5-4.5 3.25 3.25L16 13.5l2 2" />
        </>
      )}
      {name === 'insertImage' && (
        <>
          <rect x="3" y="5" width="14" height="12" rx="2" />
          <circle cx="7.5" cy="9" r="1" />
          <path d="m5 15 3.5-3.5 2.5 2.5L13 12l1.5 1.5" />
          <path d="M18 4v6m-3-3h6" />
        </>
      )}
      {name === 'undo' && <path d="m9 7-5 5 5 5M5 12h9a6 6 0 0 1 6 6" />}
      {name === 'redo' && <path d="m15 7 5 5-5 5m4-5h-9a6 6 0 0 0-6 6" />}
      {name === 'newSlide' && (
        <>
          <rect x="3" y="5" width="15" height="14" rx="1.5" />
          <path d="M7 9h7M7 13h5M20 9v6m-3-3h6" />
        </>
      )}
      {name === 'select' && <path d="m6 3 12 9-6 1.5L9 19Z" />}
      {name === 'textBox' && (
        <>
          <rect x="3" y="4" width="18" height="16" rx="1.5" />
          <path d="M8 8h8m-4 0v8m-3 0h6" />
        </>
      )}
      {name === 'shape' && <rect x="4" y="6" width="16" height="12" rx="3" />}
      {name === 'bringToFront' && (
        <>
          <rect x="2" y="8" width="9" height="9" rx="1.3" opacity="0.4" />
          <rect x="6" y="4" width="9" height="9" rx="1.3" fill="currentColor" fillOpacity="0.15" />
          <path d="M17 17 20 14 23 17" />
          <path d="M17 12 20 9 23 12" />
        </>
      )}
      {name === 'sendToBack' && (
        <>
          <rect x="6" y="4" width="9" height="9" rx="1.3" opacity="0.4" />
          <rect x="2" y="8" width="9" height="9" rx="1.3" fill="currentColor" fillOpacity="0.15" />
          <path d="M17 9 20 12 23 9" />
          <path d="M17 14 20 17 23 14" />
        </>
      )}
      {name === 'bringForward' && (
        <>
          <rect x="2" y="8" width="9" height="9" rx="1.3" opacity="0.4" />
          <rect x="6" y="4" width="9" height="9" rx="1.3" fill="currentColor" fillOpacity="0.15" />
          <path d="M17 15 20 12 23 15" />
        </>
      )}
      {name === 'sendBackward' && (
        <>
          <rect x="6" y="4" width="9" height="9" rx="1.3" opacity="0.4" />
          <rect x="2" y="8" width="9" height="9" rx="1.3" fill="currentColor" fillOpacity="0.15" />
          <path d="M17 11 20 14 23 11" />
        </>
      )}
      {name === 'fillColor' && (
        <>
          <path d="m7 4 10 10-5 5-7-7Z" />
          <path d="m5 12 7-7m6 12h2" />
        </>
      )}
      {name === 'borderColor' && (
        <>
          <path d="M5 18h14M8 15l8-8 2 2-8 8H8Z" />
        </>
      )}
      {name === 'borderWidth' && (
        <>
          <path d="M5 7h14" strokeWidth="1" />
          <path d="M5 12h14" strokeWidth="2" />
          <path d="M5 18h14" strokeWidth="3" />
        </>
      )}
      {name === 'bold' && <path d="M7 4h6a4 4 0 0 1 0 8H7zm0 8h7a4 4 0 0 1 0 8H7z" />}
      {name === 'italic' && <path d="M10 4h8M6 20h8M14 4 10 20" />}
      {name === 'underline' && (
        <>
          <path d="M7 4v7a5 5 0 0 0 10 0V4" />
          <path d="M5 20h14" />
        </>
      )}
      {name === 'textColor' && (
        <>
          <path d="m7 16 5-12 5 12M9 11h6" />
          <path d="M5 20h14" strokeWidth="3" />
        </>
      )}
      {name === 'alignLeft' && <path d="M4 6h16M4 10h10M4 14h16M4 18h10" />}
      {name === 'alignCenter' && <path d="M4 6h16M7 10h10M4 14h16M7 18h10" />}
      {name === 'alignRight' && <path d="M4 6h16M10 10h10M4 14h16M10 18h10" />}
      {name === 'alignJustify' && <path d="M4 6h16M4 10h16M4 14h16M4 18h16" />}
      {name === 'highlight' && (
        <>
          <path d="m8 14 7-9 3 2.5-6 9.5Z" />
          <path d="m8 14-1.5 3h3.5l2-1" />
          <path d="M5 20h14" strokeWidth="3" />
        </>
      )}
      {name === 'alignTop' && <path d="M4 4h16M12 20V8m-4 4 4-4 4 4" />}
      {name === 'alignMiddle' && <path d="M4 12h16M12 3v5m-3-2 3 2 3-2M12 21v-5m-3 2 3-2 3 2" />}
      {name === 'alignBottom' && <path d="M4 20h16M12 4v12m-4-4 4 4 4-4" />}
      {name === 'lineSpacing' && (
        <path d="M11 6h9M11 12h9M11 18h9M6 5v14M3.5 7.5 6 5l2.5 2.5M3.5 16.5 6 19l2.5-2.5" />
      )}
      {name === 'bulletedList' && (
        <>
          <path d="M9 6h11M9 12h11M9 18h11" />
          <circle cx="4.5" cy="6" r="1" fill="currentColor" />
          <circle cx="4.5" cy="12" r="1" fill="currentColor" />
          <circle cx="4.5" cy="18" r="1" fill="currentColor" />
        </>
      )}
      {name === 'numberedList' && (
        <path d="M10 6h10M10 12h10M10 18h10M4 4.5 5.5 4v4M4 10.5a1.5 1.5 0 0 1 3 0c0 1.5-3 2-3 3.5h3M4 16h3l-1.5 1.5a1.5 1.5 0 1 1-1.5 2" />
      )}
      {name === 'indentDecrease' && <path d="M4 5h16M11 10h9M11 14h9M4 19h16M7 9.5 4.5 12 7 14.5" />}
      {name === 'indentIncrease' && <path d="M4 5h16M11 10h9M11 14h9M4 19h16M4.5 9.5 7 12l-2.5 2.5" />}
      {name === 'clearFormatting' && <path d="M5 6h12M11 6l-3 13M15 13l5 5m0-5-5 5" />}
      {name === 'more' && (
        <>
          <circle cx="5" cy="12" r="1" fill="currentColor" stroke="none" />
          <circle cx="12" cy="12" r="1" fill="currentColor" stroke="none" />
          <circle cx="19" cy="12" r="1" fill="currentColor" stroke="none" />
        </>
      )}
      {name === 'chevronDown' && <path d="m7 10 5 5 5-5" />}
      {name === 'remove' && <path d="M5 12h14" />}
      {name === 'add' && <path d="M12 5v14M5 12h14" />}
      {name === 'check' && <path d="m5 12 5 5 9-10" />}
      {name === 'notes' && (
        <>
          <rect x="4" y="4" width="16" height="16" rx="2" />
          <path d="M8 9h8M8 13h8M8 17h5" />
        </>
      )}
    </svg>
  );
}
