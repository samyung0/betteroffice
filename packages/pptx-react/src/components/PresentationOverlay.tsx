import type { CanvasImageResolver } from '@betteroffice/pptx/viewer';
import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent } from 'react';
import { createPortal } from 'react-dom';
import { useTranslation } from '../i18n';
import { type NotesWindow, type NotesWindowState, isWindow } from '../notesWindow';
import {
  notesSize as validNotesSize,
  pauseTimer,
  resetTimer,
  resumeTimer,
  startTimer,
} from '../presenter';
import { PresenterWindow } from './PresenterWindow';
import { type PresentationSource, useSlidePaint } from './slidePaint';
import { DrawnIcon } from './ui/ToolbarIcon';

export type { PresentationSource } from './slidePaint';

const noWindow = () => null;
const noSubscription = () => () => {};

/**
 * The show: the slides full screen in this window, and with `notesWindow`
 * Presenter view's speaker notes in a pop-up (Google Slides' arrangement).
 * Both windows drive one slide. Leaving full screen (Esc) ends the show,
 * except when this window leaves it to open the notes window; without full
 * screen (refused, or no permission handed over) the show runs windowed with
 * a Full screen button.
 */
export function PresentationOverlay({
  handle,
  slideCount,
  startIndex,
  resolveImage,
  notesWindow,
  notesFor,
  defaultNotesSize,
  onNotesSizeChange,
  onExit,
  onError,
}: {
  handle: PresentationSource;
  slideCount: number;
  startIndex: number;
  resolveImage: CanvasImageResolver;
  /** Presenter view's window; without it the show offers no speaker notes. */
  notesWindow?: NotesWindow;
  /** A slide's speaker notes, for the notes window. */
  notesFor?: (index: number) => string;
  defaultNotesSize?: number;
  onNotesSizeChange?: (size: number) => void;
  /** The show ended at this slide (0-based). */
  onExit: (index: number) => void;
  onError: (error: unknown) => void;
}) {
  const { t } = useTranslation();
  const containerRef = useRef<HTMLDivElement>(null);
  const dialogRef = useRef<HTMLDialogElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [requestedIndex, setIndex] = useState(startIndex);
  const index = Math.max(0, Math.min(requestedIndex, slideCount - 1));
  const [viewport, setViewport] = useState({ width: 0, height: 0 });
  const [fullscreen, setFullscreen] = useState(false);
  const [timer, setTimer] = useState(() => startTimer(Date.now()));
  const [notesSize, setNotesSize] = useState(() => validNotesSize(defaultNotesSize));
  const [menuOpen, setMenuOpen] = useState(false);
  const notes: NotesWindowState = useSyncExternalStore(
    notesWindow?.subscribe ?? noSubscription,
    notesWindow?.get ?? noWindow
  );
  // The next exit from full screen is this window's own, made to open the
  // notes window (Chrome would drop full screen for it anyway): the show stays.
  const keepShowRef = useRef(false);
  const latest = useRef({ index, onExit, onError, notesWindow });
  latest.current = { index, onExit, onError, notesWindow };
  const exit = () => latest.current.onExit(latest.current.index);
  const step = (delta: number) =>
    setIndex(Math.max(0, Math.min(index + delta, slideCount - 1)));

  const enterFullscreen = () => {
    const dialog = dialogRef.current;
    const container = containerRef.current;
    if (!dialog || !container) return;
    void container
      .requestFullscreen?.()
      .then(() => {
        if (!dialog.isConnected) {
          if (document.fullscreenElement === container)
            void document.exitFullscreen().catch(() => undefined);
          return;
        }
        // Back on top of the full-screen element.
        dialog.close();
        dialog.showModal();
        dialog.focus();
      })
      .catch(() => undefined);
  };

  useEffect(() => {
    setIndex((current) => Math.max(0, Math.min(current, slideCount - 1)));
    if (slideCount === 0) exit();
  }, [slideCount]);

  useEffect(() => {
    const dialog = dialogRef.current!;
    const container = containerRef.current!;
    const previousFocus = document.activeElement;
    dialog.showModal();
    dialog.focus();
    const onFullscreenChange = () => {
      const inFullscreen = document.fullscreenElement === container;
      setFullscreen(inFullscreen);
      if (inFullscreen || document.fullscreenElement) return;
      if (keepShowRef.current) keepShowRef.current = false;
      else exit();
    };
    const onResize = () =>
      setViewport({ width: window.innerWidth, height: window.innerHeight });
    onResize();
    window.addEventListener('resize', onResize);
    document.addEventListener('fullscreenchange', onFullscreenChange);
    enterFullscreen();
    return () => {
      window.removeEventListener('resize', onResize);
      document.removeEventListener('fullscreenchange', onFullscreenChange);
      if (document.fullscreenElement === container) {
        void document.exitFullscreen().catch(() => undefined);
      }
      dialog.close();
      latest.current.notesWindow?.close();
      if (previousFocus instanceof HTMLElement && previousFocus.isConnected)
        previousFocus.focus();
    };
  }, []);

  useSlidePaint(canvasRef, {
    handle,
    index,
    resolveImage,
    dpr: window.devicePixelRatio || 1,
    box: viewport,
    onError: (error) => {
      latest.current.onError(error);
      exit();
    },
  });

  // The same keys in both windows.
  const onKey = (event: KeyboardEvent | ReactKeyboardEvent) => {
    if (event.altKey || event.ctrlKey || event.metaKey) return;
    const target = event.target as { closest?: (selector: string) => unknown } | null;
    const control = (selector: string) =>
      typeof target?.closest === 'function' && Boolean(target.closest(selector));
    if (event.key !== 'Escape' && control('select')) return;
    switch (event.key) {
      case ' ':
        if (control('button')) return;
        step(1);
        break;
      case 'ArrowRight':
      case 'ArrowDown':
      case 'PageDown':
        step(1);
        break;
      case 'ArrowLeft':
      case 'ArrowUp':
      case 'PageUp':
        step(-1);
        break;
      case 'Home':
        setIndex(0);
        break;
      case 'End':
        setIndex(Math.max(0, slideCount - 1));
        break;
      case 'Escape':
        exit();
        break;
      default:
        return;
    }
    event.preventDefault();
    event.stopPropagation();
  };
  const keyRef = useRef(onKey);
  keyRef.current = onKey;

  useEffect(() => {
    if (!isWindow(notes)) return;
    const doc = notes.document;
    doc.title = t('presentation.presenterView');
    doc.documentElement.style.colorScheme = 'dark';
    doc.body.style.margin = '0';
    const listener = (event: KeyboardEvent) => keyRef.current(event);
    notes.addEventListener('keydown', listener);
    return () => notes.removeEventListener('keydown', listener);
  }, [notes, t]);

  const openNotes = () => {
    setMenuOpen(false);
    if (document.fullscreenElement) {
      keepShowRef.current = true;
      void document.exitFullscreen().catch(() => {
        keepShowRef.current = false;
      });
    }
    notesWindow?.open();
  };

  const timerAction = (action: 'pause' | 'resume' | 'reset') => {
    const now = Date.now();
    setTimer((current) =>
      action === 'pause'
        ? pauseTimer(current, now)
        : action === 'resume'
          ? resumeTimer(current, now)
          : resetTimer(current, now)
    );
  };

  const offerNotes = notesWindow && !isWindow(notes);

  return (
    <div ref={containerRef}>
      <dialog
        ref={dialogRef}
        aria-label={t('presentation.label')}
        tabIndex={-1}
        style={styles.presentationOverlay}
        onCancel={(event) => {
          event.preventDefault();
          exit();
        }}
        onClick={(event) => {
          if (menuOpen) setMenuOpen(false);
          else if (event.target === dialogRef.current) step(1);
        }}
        onKeyDown={onKey}
      >
        <canvas ref={canvasRef} style={styles.presentationCanvas} />
        <button
          type="button"
          onClick={exit}
          aria-label={t('presentation.exit')}
          title={t('presentation.exit')}
          style={styles.presentationExit}
        >
          <DrawnIcon name="presentationExit" size={16}>
            <svg
              width="16"
              height="16"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              aria-hidden="true"
            >
              <path d="M6 6 18 18M18 6 6 18" />
            </svg>
          </DrawnIcon>
        </button>
        <div style={styles.presentationDock}>
          {notes === 'blocked' && (
            <div role="alert" style={styles.presentationNotice}>
              <span>{t('presentation.notesBlocked')}</span>
              <button type="button" onClick={openNotes} style={styles.presentationTextButton}>
                {t('presentation.openNotes')}
              </button>
            </div>
          )}
          {menuOpen && offerNotes && (
            <div role="menu" style={styles.presentationMenu}>
              <button
                type="button"
                role="menuitem"
                onClick={openNotes}
                style={styles.presentationMenuItem}
              >
                {t('presentation.openNotes')}
              </button>
            </div>
          )}
          <div style={styles.presentationControls}>
            <button
              type="button"
              onClick={() => step(-1)}
              disabled={index === 0}
              aria-label={t('presentation.previousSlide')}
              style={styles.presentationNavButton}
            >
              <DrawnIcon name="presentationPrevious" size={14}>
                <svg
                  width="14"
                  height="14"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.4"
                  aria-hidden="true"
                >
                  <path d="m15 6-6 6 6 6" />
                </svg>
              </DrawnIcon>
            </button>
            <span style={styles.presentationCounter} aria-live="polite">
              {t('presentation.slideCounter', { current: index + 1, total: slideCount })}
            </span>
            <button
              type="button"
              onClick={() => step(1)}
              disabled={index >= slideCount - 1}
              aria-label={t('presentation.nextSlide')}
              style={styles.presentationNavButton}
            >
              <DrawnIcon name="presentationNext" size={14}>
                <svg
                  width="14"
                  height="14"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.4"
                  aria-hidden="true"
                >
                  <path d="m9 6 6 6-6 6" />
                </svg>
              </DrawnIcon>
            </button>
            {!fullscreen && (
              <button type="button" onClick={enterFullscreen} style={styles.presentationTextButton}>
                <DrawnIcon name="presentationFullscreen" size={14}>
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2.2"
                    aria-hidden="true"
                  >
                    <path d="M4 9V4h5M15 4h5v5M20 15v5h-5M9 20H4v-5" />
                  </svg>
                </DrawnIcon>
                {t('presentation.fullScreen')}
              </button>
            )}
            {offerNotes && (
              <button
                type="button"
                aria-label={t('presentation.moreActions')}
                aria-haspopup="menu"
                aria-expanded={menuOpen}
                onClick={(event) => {
                  event.stopPropagation();
                  setMenuOpen(!menuOpen);
                }}
                style={styles.presentationNavButton}
              >
                <DrawnIcon name="presentationMore" size={14}>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                    <circle cx="12" cy="5" r="2" />
                    <circle cx="12" cy="12" r="2" />
                    <circle cx="12" cy="19" r="2" />
                  </svg>
                </DrawnIcon>
              </button>
            )}
          </div>
        </div>
      </dialog>
      {isWindow(notes) &&
        createPortal(
          <PresenterWindow
            target={notes}
            handle={handle}
            resolveImage={resolveImage}
            index={index}
            slideCount={slideCount}
            notes={notesFor?.(index) ?? ''}
            fullscreen={fullscreen}
            timer={timer}
            onTimer={timerAction}
            notesSize={notesSize}
            onNotesSize={(size) => {
              setNotesSize(size);
              onNotesSizeChange?.(size);
            }}
            onStep={step}
            onEnd={exit}
            onError={onError}
          />,
          notes.document.body
        )}
    </div>
  );
}

const styles = {
  presentationOverlay: {
    position: 'fixed',
    margin: 0,
    padding: 0,
    border: 0,
    width: '100vw',
    height: '100vh',
    maxWidth: 'none',
    maxHeight: 'none',
    inset: 0,
    zIndex: 2147483000,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: '#000000',
    cursor: 'pointer',
    outline: 'none',
  },
  presentationCanvas: { display: 'block', pointerEvents: 'none' },
  presentationExit: {
    position: 'absolute',
    top: 16,
    right: 16,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 36,
    height: 36,
    border: 0,
    borderRadius: 999,
    background: 'rgba(255, 255, 255, 0.14)',
    color: '#ffffff',
    cursor: 'pointer',
  },
  presentationDock: {
    position: 'absolute',
    bottom: 20,
    left: '50%',
    transform: 'translateX(-50%)',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: 8,
    cursor: 'default',
  },
  presentationControls: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
    padding: '6px 14px',
    borderRadius: 999,
    background: 'rgba(255, 255, 255, 0.12)',
  },
  presentationNavButton: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 28,
    height: 28,
    border: 0,
    borderRadius: 999,
    background: 'rgba(255, 255, 255, 0.14)',
    color: '#ffffff',
    cursor: 'pointer',
  },
  presentationTextButton: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: 6,
    height: 28,
    padding: '0 12px',
    border: 0,
    borderRadius: 999,
    background: 'rgba(255, 255, 255, 0.14)',
    color: '#ffffff',
    font: 'inherit',
    fontSize: 12,
    fontWeight: 600,
    whiteSpace: 'nowrap',
    cursor: 'pointer',
  },
  presentationCounter: {
    minWidth: 64,
    color: '#ffffff',
    fontSize: 12,
    fontWeight: 600,
    textAlign: 'center',
  },
  presentationNotice: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
    padding: '6px 6px 6px 14px',
    borderRadius: 999,
    background: 'rgba(30, 30, 46, 0.92)',
    color: '#ffffff',
    fontSize: 12,
    fontWeight: 600,
    whiteSpace: 'nowrap',
  },
  presentationMenu: {
    alignSelf: 'flex-end',
    padding: 4,
    borderRadius: 10,
    background: 'rgba(30, 30, 46, 0.96)',
  },
  presentationMenuItem: {
    display: 'flex',
    alignItems: 'center',
    height: 32,
    padding: '0 10px',
    border: 0,
    borderRadius: 6,
    background: 'transparent',
    color: '#ffffff',
    font: 'inherit',
    fontSize: 13,
    fontWeight: 500,
    whiteSpace: 'nowrap',
    cursor: 'pointer',
  },
} satisfies Record<string, CSSProperties>;
