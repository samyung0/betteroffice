import type { CanvasImageResolver } from '@betteroffice/pptx/viewer';
import { useEffect, useRef, useState } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { useTranslation } from '../i18n';
import { NOTES_SIZES, elapsedMs, formatElapsed } from '../presenter';
import type { PresenterTimer } from '../presenter';
import { type PresentationSource, useSlidePaint } from './slidePaint';
import { DrawnIcon } from './ui/ToolbarIcon';

/**
 * Presenter view's speaker notes window, after Google Slides' (notes first):
 * the timer, the slides' state and the controls on top, the current and next
 * slide small below, the notes across the full width. Always dark.
 */
export function PresenterWindow({
  target,
  handle,
  resolveImage,
  index,
  slideCount,
  notes,
  fullscreen,
  timer,
  onTimer,
  notesSize,
  onNotesSize,
  onStep,
  onEnd,
  onError,
}: {
  /** The pop-up this renders into: its width and pixel ratio size the slides. */
  target: Window;
  handle: PresentationSource;
  resolveImage: CanvasImageResolver;
  index: number;
  slideCount: number;
  notes: string;
  /** Whether the show's own window (the slides) is in full screen. */
  fullscreen: boolean;
  timer: PresenterTimer;
  onTimer: (action: 'pause' | 'resume' | 'reset') => void;
  notesSize: number;
  onNotesSize: (size: number) => void;
  onStep: (delta: number) => void;
  onEnd: () => void;
  onError: (error: unknown) => void;
}) {
  const { t } = useTranslation();
  const [width, setWidth] = useState(() => target.innerWidth);
  const [now, setNow] = useState(() => Date.now());
  const running = timer.since !== null;

  useEffect(() => {
    const onResize = () => setWidth(target.innerWidth);
    target.addEventListener('resize', onResize);
    return () => target.removeEventListener('resize', onResize);
  }, [target]);

  useEffect(() => {
    setNow(Date.now());
    if (!running) return;
    const tick = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(tick);
  }, [running, timer]);

  const current = Math.max(200, Math.min(520, (width - 60) * 0.58));
  const hasNext = index < slideCount - 1;
  const status = fullscreen ? t('presentation.slidesFullScreen') : t('presentation.slidesWindowed');

  return (
    <main aria-label={t('presentation.presenterView')} style={styles.root}>
      <div style={styles.top}>
        <span style={styles.timer} role="timer" aria-label={t('presentation.elapsed')}>
          {formatElapsed(elapsedMs(timer, now))}
        </span>
        <IconButton
          label={running ? t('presentation.pauseTimer') : t('presentation.resumeTimer')}
          onClick={() => onTimer(running ? 'pause' : 'resume')}
        >
          {running ? (
            <DrawnIcon name="presenterPause" size={16}>
              <Glyph d="M8 5v14M16 5v14" />
            </DrawnIcon>
          ) : (
            <DrawnIcon name="presenterResume" size={16}>
              <Glyph d="M7 5v14l12-7z" />
            </DrawnIcon>
          )}
        </IconButton>
        <IconButton label={t('presentation.resetTimer')} onClick={() => onTimer('reset')}>
          <DrawnIcon name="presenterReset" size={16}>
            <Glyph d="M4 12a8 8 0 1 0 2.4-5.7M4 4v4h4" />
          </DrawnIcon>
        </IconButton>
        <span style={styles.separator} aria-hidden="true" />
        <span style={styles.status} role="status">
          <span
            style={{ ...styles.dot, background: fullscreen ? palette.success : palette.warning }}
            aria-hidden="true"
          />
          {status}
        </span>
        <span style={{ flex: 1 }} />
        <label style={styles.size}>
          <span style={styles.nudged} aria-hidden="true">
            <DrawnIcon name="presenterNotesSize" size={16}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                <text x="1" y="18" fontSize="15" fontWeight="600">Aa</text>
              </svg>
            </DrawnIcon>
          </span>
          <span style={{ color: palette.muted }}>{t('presentation.notesSize')}</span>
          <select
            aria-label={t('presentation.notesSizeLabel')}
            value={notesSize}
            onChange={(event) => onNotesSize(Number(event.target.value))}
            style={styles.select}
          >
            {NOTES_SIZES.map((size) => (
              <option key={size} value={size}>
                {size}
              </option>
            ))}
          </select>
        </label>
        <button type="button" onClick={onEnd} style={styles.end}>
          {t('presentation.end')}
        </button>
      </div>
      <div style={styles.slides}>
        <div style={{ width: current }}>
          <div style={styles.caps}>{t('presentation.now')}</div>
          <SlideCanvas
            target={target}
            handle={handle}
            resolveImage={resolveImage}
            index={index}
            width={current}
            onClick={() => onStep(1)}
            onError={onError}
          />
          <div style={styles.nav}>
            <IconButton
              label={t('presentation.previousSlide')}
              onClick={() => onStep(-1)}
              disabled={index === 0}
              size={36}
            >
              <DrawnIcon name="presentationPrevious" size={16}>
                <Glyph d="m15 6-6 6 6 6" />
              </DrawnIcon>
            </IconButton>
            <span style={styles.counter} aria-live="polite">
              {t('presentation.slidePosition', { current: index + 1, total: slideCount })}
            </span>
            <IconButton
              label={t('presentation.nextSlide')}
              onClick={() => onStep(1)}
              disabled={!hasNext}
              size={36}
            >
              <DrawnIcon name="presentationNext" size={16}>
                <Glyph d="m9 6 6 6-6 6" />
              </DrawnIcon>
            </IconButton>
          </div>
        </div>
        <div style={{ width: current * 0.6 }}>
          <div style={styles.caps}>{t('presentation.next')}</div>
          {hasNext ? (
            <SlideCanvas
              target={target}
              handle={handle}
              resolveImage={resolveImage}
              index={index + 1}
              width={current * 0.6}
              onError={onError}
            />
          ) : (
            <div style={styles.endOfSlides}>{t('presentation.endOfSlides')}</div>
          )}
        </div>
      </div>
      <section style={styles.notes} aria-label={t('notes.panelLabel')}>
        <div style={styles.caps}>{t('presentation.notes')}</div>
        {notes.trim() ? (
          <div style={{ ...styles.notesText, fontSize: notesSize }} data-testid="pptx-presenter-notes">
            {notes}
          </div>
        ) : (
          <div style={{ ...styles.notesText, fontSize: notesSize, color: palette.muted }}>
            {t('presentation.noNotes')}
          </div>
        )}
      </section>
    </main>
  );
}

function SlideCanvas({
  target,
  handle,
  resolveImage,
  index,
  width,
  onClick,
  onError,
}: {
  target: Window;
  handle: PresentationSource;
  resolveImage: CanvasImageResolver;
  index: number;
  width: number;
  onClick?: () => void;
  onError: (error: unknown) => void;
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  useSlidePaint(canvasRef, {
    handle,
    index,
    resolveImage,
    dpr: target.devicePixelRatio || 1,
    box: { width, height: Number.POSITIVE_INFINITY },
    onError,
  });
  return (
    <canvas
      ref={canvasRef}
      onClick={onClick}
      style={{ ...styles.slide, cursor: onClick ? 'pointer' : 'default' }}
    />
  );
}

function IconButton({
  label,
  onClick,
  disabled,
  size = 32,
  children,
}: {
  label: string;
  onClick: () => void;
  disabled?: boolean;
  size?: number;
  children: ReactNode;
}) {
  return (
    <button
      type="button"
      aria-label={label}
      title={label}
      onClick={onClick}
      disabled={disabled}
      style={{
        ...styles.iconButton,
        width: size,
        height: size,
        opacity: disabled ? 0.4 : 1,
        cursor: disabled ? 'default' : 'pointer',
      }}
    >
      <span style={styles.nudged}>{children}</span>
    </button>
  );
}

function Glyph({ d }: { d: string }) {
  return (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d={d} />
    </svg>
  );
}

/** Catppuccin Mocha, Capy's dark theme: the window stays dark in a light app. */
const palette = {
  page: '#14151f',
  bar: '#1e1e2e',
  raised: '#313244',
  text: '#cdd6f4',
  secondary: '#a6adc8',
  muted: '#7f849c',
  divider: '#37384a',
  border: '#414256',
  primary: '#c2c7e0',
  onPrimary: '#1e1e2e',
  success: '#a6e3a1',
  warning: '#f9e2af',
};

const styles = {
  root: {
    position: 'fixed',
    inset: 0,
    display: 'flex',
    flexDirection: 'column',
    background: palette.page,
    color: palette.text,
    colorScheme: 'dark',
    fontFamily: 'var(--font-sans, system-ui, -apple-system, "Segoe UI", sans-serif)',
    fontSize: 14,
    fontWeight: 500,
    lineHeight: 1.45,
  },
  top: {
    flex: 'none',
    height: 48,
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    padding: '0 12px 0 16px',
    borderBottom: `1px solid ${palette.divider}`,
    background: palette.bar,
    whiteSpace: 'nowrap',
  },
  timer: {
    fontSize: 20,
    fontWeight: 600,
    fontVariantNumeric: 'tabular-nums',
    minWidth: 56,
    marginRight: 2,
  },
  separator: { width: 1, height: 20, background: palette.divider, margin: '0 4px', flex: 'none' },
  status: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: 6,
    fontSize: 13,
    color: palette.secondary,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  // Capy's Fustat sits its cap band about 1px above the flex centre: icons
  // beside text move up to meet it.
  dot: { width: 8, height: 8, borderRadius: '50%', flex: 'none', transform: 'translateY(-1px)' },
  nudged: { display: 'inline-flex', transform: 'translateY(-1px)' },
  size: {
    height: 32,
    display: 'inline-flex',
    alignItems: 'center',
    gap: 6,
    padding: '0 4px 0 10px',
    borderRadius: 8,
    background: palette.raised,
    fontSize: 14,
  },
  select: {
    height: 28,
    border: 0,
    borderRadius: 6,
    background: 'transparent',
    color: palette.text,
    font: 'inherit',
    fontVariantNumeric: 'tabular-nums',
    cursor: 'pointer',
  },
  end: {
    height: 32,
    marginLeft: 6,
    padding: '0 14px',
    border: 0,
    borderRadius: 10,
    background: palette.primary,
    color: palette.onPrimary,
    font: 'inherit',
    fontWeight: 600,
    cursor: 'pointer',
  },
  iconButton: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    flex: 'none',
    padding: 0,
    border: 0,
    borderRadius: 8,
    background: 'transparent',
    color: palette.secondary,
  },
  slides: {
    flex: 'none',
    display: 'flex',
    gap: 20,
    alignItems: 'flex-start',
    padding: '18px 20px 0',
  },
  caps: {
    margin: '0 0 8px',
    fontSize: 11,
    fontWeight: 600,
    letterSpacing: '0.06rem',
    lineHeight: 1,
    textTransform: 'uppercase',
    color: palette.muted,
  },
  slide: { display: 'block', borderRadius: 4, background: '#000000' },
  nav: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 10,
    marginTop: 8,
    color: palette.secondary,
  },
  counter: { minWidth: 96, textAlign: 'center', fontVariantNumeric: 'tabular-nums' },
  endOfSlides: {
    aspectRatio: '16 / 9',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    border: `1px dashed ${palette.border}`,
    borderRadius: 4,
    color: palette.muted,
    fontSize: 13,
  },
  notes: { flex: 1, minHeight: 0, overflow: 'auto', padding: '18px 20px 20px' },
  notesText: { lineHeight: 1.5, fontWeight: 500, whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' },
} satisfies Record<string, CSSProperties>;
