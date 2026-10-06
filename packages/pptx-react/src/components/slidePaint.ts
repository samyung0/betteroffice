import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx/viewer';
import type { CanvasImageResolver, SlideDisplayList } from '@betteroffice/pptx/viewer';
import { useEffect, useRef } from 'react';
import type { RefObject } from 'react';

/** What presenting needs from a deck: the editor's or the viewer's handle. */
export interface PresentationSource {
  layoutSlide(slideIndex: number): SlideDisplayList;
}

/**
 * Paints slide `index` into the canvas, fitted into `box` (CSS px), off screen
 * first so a slow picture never shows a half-painted slide, and only the
 * latest request lands. The canvas may live in another window (the notes
 * pop-up): the buffer is this document's, whose fonts the slides use.
 */
export function useSlidePaint(
  canvasRef: RefObject<HTMLCanvasElement | null>,
  {
    handle,
    index,
    resolveImage,
    dpr,
    box,
    onError,
  }: {
    handle: PresentationSource;
    index: number;
    resolveImage: CanvasImageResolver;
    dpr: number;
    box: { width: number; height: number };
    onError: (error: unknown) => void;
  }
) {
  const errorRef = useRef(onError);
  errorRef.current = onError;
  const { width, height } = box;
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || width <= 0 || height <= 0) return;
    let cancelled = false;
    const paint = async () => {
      const frame = handle.layoutSlide(index);
      const buffer = document.createElement('canvas');
      const context = buffer.getContext('2d');
      if (!context) return;
      const scale = Math.min(width / frame.width, height / frame.height);
      sizeCanvasForSlide(buffer, frame, dpr, scale);
      await paintSlide(context, frame, dpr, scale, { resolveImage });
      if (cancelled) return;
      sizeCanvasForSlide(canvas, frame, dpr, scale);
      canvas.getContext('2d')?.drawImage(buffer, 0, 0);
    };
    void paint().catch((error: unknown) => {
      if (!cancelled) errorRef.current(error);
    });
    return () => {
      cancelled = true;
    };
  }, [canvasRef, handle, index, resolveImage, dpr, width, height]);
}
