import { useEffect, useMemo, useState } from 'react';
import {
  openDocumentViewer,
  type DocxViewerAnalysis,
} from '@betteroffice/docx/viewer';
import {
  createCanvasImageResolver,
  type DisplayList,
} from '@betteroffice/docx/layout/render';
import { CanvasPagesView } from './DocxEditor/CanvasPagesView';
import { DefaultLoadingIndicator, ParseError } from './DocxEditorHelpers';

export interface DocxViewerProps {
  documentBuffer: ArrayBuffer | Uint8Array;
  zoom?: number;
  onAnalysis?: (analysis: DocxViewerAnalysis) => void;
  onError?: (error: Error) => void;
}

export interface DocxDisplayListViewerProps {
  displayList: DisplayList;
  zoom?: number;
}

/** Canvas-only surface for hosts that parse/layout in a disposable worker. */
export function DocxDisplayListViewer({
  displayList,
  zoom = 1,
}: DocxDisplayListViewerProps) {
  const resolveImage = useMemo(() => createCanvasImageResolver(), []);
  return (
    <CanvasPagesView
      displayList={displayList}
      glyphOutlineProvider={false}
      interactive={false}
      resolveImage={resolveImage}
      zoom={zoom}
    />
  );
}

/** Read-only DOCX surface. The editing WASM is never imported by this component. */
export function DocxViewer({
  documentBuffer,
  zoom = 1,
  onAnalysis,
  onError,
}: DocxViewerProps) {
  const [displayList, setDisplayList] = useState<DisplayList | null>(null);
  const [error, setError] = useState<Error | null>(null);

  useEffect(() => {
    let cancelled = false;
    let dispose: (() => void) | undefined;
    setDisplayList(null);
    setError(null);
    const bytes =
      documentBuffer instanceof Uint8Array
        ? documentBuffer
        : new Uint8Array(documentBuffer);
    void openDocumentViewer(bytes)
      .then((handle) => {
        dispose = () => handle.dispose();
        if (cancelled) {
          handle.dispose();
          return;
        }
        const list = handle.displayList();
        setDisplayList(list);
        onAnalysis?.({ format: 'docx', pageCount: list.pages.length });
      })
      .catch((value: unknown) => {
        if (cancelled) return;
        const next = value instanceof Error ? value : new Error(String(value));
        setError(next);
        onError?.(next);
      });
    return () => {
      cancelled = true;
      dispose?.();
    };
  }, [documentBuffer, onAnalysis, onError]);

  if (error) return <ParseError message={error.message} />;
  if (!displayList) return <DefaultLoadingIndicator />;
  return <DocxDisplayListViewer displayList={displayList} zoom={zoom} />;
}
