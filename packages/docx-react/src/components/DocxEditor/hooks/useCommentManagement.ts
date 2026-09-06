import { useCallback, useEffect, useRef, useState } from 'react';
import type { Comment } from '@betteroffice/docx/types/content';
import { commentSharedId, projectYrsComments, type YrsRawOp } from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from '../PagedEditor';

interface FloatingCommentBtn {
  top: number;
  left: number;
}

/** Manages controlled and uncontrolled comments. */
export function useCommentManagement({
  commentsProp,
  onCommentDelete,
  onCommentsChange,
  pagedEditorRef,
}: {
  commentsProp: Comment[] | undefined;
  onCommentDelete: ((comment: Comment) => void) | undefined;
  onCommentsChange: ((comments: Comment[]) => void) | undefined;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
}) {
  const [internalComments, setInternalComments] = useState<Comment[]>([]);
  const isControlledComments = commentsProp !== undefined;
  const comments = isControlledComments ? commentsProp : internalComments;

  const [isAddingComment, setIsAddingComment] = useState(false);
  const [commentSelectionRange, setCommentSelectionRange] = useState<{
    from: number;
    to: number;
  } | null>(null);
  const [addCommentYPosition, setAddCommentYPosition] = useState<number | null>(null);
  const [floatingCommentBtn, setFloatingCommentBtn] = useState<FloatingCommentBtn | null>(null);

  // Synchronous mirrors used by stable callbacks. Assigned on every render so
  // the latest value is always visible from the callbacks that read `.current`.
  const cleanOrphanedCommentsTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const commentsRef = useRef(comments);
  commentsRef.current = comments;
  const isAddingCommentRef = useRef(isAddingComment);
  isAddingCommentRef.current = isAddingComment;
  const onCommentDeleteRef = useRef(onCommentDelete);
  onCommentDeleteRef.current = onCommentDelete;
  const onCommentsChangeRef = useRef(onCommentsChange);
  onCommentsChangeRef.current = onCommentsChange;

  // Unified setter that resolves the new value, mutates internal state when
  // uncontrolled, and always notifies via `onCommentsChange`. Reads through
  // `commentsRef.current` for the functional-update branch so the callback
  // stays stable across renders.
  const setComments = useCallback(
    (next: React.SetStateAction<Comment[]>) => {
      const resolved =
        typeof next === 'function'
          ? (next as (prev: Comment[]) => Comment[])(commentsRef.current)
          : next;
      if (resolved === commentsRef.current) return;
      if (!isControlledComments) {
        commentsRef.current = resolved;
        setInternalComments(resolved);
      }
      onCommentsChangeRef.current?.(resolved);
    },
    [isControlledComments]
  );

  const updateComments = useCallback(
    (next: React.SetStateAction<Comment[]>) => {
      const previous = commentsRef.current;
      const resolved = typeof next === 'function' ? next(previous) : next;
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      if (!session) return;
      const live = new Map(session.listComments().map((comment) => [comment.id, comment]));
      const ops: YrsRawOp[] = [];
      const remaining = new Set(resolved.map(commentSharedId));
      for (const comment of previous) {
        const id = commentSharedId(comment);
        if (!remaining.has(id) && live.has(id)) ops.push({ op: 'removeComment', id });
      }
      for (const comment of resolved) {
        const id = commentSharedId(comment);
        const old = previous.find((item) => commentSharedId(item) === id);
        const fields: Extract<YrsRawOp, { op: 'patchComment' }>['fields'] = {};
        if (!old) {
          fields.author = comment.author;
          fields.date = comment.date ?? '';
          fields.body = comment.content;
          fields.done = comment.done ?? false;
          const parent = resolved.find((item) => item.id === comment.parentId);
          fields.parentId = parent ? commentSharedId(parent) : null;
        } else {
          if (old.done !== comment.done) fields.done = comment.done ?? false;
          if (old.content !== comment.content) fields.body = comment.content;
        }
        if (Object.keys(fields).length) ops.push({ op: 'patchComment', id, fields });
      }
      if (ops.length) {
        session.applyRawOps('body', ops);
        editor?.syncYrsInputState(true);
      }
      setComments(projectYrsComments(session, resolved));
    },
    [pagedEditorRef, setComments]
  );

  // Remove comments whose sticky Yrs anchors no longer exist in the document. Called
  // debounced from the document-change handler so the user doesn't see
  // comments vanish mid-edit.
  const cleanOrphanedComments = useCallback(() => {
    if (isAddingCommentRef.current) return;
    const session = pagedEditorRef.current?.getYrsSession();
    if (!session) return;

    const liveIds = new Set<number>();
    for (const comment of commentsRef.current) {
      if (comment.parentId != null) continue;
      try {
        if (session.resolveComment(commentSharedId(comment)).length > 0) liveIds.add(comment.id);
      } catch {
        // Missing anchors are handled as orphaned comments below.
      }
    }

    const currentComments = commentsRef.current;
    const orphanedIds = new Set<number>();
    for (const c of currentComments) {
      if (c.parentId == null && !liveIds.has(c.id)) {
        orphanedIds.add(c.id);
      }
    }
    if (orphanedIds.size === 0) return;

    for (const c of currentComments) {
      if (orphanedIds.has(c.id)) onCommentDeleteRef.current?.(c);
    }
    setComments((prev) =>
      prev.filter((c) => !orphanedIds.has(c.id) && !orphanedIds.has(c.parentId!))
    );
  }, [pagedEditorRef, setComments]);

  // Unmount cleanup for the orphan-cleanup debouncer.
  useEffect(() => {
    return () => {
      if (cleanOrphanedCommentsTimerRef.current) {
        clearTimeout(cleanOrphanedCommentsTimerRef.current);
      }
    };
  }, []);

  return {
    comments,
    setComments,
    updateComments,
    isAddingComment,
    setIsAddingComment,
    isAddingCommentRef,
    commentSelectionRange,
    setCommentSelectionRange,
    addCommentYPosition,
    setAddCommentYPosition,
    floatingCommentBtn,
    setFloatingCommentBtn,
    cleanOrphanedCommentsTimerRef,
    cleanOrphanedComments,
  };
}
