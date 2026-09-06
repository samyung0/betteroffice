/**
 * Comment + tracked-change ID allocation and factories.
 *
 * The allocator and `createComment` factory now live in core (shared with Vue,
 * instance-scoped); re-exported here so existing React import sites stay stable.
 * `DocxEditor` owns one `CommentIdAllocator` per editor instance and threads it
 * into the hooks that allocate IDs.
 *
 * `EMPTY_ANCHOR_POSITIONS` is React-only sidebar state and stays here.
 */

export {
  PENDING_COMMENT_ID,
  createCommentIdAllocator,
  type CommentIdAllocator,
} from '@betteroffice/docx/utils';
import {
  createComment as createBaseComment,
  type CommentIdAllocator as Allocator,
} from '@betteroffice/docx/utils';
import { commentNumericId } from '@betteroffice/docx/yrs';

export function createComment(
  allocator: Allocator,
  text: string,
  author: string,
  parentId?: number
) {
  const sharedId = crypto.randomUUID();
  return {
    ...createBaseComment(allocator, text, author, parentId),
    id: commentNumericId(sharedId),
    sharedId,
  };
}

/** Stable empty Map used as the initial anchor-positions state. */
export const EMPTY_ANCHOR_POSITIONS = new Map<string, number>();
