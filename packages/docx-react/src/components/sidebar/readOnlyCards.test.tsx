import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { Comment } from '@betteroffice/docx/types/content';
import type { TrackedChangeEntry } from './cardUtils';
import { CommentCard } from './CommentCard';
import { TrackedChangeCard } from './TrackedChangeCard';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const expanded = { isExpanded: true, onToggleExpand: () => {}, measureRef: () => {} };
const comment = { id: 1, author: 'Ada', content: [] } as Comment;
const change = {
  type: 'insertion',
  text: 'added',
  author: 'Ada',
  from: 0,
  to: 5,
  revisionId: 7,
} as unknown as TrackedChangeEntry;

test('read-only comment and tracked-change cards show the thread without reply, resolve, accept or reject', () => {
  for (const readOnly of [false, true]) {
    const { container } = render(
      <>
        <CommentCard {...expanded} comment={comment} replies={[]} readOnly={readOnly} />
        <TrackedChangeCard {...expanded} change={change} replies={[]} readOnly={readOnly} />
      </>
    );
    expect(container.querySelectorAll('.docx-comment-card').length).toBe(1);
    expect(container.querySelectorAll('button').length > 0).toBe(!readOnly);
    expect(container.querySelectorAll('textarea, input').length > 0).toBe(!readOnly);
    cleanup();
  }
});
