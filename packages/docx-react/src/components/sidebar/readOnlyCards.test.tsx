import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { Comment } from '@betteroffice/docx/types/content';
import type { TrackedChangeEntry } from './cardUtils';
import { useCommentSidebarItems } from '../../hooks/useCommentSidebarItems';
import { CommentCard } from './CommentCard';
import { TrackedChangeCard } from './TrackedChangeCard';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');

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
    const shown = Array.from(container.querySelectorAll('textarea, input')).filter(
      (node) => !node.closest('[hidden]')
    );
    expect(shown.length > 0).toBe(!readOnly);
    cleanup();
  }
});

test('a reply being typed survives a pause and is not sent while it lasts', () => {
  const replies: string[] = [];
  function Sidebar({ readOnly }: { readOnly: boolean }) {
    const items = useCommentSidebarItems({
      comments: [comment],
      trackedChanges: [],
      callbacks: { onCommentReply: (_id, text) => replies.push(text) },
      readOnly,
    });
    return (
      <>
        {items.map((item) => (
          <div key={item.id}>{item.render(expanded)}</div>
        ))}
      </>
    );
  }
  const view = render(<Sidebar readOnly={false} />);
  fireEvent.click(view.getByPlaceholderText('Reply or add others with @'));
  fireEvent.change(view.getByPlaceholderText('Reply or add others with @'), {
    target: { value: 'Draft' },
  });

  view.rerender(<Sidebar readOnly />);
  const draft = view.getByDisplayValue('Draft');
  expect(draft.closest('[hidden]')).not.toBeNull();

  view.rerender(<Sidebar readOnly={false} />);
  expect(view.getByDisplayValue('Draft').closest('[hidden]')).toBeNull();
  expect(replies).toEqual([]);
});

test('a comment being written survives a pause hidden, and is not sent while it lasts', () => {
  const added: string[] = [];
  function Sidebar({ readOnly }: { readOnly: boolean }) {
    const items = useCommentSidebarItems({
      comments: [],
      trackedChanges: [],
      callbacks: { onAddComment: (text) => added.push(text) },
      isAddingComment: true,
      addCommentYPosition: 0,
      readOnly,
    });
    return (
      <>
        {items.map((item) => (
          <div key={item.id}>{item.render(expanded)}</div>
        ))}
      </>
    );
  }
  const view = render(<Sidebar readOnly={false} />);
  fireEvent.change(view.getByPlaceholderText('Add a comment...'), {
    target: { value: 'Note' },
  });

  view.rerender(<Sidebar readOnly />);
  expect(view.getByDisplayValue('Note').closest('[hidden]')).not.toBeNull();

  view.rerender(<Sidebar readOnly={false} />);
  expect(view.getByDisplayValue('Note').closest('[hidden]')).toBeNull();
  expect(added).toEqual([]);
});
