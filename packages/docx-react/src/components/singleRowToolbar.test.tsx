import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { en } from '@betteroffice/docx-i18n';
import type { ComponentProps } from 'react';
import { EditorToolbar } from './EditorToolbar';
import { DefaultPlaceholder, ParseError } from './DocxEditorHelpers';
import { FindReplaceDialog } from './dialogs/FindReplaceDialog';
import { HostMenus, type DocxMenuModel, type HostMenuEntry } from './DocxEditor/hostMenus';
import {
  DrawnIcon,
  ICON_NAMES,
  IconSetContext,
  MaterialSymbol,
  type IconProps,
  type IconSet,
} from './ui/Icons';

const { cleanup, fireEvent, render, within } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function SingleRow(props: Omit<ComponentProps<typeof EditorToolbar>, 'children'>) {
  return (
    <EditorToolbar singleRow leading={<EditorToolbar.MenuBar folded />} zoom={1} {...props}>
      <EditorToolbar.Toolbar>
        <button type="button" data-testid="trailing" />
      </EditorToolbar.Toolbar>
    </EditorToolbar>
  );
}

const zoomLabel = en.zoom.ariaLabel.replace('{label}', '100%');

test('a host icon set replaces the built-in icons, drawn ones included', () => {
  const hostIcon = (name: string) =>
    function HostIcon({ size }: IconProps) {
      return <i data-host-icon={name} data-size={size} />;
    };
  const icons = Object.fromEntries(ICON_NAMES.map((name) => [name, hostIcon(name)])) as IconSet;
  const icon = (
    <>
      <MaterialSymbol name="format_bold" size={20} />
      <DrawnIcon name="menu-cut" size={16}>
        <svg data-testid="own-cut" />
      </DrawnIcon>
    </>
  );

  const hosted = render(<IconSetContext.Provider value={icons}>{icon}</IconSetContext.Provider>);
  expect(hosted.container.querySelector('[data-host-icon="format_bold"]')).not.toBeNull();
  expect(
    hosted.container.querySelector('[data-host-icon="menu-cut"]')?.getAttribute('data-size')
  ).toBe('16');
  expect(hosted.queryByTestId('own-cut')).toBeNull();
  cleanup();

  const own = render(icon);
  expect(own.container.querySelector('[data-host-icon]')).toBeNull();
  expect(own.getByTestId('own-cut')).toBeTruthy();
  expect(own.container.querySelectorAll('svg')).toHaveLength(2);
});

test('the find dialog and the placeholders draw the host icons', () => {
  const icons = Object.fromEntries(
    ICON_NAMES.map((name) => [
      name,
      function HostIcon({ size }: IconProps) {
        return <i data-host-icon={name} data-size={size} />;
      },
    ])
  ) as IconSet;
  const { container } = render(
    <IconSetContext.Provider value={icons}>
      <FindReplaceDialog
        isOpen
        onClose={() => {}}
        onFind={() => null}
        onFindNext={() => null}
        onFindPrevious={() => null}
        onReplace={() => false}
        onReplaceAll={() => 0}
      />
      <DefaultPlaceholder />
      <ParseError message="Broken" />
    </IconSetContext.Provider>
  );
  expect(container.querySelector('svg')).toBeNull();
  for (const name of [
    'dialog-close',
    'find-previous',
    'find-next',
    'placeholder-document',
    'placeholder-error',
  ])
    expect(container.querySelector(`[data-host-icon="${name}"]`)).not.toBeNull();
});

test('the single row scrolls the menu, history, zoom and groups and pins the trailing items', () => {
  const { getByTestId } = render(<SingleRow />);
  const bar = getByTestId('formatting-bar');
  expect(bar.dataset.layout).toBe('single-row');
  const scroll = within(bar.querySelector('.oox-formatting-bar__scroll') as HTMLElement);
  scroll.getByRole('menubar');
  scroll.getByRole('button', { name: en.formattingBar.undo });
  scroll.getByRole('combobox', { name: zoomLabel });
  scroll.getByRole('button', { name: en.formattingBar.bold });
  scroll.getByRole('combobox', { name: en.font.selectAriaLabel });
  const end = within(bar.querySelector('.oox-formatting-bar__end') as HTMLElement);
  end.getByTestId('trailing');
  expect(end.queryByRole('combobox', { name: zoomLabel })).toBeNull();
  // One size box and one zoom dropdown: no step buttons in a single row.
  expect(bar.querySelector('[data-testid="font-size-decrease"]')).toBeNull();
  expect(bar.querySelector('[data-testid="font-size-increase"]')).toBeNull();
  expect(within(bar).queryByRole('button', { name: en.zoom.zoomOut })).toBeNull();
  expect(within(bar).queryByRole('button', { name: en.zoom.zoomIn })).toBeNull();
});

test('showFontPicker, showFontSizePicker and showZoomControl drop their controls', () => {
  const { getByTestId } = render(
    <SingleRow showFontPicker={false} showFontSizePicker={false} showZoomControl={false} />
  );
  const bar = within(getByTestId('formatting-bar'));
  expect(bar.queryByRole('combobox', { name: en.font.selectAriaLabel })).toBeNull();
  expect(bar.queryAllByLabelText(en.fontSize.selectAriaLabel)).toHaveLength(0);
  expect(bar.queryByRole('combobox', { name: zoomLabel })).toBeNull();
  bar.getByRole('button', { name: en.formattingBar.bold });
});

test("the folded menu's View item toggles the outline and shows whether it is open", () => {
  const onToggleOutline = mock(() => {});
  const menu = (outlineOpen: boolean) => (
    <SingleRow outlineOpen={outlineOpen} onToggleOutline={onToggleOutline} />
  );
  const { getByRole, rerender } = render(menu(false));
  const open = () => fireEvent.click(getByRole('button', { name: en.titleBar.menuBarAriaLabel }));

  open();
  const item = getByRole('menuitemcheckbox', { name: en.editor.showDocumentOutline });
  expect(item.getAttribute('aria-checked')).toBe('false');
  fireEvent.click(item);
  expect(onToggleOutline).toHaveBeenCalledTimes(1);

  rerender(menu(true));
  open();
  expect(
    getByRole('menuitemcheckbox', { name: en.editor.showDocumentOutline }).getAttribute(
      'aria-checked'
    )
  ).toBe('true');
});

test('a vertical wheel scrolls the clipped row sideways and marks the edges left to scroll', () => {
  const { getByTestId } = render(<SingleRow />);
  const row = getByTestId('formatting-bar').querySelector(
    '.oox-formatting-bar__scroll'
  ) as HTMLElement;
  Object.defineProperty(row, 'scrollWidth', { configurable: true, value: 500 });
  Object.defineProperty(row, 'clientWidth', { configurable: true, value: 200 });
  const wheel = (deltaY: number) => {
    const event = new WheelEvent('wheel', { cancelable: true, deltaY });
    row.dispatchEvent(event);
    row.dispatchEvent(new Event('scroll'));
    return event;
  };

  expect(wheel(120).defaultPrevented).toBe(true);
  expect(row.scrollLeft).toBe(120);
  expect(row.hasAttribute('data-scroll-start')).toBe(true);
  expect(row.hasAttribute('data-scroll-end')).toBe(true);

  wheel(1000);
  expect(row.scrollLeft).toBe(300);
  expect(row.hasAttribute('data-scroll-end')).toBe(false);
  // At the end the wheel is left to the page.
  expect(wheel(100).defaultPrevented).toBe(false);
});

test('a wheel over a dropdown open inside the row leaves the row and the dropdown alone', () => {
  const { getByRole, getByTestId } = render(<SingleRow onToggleOutline={() => {}} />);
  const row = getByTestId('formatting-bar').querySelector(
    '.oox-formatting-bar__scroll'
  ) as HTMLElement;
  Object.defineProperty(row, 'scrollWidth', { configurable: true, value: 500 });
  Object.defineProperty(row, 'clientWidth', { configurable: true, value: 200 });
  fireEvent.click(getByRole('button', { name: en.titleBar.menuBarAriaLabel }));
  const item = getByRole('menuitemcheckbox', { name: en.editor.showDocumentOutline });
  expect(row.contains(item)).toBe(true);

  const event = new WheelEvent('wheel', { bubbles: true, cancelable: true, deltaY: 120 });
  item.dispatchEvent(event);
  expect(event.defaultPrevented).toBe(false);
  expect(row.scrollLeft).toBe(0);
  getByRole('menuitemcheckbox', { name: en.editor.showDocumentOutline });
});

test('with the menus in the host the row drops the ☰, strikethrough and super/subscript and offers comment and image', () => {
  const onAddComment = mock(() => {});
  const { getByTestId } = render(
    <EditorToolbar
      singleRow
      hostMenus
      onAddComment={onAddComment}
      onInsertImage={() => {}}
      zoom={1}
    >
      <EditorToolbar.Toolbar />
    </EditorToolbar>
  );
  const bar = within(getByTestId('formatting-bar'));
  expect(bar.queryByRole('menubar')).toBeNull();
  for (const name of [
    en.formattingBar.strikethrough,
    en.formattingBar.superscript,
    en.formattingBar.subscript,
  ])
    expect(bar.queryByRole('button', { name })).toBeNull();
  bar.getByRole('button', { name: en.toolbar.image });
  fireEvent.click(bar.getByRole('button', { name: en.common.comment }));
  expect(onAddComment).toHaveBeenCalledTimes(1);
});

test('host menus describe the editor and run their commands', () => {
  const onFormat = mock(() => {});
  const onZoomChange = mock(() => {});
  const onInsertTable = mock(() => {});
  const onInsertImage = mock(() => {});
  const actions = {
    onAddComment: mock(() => {}),
    onEditAction: mock(() => {}),
    onFindReplace: mock(() => {}),
    onInsertImageFile: mock(() => {}),
    onToggleComments: mock(() => {}),
    showComments: false,
  };
  let model: DocxMenuModel | null = null;
  render(
    <EditorToolbar
      singleRow
      hostMenus
      zoom={1}
      outlineOpen
      onToggleOutline={() => {}}
      onFormat={onFormat}
      onSave={() => {}}
      onPageSetup={() => {}}
      onZoomChange={onZoomChange}
      onInsertTable={onInsertTable}
      onInsertImage={onInsertImage}
    >
      <HostMenus onMenus={(next) => (model = next)} actions={actions} />
    </EditorToolbar>
  );
  const reported = model as DocxMenuModel | null;
  if (!reported) throw new Error('no menus');
  expect(reported.menus.map((menu) => menu.id)).toEqual([
    'file',
    'edit',
    'view',
    'insert',
    'format',
  ]);
  const flat = (entries: HostMenuEntry[]): HostMenuEntry[] =>
    entries.flatMap((entry) =>
      entry.kind === 'submenu' ? [entry, ...flat(entry.items)] : [entry]
    );
  const items = new Map(
    flat(reported.menus.flatMap((menu) => menu.items)).flatMap((entry) =>
      entry.kind === 'item' || entry.kind === 'submenu' ? [[entry.id, entry] as const] : []
    )
  );
  // Placeholders that do nothing, and the clipboard, stay out.
  for (const id of ['insert-toc', 'cut', 'copy', 'paste']) expect(items.has(id)).toBe(false);
  expect(items.get('zoom:100')).toMatchObject({ checked: true });
  expect(items.get('show-outline')).toMatchObject({ checked: true });
  expect(items.get('show-comments')).toMatchObject({ checked: false });

  reported.run('superscript');
  expect(onFormat).toHaveBeenCalledWith('superscript');
  reported.run('insert-table', '3x4');
  expect(onInsertTable).toHaveBeenCalledWith(3, 4);
  reported.run('zoom:150');
  expect(onZoomChange).toHaveBeenCalledWith(1.5);
  const file = new File(['png'], 'cell.png', { type: 'image/png' });
  reported.run('insert-image', undefined, file);
  expect(actions.onInsertImageFile).toHaveBeenCalledWith(file);
  reported.run('insert-image');
  expect(onInsertImage).toHaveBeenCalledTimes(1);
  reported.run('find-replace');
  expect(actions.onFindReplace).toHaveBeenCalledTimes(1);
});

test('host menus list the first 40 paragraph styles and refuse a disabled item', () => {
  const onFormat = mock(() => {});
  const actions = {
    onAddComment: mock(() => {}),
    onEditAction: mock(() => {}),
    onInsertImageFile: mock(() => {}),
    onToggleComments: mock(() => {}),
    showComments: false,
  };
  const documentStyles = Array.from({ length: 50 }, (_, index) => ({
    styleId: `Style${index}`,
    name: `Style ${index}`,
    type: 'paragraph' as const,
  }));
  let model: DocxMenuModel | null = null;
  render(
    <EditorToolbar
      singleRow
      hostMenus
      zoom={1}
      disabled
      onFormat={onFormat}
      documentStyles={documentStyles}
    >
      <HostMenus onMenus={(next) => (model = next)} actions={actions} />
    </EditorToolbar>
  );
  const reported = model as DocxMenuModel | null;
  if (!reported) throw new Error('no menus');
  const styles = reported.menus
    .flatMap((menu) => menu.items)
    .find((entry) => entry.kind === 'submenu' && entry.id === 'format-styles');
  expect(styles?.kind === 'submenu' && styles.items.length).toBe(40);

  reported.run('bold');
  reported.run('style:Style1');
  expect(onFormat).not.toHaveBeenCalled();
});
