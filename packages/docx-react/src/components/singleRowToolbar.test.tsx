import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { en } from '@betteroffice/docx-i18n';
import type { ComponentProps } from 'react';
import { EditorToolbar } from './EditorToolbar';
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

test('the single row scrolls the menu, history and groups and pins zoom and trailing items', () => {
  const { getByTestId } = render(<SingleRow />);
  const bar = getByTestId('formatting-bar');
  expect(bar.dataset.layout).toBe('single-row');
  const scroll = within(bar.querySelector('.oox-formatting-bar__scroll') as HTMLElement);
  scroll.getByRole('menubar');
  scroll.getByRole('button', { name: en.formattingBar.undo });
  scroll.getByRole('button', { name: en.formattingBar.bold });
  scroll.getByRole('combobox', { name: en.font.selectAriaLabel });
  const end = within(bar.querySelector('.oox-formatting-bar__end') as HTMLElement);
  end.getByRole('combobox', { name: zoomLabel });
  end.getByTestId('trailing');
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
