import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { en } from '@betteroffice/xlsx-i18n';
import type { ComponentProps } from 'react';
import { Toolbar } from './Toolbar';
import {
  DRAWN_ICON_NAMES,
  IconSetContext,
  TOOLBAR_ICON_NAMES,
  type IconProps,
  type IconSet,
} from './ui/ToolbarIcon';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { cleanup, fireEvent, render, within } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

const noop = () => {};
const formatting = { onFormat: noop, onMerge: noop, onUndo: noop, onRedo: noop };

function SingleRow(props: ComponentProps<typeof Toolbar>) {
  return <Toolbar singleRow onZoomChange={noop} {...formatting} {...props} />;
}

test('a host icon set replaces every built-in icon, the drawn glyphs included', () => {
  const icons = Object.fromEntries(
    [...TOOLBAR_ICON_NAMES, ...DRAWN_ICON_NAMES].map((name) => [
      name,
      function HostIcon({ size }: IconProps) {
        return <i data-host-icon={name} data-size={size} />;
      },
    ])
  ) as IconSet;
  const { container } = render(
    <IconSetContext.Provider value={icons}>
      <SingleRow />
    </IconSetContext.Provider>
  );
  expect(container.querySelector('svg')).toBeNull();
  for (const name of ['menu', 'undo', 'bold', 'borderAll', 'alignTextLeft', 'alignCellMiddle'])
    expect(container.querySelector(`[data-host-icon="${name}"]`)).not.toBeNull();
});

test('the menu holds save, PNG export and print, and the row drops search and size steps', () => {
  const onSave = mock(noop);
  const onExportPng = mock(noop);
  const onPrint = mock(noop);
  const { getByRole, getByTestId, queryByRole } = render(
    <SingleRow
      onSave={onSave}
      onExportPng={onExportPng}
      onPrint={onPrint}
      showSearchMenus={false}
    />
  );
  const bar = getByRole('toolbar');
  expect(bar.dataset.layout).toBe('single-row');
  expect(queryByRole('button', { name: en.toolbar.searchMenus })).toBeNull();
  expect(queryByRole('button', { name: en.toolbar.print })).toBeNull();
  expect(queryByRole('button', { name: en.toolbar.increaseFontSize })).toBeNull();
  expect(queryByRole('button', { name: en.toolbar.decreaseFontSize })).toBeNull();
  expect(queryByRole('button', { name: en.toolbar.more })).toBeNull();
  // Zoom stays pinned outside the scrolling groups.
  const row = getByTestId('xlsx-toolbar-row');
  expect(within(row).queryByTestId('xlsx-zoom')).toBeNull();
  getByTestId('xlsx-zoom');

  for (const [label, handler] of [
    [en.toolbar.save, onSave],
    [en.toolbar.exportPng, onExportPng],
    [en.toolbar.print, onPrint],
  ] as const) {
    fireEvent.click(getByRole('button', { name: en.toolbar.menu }));
    fireEvent.click(getByRole('menuitem', { name: label }));
    expect(handler).toHaveBeenCalledTimes(1);
  }
});

test('showFontPicker, showFontSizePicker and showZoomControl drop their controls', () => {
  const { queryByRole, queryByTestId, getByRole } = render(
    <SingleRow showFontPicker={false} showFontSizePicker={false} showZoomControl={false} />
  );
  expect(queryByRole('button', { name: en.toolbar.fontFamily })).toBeNull();
  expect(queryByRole('combobox', { name: en.toolbar.fontSize })).toBeNull();
  expect(queryByTestId('xlsx-zoom')).toBeNull();
  getByRole('button', { name: en.toolbar.bold });
});

test('a vertical wheel scrolls the clipped row sideways', () => {
  const { getByTestId } = render(<SingleRow />);
  const row = getByTestId('xlsx-toolbar-row');
  Object.defineProperty(row, 'scrollWidth', { configurable: true, value: 500 });
  Object.defineProperty(row, 'clientWidth', { configurable: true, value: 200 });
  const wheel = (deltaY: number) => {
    const event = new WheelEvent('wheel', { cancelable: true, deltaY });
    row.dispatchEvent(event);
    return event;
  };

  expect(wheel(120).defaultPrevented).toBe(true);
  expect(row.scrollLeft).toBe(120);
  wheel(1000);
  expect(row.scrollLeft).toBe(300);
  // At the end the wheel is left to the page.
  expect(wheel(50).defaultPrevented).toBe(false);
});
