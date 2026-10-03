import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { en } from '@betteroffice/xlsx-i18n';
import type { ComponentProps } from 'react';
import { type FormattingAction, Toolbar } from './Toolbar';
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
  for (const name of ['undo', 'bold', 'borderAll', 'alignTextLeft', 'alignCellMiddle'])
    expect(container.querySelector(`[data-host-icon="${name}"]`)).not.toBeNull();
});

test("the row follows Google Sheets' order, with no menu, print, search, size steps or overflow", () => {
  const { getByRole, getByTestId, queryByRole } = render(
    <SingleRow onPrint={noop} showSearchMenus={false} showCustomNumberFormat={false} />
  );
  const bar = getByRole('toolbar');
  expect(bar.dataset.layout).toBe('single-row');
  const groups = within(getByTestId('xlsx-toolbar-row'))
    .getAllByRole('group')
    .map((group) => group.getAttribute('aria-label'));
  expect(groups).toEqual([
    en.toolbar.groups.history,
    en.toolbar.groups.zoom,
    en.toolbar.groups.number,
    en.toolbar.groups.font,
    en.toolbar.fontSize,
    en.toolbar.groups.text,
    en.toolbar.groups.borders,
    en.toolbar.groups.alignment,
  ]);
  for (const name of [
    en.toolbar.searchMenus,
    en.toolbar.print,
    en.toolbar.increaseFontSize,
    en.toolbar.decreaseFontSize,
    en.toolbar.more,
  ])
    expect(queryByRole('button', { name })).toBeNull();

  fireEvent.click(getByRole('button', { name: en.toolbar.moreNumberFormats }));
  getByRole('menuitem', { name: en.toolbar.numberFormats.percent });
  expect(queryByRole('menuitem', { name: en.toolbar.numberFormats.custom })).toBeNull();
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

test('the colour buttons open a palette that applies, resets and takes a custom colour', () => {
  const onFormat = mock((_action: FormattingAction) => {});
  const { getByRole } = render(<SingleRow onFormat={onFormat} />);
  fireEvent.click(getByRole('button', { name: en.toolbar.textColor }));
  fireEvent.click(getByRole('gridcell', { name: 'Dark blue' }));
  fireEvent.click(getByRole('button', { name: en.toolbar.fillColor }));
  fireEvent.click(getByRole('button', { name: en.toolbar.palette.default }));
  expect(onFormat.mock.calls).toEqual([
    [{ type: 'textColor', value: '#1d4ed8' }],
    [{ type: 'clearColor', value: 'fillColor' }],
  ]);
});

test('the alignment popovers stack one pressed button per alignment and apply the pick', () => {
  const onFormat = mock((_action: FormattingAction) => {});
  const { getByRole } = render(
    <SingleRow
      onFormat={onFormat}
      currentFormatting={{ horizontalAlignment: 'center', verticalAlignment: 'bottom' }}
    />
  );
  fireEvent.click(getByRole('button', { name: en.toolbar.horizontalAlignment }));
  const menu = getByRole('menu', { name: en.toolbar.horizontalAlignment });
  const buttons = within(menu).getAllByRole('button');
  expect(buttons.map((button) => button.getAttribute('aria-label'))).toEqual([
    en.toolbar.horizontalAlign.left,
    en.toolbar.horizontalAlign.center,
    en.toolbar.horizontalAlign.right,
  ]);
  expect(buttons.map((button) => button.getAttribute('aria-pressed'))).toEqual([
    null,
    'true',
    null,
  ]);
  expect((buttons[0].parentElement as HTMLElement).style.flexDirection).toBe('column');
  fireEvent.click(buttons[2]);
  fireEvent.click(getByRole('button', { name: en.toolbar.verticalAlignment }));
  fireEvent.click(getByRole('button', { name: en.toolbar.verticalAlign.top }));
  expect(onFormat.mock.calls).toEqual([
    [{ type: 'horizontalAlignment', value: 'right' }],
    [{ type: 'verticalAlignment', value: 'top' }],
  ]);
});
