import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { ComponentProps } from 'react';
import { LocaleProvider } from '../i18n';
import type { FormattingAction } from './Toolbar';
import { Toolbar } from './Toolbar';
import { IconSetContext, TOOLBAR_ICON_NAMES, type IconProps, type IconSet } from './ui/ToolbarIcon';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

const noop = () => {};
const handlers = {
  onFormat: noop,
  onShapeFormat: noop,
  onInsertSlide: noop,
  onInsertImage: noop,
  onSave: noop,
  onExportPng: noop,
  onUndo: noop,
  onRedo: noop,
  onZoomChange: noop,
  onToolChange: noop,
};

function SingleRow(props: ComponentProps<typeof Toolbar>) {
  return (
    <LocaleProvider>
      <Toolbar singleRow slideLayouts={[{ partPath: 'layout1.xml' }]} {...handlers} {...props} />
    </LocaleProvider>
  );
}

const groups = (container: HTMLElement) =>
  Array.from(container.querySelectorAll('[role="group"]'), (group) =>
    group.getAttribute('aria-label')
  );

test('a host icon set replaces every built-in icon', () => {
  const icons = Object.fromEntries(
    TOOLBAR_ICON_NAMES.map((name) => [
      name,
      function HostIcon({ size }: IconProps) {
        return <i data-host-icon={name} data-size={size} />;
      },
    ])
  ) as IconSet;
  const { container } = render(
    <IconSetContext.Provider value={icons}>
      <SingleRow
        textSelectionActive
        shapeSelectionActive
        currentShapeFormatting={{ geometry: 'roundRect', adjustments: { adj: 0.2 } }}
      />
    </IconSetContext.Provider>
  );
  expect(container.querySelector('svg')).toBeNull();
  for (const name of ['newSlide', 'undo', 'select', 'shape', 'bold', 'textColor', 'alignLeft', 'fillColor'])
    expect(container.querySelector(`[data-host-icon="${name}"]`)).not.toBeNull();
});

test("follows Google Slides' order and leaves file, arrange and overflow to the host", () => {
  const { container, queryByTestId } = render(<SingleRow />);
  expect(groups(container)).toEqual(['Slides', 'History', 'Zoom', 'Tools']);
  for (const id of ['pptx-save', 'pptx-export-png', 'pptx-shape-arrange', 'pptx-toolbar-more', 'pptx-bold'])
    expect(queryByTestId(id)).toBeNull();
});

test('shows the text and shape sections only for a matching selection', () => {
  const { container, getByTestId, queryByLabelText } = render(
    <SingleRow textSelectionActive shapeSelectionActive currentFormatting={{ fontSize: 20 }} />
  );
  expect(groups(container)).toEqual([
    'Slides',
    'History',
    'Zoom',
    'Tools',
    'Font',
    'Text formatting',
    'Alignment',
    'Shape formatting',
  ]);
  expect((getByTestId('pptx-font-size') as HTMLInputElement).value).toBe('20');
  expect(queryByLabelText('Decrease font size')).toBeNull();
  expect(queryByLabelText('Increase font size')).toBeNull();
});

test('the alignment dropdown applies the picked alignment', () => {
  const actions: FormattingAction[] = [];
  const { getByLabelText, getByTestId } = render(
    <SingleRow textSelectionActive currentFormatting={{ align: 'l' }} onFormat={(action) => actions.push(action)} />
  );
  fireEvent.click(getByTestId('pptx-align'));
  fireEvent.click(getByLabelText('Justify'));
  expect(actions).toEqual([{ type: 'align', value: 'just' }]);
});

test('hides zoom, font and size when the host asks', () => {
  const { container, queryByTestId } = render(
    <SingleRow
      textSelectionActive
      showFontPicker={false}
      showFontSizePicker={false}
      showZoomControl={false}
    />
  );
  expect(groups(container)).toEqual(['Slides', 'History', 'Tools', 'Text formatting', 'Alignment']);
  expect(queryByTestId('pptx-zoom')).toBeNull();
});
