import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { ComponentProps } from 'react';
import { LocaleProvider } from '../i18n';
import type { ParagraphFormatting } from '../paragraphFormatting';
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
  for (const name of [
    'newSlide',
    'undo',
    'select',
    'shape',
    'bold',
    'textColor',
    'highlight',
    'alignLeft',
    'lineSpacing',
    'bulletedList',
    'numberedList',
    'indentDecrease',
    'indentIncrease',
    'clearFormatting',
    'fillColor',
  ])
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
    'Lists',
    'Clear formatting',
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
  expect(groups(container)).toEqual([
    'Slides',
    'History',
    'Tools',
    'Text formatting',
    'Alignment',
    'Lists',
    'Clear formatting',
  ]);
  expect(queryByTestId('pptx-zoom')).toBeNull();
});

test('the alignment dropdown also sets the text box vertical alignment', () => {
  const actions: FormattingAction[] = [];
  const { getByLabelText, getByTestId } = render(
    <SingleRow
      textSelectionActive
      currentParagraph={paragraph({ anchor: 'ctr' })}
      onFormat={(action) => actions.push(action)}
    />
  );
  fireEvent.click(getByTestId('pptx-align'));
  expect(getByLabelText('Middle').getAttribute('aria-pressed')).toBe('true');
  fireEvent.click(getByLabelText('Bottom'));
  expect(actions).toEqual([{ type: 'verticalAlign', value: 'b' }]);
});

test('line and paragraph spacing tick the current spacing and offer to add or remove space', () => {
  const actions: FormattingAction[] = [];
  const { getByLabelText, getByTestId } = render(
    <SingleRow
      textSelectionActive
      currentParagraph={paragraph({ lineSpacing: '1.15', spaceBefore: true })}
      onFormat={(action) => actions.push(action)}
    />
  );
  fireEvent.click(getByTestId('pptx-line-spacing'));
  expect(getByLabelText('1.15').querySelector('svg')).not.toBeNull();
  expect(getByLabelText('Single').querySelector('svg')).toBeNull();
  getByLabelText('Remove space before paragraph');
  fireEvent.click(getByLabelText('Double'));
  fireEvent.click(getByTestId('pptx-line-spacing'));
  fireEvent.click(getByLabelText('Add space after paragraph'));
  expect(actions).toEqual([{ type: 'lineSpacing', value: '2' }, { type: 'spaceAfter' }]);
});

test('list buttons toggle their kind and their menus pick a style', () => {
  const actions: FormattingAction[] = [];
  const { getByLabelText, getByTestId } = render(
    <SingleRow
      textSelectionActive
      currentParagraph={paragraph({ list: 'bullet', listPreset: 'disc', canOutdent: false })}
      onFormat={(action) => actions.push(action)}
    />
  );
  expect(getByTestId('pptx-bulleted-list').getAttribute('aria-pressed')).toBe('true');
  expect(getByTestId('pptx-numbered-list').getAttribute('aria-pressed')).toBeNull();
  expect((getByTestId('pptx-decrease-indent') as HTMLButtonElement).disabled).toBe(true);
  fireEvent.click(getByTestId('pptx-bulleted-list'));
  fireEvent.click(getByTestId('pptx-bulleted-list-styles'));
  expect(getByLabelText('● ○ ■').querySelector('svg')).not.toBeNull();
  fireEvent.click(getByLabelText('• ◦ ▪'));
  fireEvent.click(getByTestId('pptx-numbered-list-styles'));
  fireEvent.click(getByLabelText('1) a) i)'));
  fireEvent.click(getByTestId('pptx-increase-indent'));
  fireEvent.click(getByTestId('pptx-clear-formatting'));
  expect(actions).toEqual([
    { type: 'list', kind: 'bullet' },
    { type: 'list', kind: 'bullet', preset: 'bullet' },
    { type: 'list', kind: 'number', preset: 'paren' },
    { type: 'indent', delta: 1 },
    'clearFormatting',
  ]);
});

test('the highlight picker applies a colour or clears it', () => {
  const actions: FormattingAction[] = [];
  const { getByLabelText, getByTestId } = render(
    <SingleRow
      textSelectionActive
      currentFormatting={{ highlight: '#fde047' }}
      onFormat={(action) => actions.push(action)}
    />
  );
  fireEvent.click(getByTestId('pptx-highlight-color'));
  fireEvent.click(getByLabelText('None'));
  expect(actions).toEqual([{ type: 'highlight', value: null }]);
});

function paragraph(overrides: Partial<ParagraphFormatting>): ParagraphFormatting {
  return {
    spaceBefore: false,
    spaceAfter: false,
    canIndent: true,
    canOutdent: true,
    ...overrides,
  };
}
