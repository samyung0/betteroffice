import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { ColorPicker, type ColorPaletteColor } from './ColorPicker';

const { act, cleanup, fireEvent, render } = await import('@testing-library/react');

const PALETTE: ColorPaletteColor[] = [
  { name: 'Black', value: '#000000' },
  { name: 'Red', value: '#dc2626' },
  { name: 'Blue', value: '#2563eb' },
];

const realRect = Element.prototype.getBoundingClientRect;

afterEach(() => {
  cleanup();
  Element.prototype.getBoundingClientRect = realRect;
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function open(container: HTMLElement) {
  fireEvent.click(container.querySelector('.docx-color-picker-arrow')!);
}

test('a host palette replaces the theme colours and picks as the standard colours do', () => {
  const onChange = mock();
  const { container, getByRole, queryByText } = render(
    <ColorPicker mode="text" value="DC2626" onChange={onChange} palette={PALETTE} />
  );
  open(container);

  expect(container.querySelectorAll('.docx-color-palette [role=gridcell]')).toHaveLength(3);
  expect(queryByText('Theme Colors')).toBeNull();
  expect(getByRole('gridcell', { name: 'Red' }).getAttribute('aria-selected')).toBe('true');
  // The custom row is the label and the swatch, without the hex code.
  expect(container.querySelector('.docx-color-palette__custom')?.textContent).not.toContain('#');

  fireEvent.click(getByRole('gridcell', { name: 'Blue' }));
  expect(onChange).toHaveBeenCalledWith({ rgb: '2563EB' });
});

test('the palette highlights with a hex and clears with none', () => {
  const onChange = mock();
  const { container, getByRole } = render(
    <ColorPicker mode="highlight" onChange={onChange} palette={PALETTE} />
  );
  open(container);
  fireEvent.click(getByRole('gridcell', { name: 'Blue' }));
  expect(onChange).toHaveBeenLastCalledWith('2563EB');

  open(container);
  fireEvent.click(getByRole('button', { name: 'No Color' }));
  expect(onChange).toHaveBeenLastCalledWith('none');
});

test('the palette ticks the selection\'s highlight, named or hex', () => {
  const palette = [...PALETTE, { name: 'Yellow', value: '#ffff00' }];
  const ticked = (value: string) => {
    const { container, getByRole, unmount } = render(
      <ColorPicker mode="highlight" value={value} palette={palette} />
    );
    open(container);
    const selected = palette
      .filter(({ name }) => getByRole('gridcell', { name }).getAttribute('aria-selected') === 'true')
      .map(({ name }) => name);
    unmount();
    return selected;
  };
  expect(ticked('yellow')).toEqual(['Yellow']);
  expect(ticked('2563EB')).toEqual(['Blue']);
  expect(ticked('#2563EB')).toEqual(['Blue']);
  expect(ticked('none')).toEqual([]);
});

test('a dropdown near the window edge moves back inside it', async () => {
  Object.defineProperty(window, 'innerWidth', { configurable: true, value: 300 });
  Element.prototype.getBoundingClientRect = function (this: Element) {
    const rect = (left: number, width: number) =>
      ({
        left,
        right: left + width,
        top: 0,
        bottom: 32,
        width,
        height: 32,
        x: left,
        y: 0,
      } as DOMRect);
    if (this.classList.contains('docx-color-picker-dropdown')) return rect(0, 256);
    if (this.classList.contains('docx-color-picker')) return rect(280, 40);
    return realRect.call(this);
  };
  const { container } = render(<ColorPicker mode="text" palette={PALETTE} />);
  open(container);
  await act(() => new Promise((resolve) => setTimeout(resolve, 50)));

  const dropdown = container.querySelector<HTMLElement>('.docx-color-picker-dropdown')!;
  expect(dropdown.style.left).toBe(`${300 - 256 - 4}px`);
});
