import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { en } from '@betteroffice/docx-i18n';
import { TableMoreDropdown } from './TableMoreDropdown';

const { cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const context = {
  isInTable: true,
  rowCount: 2,
  columnCount: 2,
  wrapText: true,
  headerRow: false,
};

test('the table ⋮ menu follows Google Docs’ order and runs the table items', () => {
  const onAction = mock((_action: unknown) => {});
  const view = render(<TableMoreDropdown onAction={onAction} tableContext={context} />);
  fireEvent.click(view.getByTestId('toolbar-table-more'));
  const rows = [...view.container.ownerDocument.querySelectorAll('[role^="menuitem"]')].map(
    (item) => item.textContent
  );
  // Delete comes right after Insert, so it shows without scrolling.
  expect(rows).toEqual([
    en.table.insertRowAbove,
    en.table.insertRowBelow,
    en.table.insertColumnLeft,
    en.table.insertColumnRight,
    en.table.deleteRow,
    en.table.deleteColumn,
    en.table.deleteTable,
    en.table.mergeCells,
    en.table.splitCell,
    en.tableAdvanced.distributeColumns,
    en.tableAdvanced.autoFit,
    en.tableAdvanced.pinHeaderRow,
    en.tableAdvanced.wrapText,
    en.table.selectTable,
    en.tableAdvanced.tableProperties,
  ]);
  const ticked = [...view.container.ownerDocument.querySelectorAll('[aria-checked="true"]')].map(
    (item) => item.textContent
  );
  expect(ticked).toEqual([en.tableAdvanced.wrapText]);
  fireEvent.click(view.getByTestId('toolbar-table-more'));
  const run = (id: string) => {
    fireEvent.click(view.getByTestId('toolbar-table-more'));
    fireEvent.click(view.getByTestId(`table-menu-${id}`));
  };
  run('header-row');
  run('wrap-text');
  run('distribute');
  run('autofit');
  expect(onAction.mock.calls.map(([action]) => action)).toEqual([
    { type: 'pinHeaderRow', pinned: true },
    { type: 'wrapText', wrap: false },
    { type: 'distributeColumns' },
    { type: 'autoFitContents' },
  ]);
});

test('a read-only editor’s table ⋮ menu does not open', () => {
  const onAction = mock((_action: unknown) => {});
  const view = render(<TableMoreDropdown onAction={onAction} tableContext={context} disabled />);
  fireEvent.click(view.getByTestId('toolbar-table-more'));
  expect(view.queryByTestId('table-menu-autofit')).toBeNull();
});
