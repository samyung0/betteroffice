/**
 * TableMoreDropdown - Compact dropdown for less-used table actions
 *
 * Contains: delete row/column/table, vertical alignment, header row,
 * distribute columns, auto-fit, table alignment, cell margins,
 * text direction, no-wrap, row height, table properties.
 */

import { useState, useCallback } from 'react';
import type { CSSProperties } from 'react';
import { Button } from './Button';
import { Tooltip } from './Tooltip';
import { MaterialSymbol } from './MaterialSymbol';
import { cn } from '../../lib/utils';
import type { TableAction } from './TableToolbar';
import { useFixedDropdown } from '../../hooks/useFixedDropdown';
import { useTranslation } from '../../i18n';

export interface TableMoreDropdownProps {
  onAction: (action: TableAction) => void;
  disabled?: boolean;
  tableContext?: {
    isInTable: boolean;
    rowCount?: number;
    columnCount?: number;
    canSplitCell?: boolean;
    hasMultiCellSelection?: boolean;
    table?: { attrs?: { justification?: string } };
  } | null;
}

const menuItemStyles: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 10,
  padding: '7px 14px',
  fontSize: 13,
  color: 'var(--doc-text)',
  cursor: 'pointer',
  border: 'none',
  backgroundColor: 'transparent',
  width: '100%',
  textAlign: 'left',
};

const separatorStyles: CSSProperties = {
  height: 1,
  backgroundColor: 'var(--doc-border)',
  margin: '4px 0',
};

export function TableMoreDropdown({
  onAction,
  disabled = false,
  tableContext,
}: TableMoreDropdownProps) {
  const { t } = useTranslation();
  const [isOpen, setIsOpen] = useState(false);
  const [hoveredItem, setHoveredItem] = useState<string | null>(null);
  const close = useCallback(() => setIsOpen(false), []);
  const { containerRef, dropdownRef, dropdownStyle, handleMouseDown } = useFixedDropdown({
    isOpen,
    onClose: close,
    align: 'right',
  });

  const handleAction = useCallback(
    (action: TableAction) => {
      onAction(action);
      setIsOpen(false);
    },
    [onAction]
  );

  const menuItem = (
    id: string,
    icon: string,
    label: string,
    action: TableAction,
    opts?: { danger?: boolean; itemDisabled?: boolean }
  ) => {
    const isItemDisabled = disabled || opts?.itemDisabled;
    return (
      <button
        key={id}
        type="button"
        role="menuitem"
        style={{
          ...menuItemStyles,
          backgroundColor:
            hoveredItem === id && !isItemDisabled ? 'var(--doc-bg-hover)' : 'transparent',
          color: isItemDisabled
            ? 'var(--doc-text-muted)'
            : opts?.danger
              ? 'var(--doc-error)'
              : 'var(--doc-text)',
          cursor: isItemDisabled ? 'not-allowed' : 'pointer',
        }}
        onClick={() => !isItemDisabled && handleAction(action)}
        onMouseEnter={() => setHoveredItem(id)}
        onMouseLeave={() => setHoveredItem(null)}
        disabled={isItemDisabled}
      >
        <MaterialSymbol
          name={icon}
          size={16}
          className={opts?.danger && !isItemDisabled ? 'text-destructive' : ''}
        />
        <span style={{ flex: 1 }}>{label}</span>
      </button>
    );
  };

  const button = (
    <Button
      variant="ghost"
      size="icon-sm"
      className={cn(
        'text-muted-foreground hover:text-foreground hover:bg-muted/80',
        isOpen && 'bg-muted',
        disabled && 'opacity-30 cursor-not-allowed'
      )}
      onMouseDown={handleMouseDown}
      onClick={() => !disabled && setIsOpen((prev) => !prev)}
      disabled={disabled}
      aria-label={t('table.moreOptions')}
      aria-expanded={isOpen}
      aria-haspopup="menu"
      data-testid="toolbar-table-more"
    >
      <MaterialSymbol name="more_vert" size={20} />
    </Button>
  );

  return (
    <div ref={containerRef} style={{ position: 'relative', display: 'inline-block' }}>
      {!isOpen ? <Tooltip content={t('table.moreOptions')}>{button}</Tooltip> : button}

      {isOpen && !disabled && (
        <div
          ref={dropdownRef}
          style={{
            ...dropdownStyle,
            backgroundColor: 'var(--doc-surface)',
            border: '1px solid var(--doc-border)',
            borderRadius: 8,
            boxShadow: '0 4px 16px var(--doc-shadow)',
            padding: '4px 0',
            minWidth: 200,
            maxHeight: '70vh',
            overflowY: 'auto',
          }}
          role="menu"
          onMouseDown={(e) => e.stopPropagation()}
        >
          {/* Insert actions */}
          {menuItem('addRowAbove', 'add', t('table.insertRowAbove'), 'addRowAbove')}
          {menuItem('addRowBelow', 'add', t('table.insertRowBelow'), 'addRowBelow')}
          {menuItem('addColumnLeft', 'add', t('table.insertColumnLeft'), 'addColumnLeft')}
          {menuItem('addColumnRight', 'add', t('table.insertColumnRight'), 'addColumnRight')}

          <div style={separatorStyles} role="separator" />

          {/* Merge/Split */}
          {menuItem('mergeCells', 'call_merge', t('table.mergeCells'), 'mergeCells', {
            itemDisabled: !tableContext?.hasMultiCellSelection,
          })}
          {menuItem('splitCell', 'call_split', t('table.splitCell'), 'splitCell', {
            itemDisabled: !tableContext?.canSplitCell,
          })}

          <div style={separatorStyles} role="separator" />

          {/* Select */}
          {menuItem('selectTable', 'select_all', t('table.selectTable'), 'selectTable')}

          <div style={separatorStyles} role="separator" />

          {/* Delete actions */}
          {menuItem('deleteRow', 'delete', t('table.deleteRow'), 'deleteRow', {
            danger: true,
            itemDisabled: (tableContext?.rowCount ?? 0) <= 1,
          })}
          {menuItem('deleteColumn', 'delete', t('table.deleteColumn'), 'deleteColumn', {
            danger: true,
            itemDisabled: (tableContext?.columnCount ?? 0) <= 1,
          })}
          {menuItem('deleteTable', 'delete', t('table.deleteTable'), 'deleteTable', {
            danger: true,
          })}

          <div style={separatorStyles} role="separator" />

          {menuItem('properties', 'settings', t('tableAdvanced.tableProperties'), {
            type: 'openTableProperties',
          })}
        </div>
      )}
    </div>
  );
}

export default TableMoreDropdown;
