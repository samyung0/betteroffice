import type { TextCmd } from '../display-list/types';

const PX_PER_PT = 96 / 72;

// css font shorthand: `[italic] [bold] <size>px <family>`. Workbook fonts are
// often not installed on the host (Calibri on macOS); without a fallback the
// browser's unknown-family default is serif.
export function fontString(cmd: TextCmd): string {
  const style = cmd.italic ? 'italic ' : '';
  const weight = cmd.bold ? 'bold ' : '';
  const family = cmd.fontFamily
    ? cmd.chart
      ? cmd.fontFamily
      : `"${cmd.fontFamily.replace(/"/g, '\\"')}", sans-serif`
    : 'sans-serif';
  return `${style}${weight}${fontSizePx(cmd)}px ${family}`;
}

/**
 * The CSS font a cell's text is painted with, `scale` times its size: DOM text
 * over the grid (the in-cell editor) passes the zoom to match the canvas.
 */
export function cellFont(
  font: Pick<TextCmd, 'fontFamily' | 'fontSize' | 'bold' | 'italic'>,
  scale = 1
): string {
  const style = font.italic ? 'italic ' : '';
  const weight = font.bold ? 'bold ' : '';
  const family = font.fontFamily
    ? `"${font.fontFamily.replace(/"/g, '\\"')}", sans-serif`
    : 'sans-serif';
  return `${style}${weight}${font.fontSize * PX_PER_PT * scale}px ${family}`;
}

export function fontSizePx(cmd: TextCmd): number {
  return cmd.fontSize * PX_PER_PT;
}
