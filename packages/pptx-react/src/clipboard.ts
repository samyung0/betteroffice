import type { ShapeSnapshot, StorySnapshot, TextStyleSnapshot } from '@betteroffice/pptx';

export interface ClipboardText {
  text: string;
  html: string;
}

/** The story's `[start, end)` as plain text and as HTML paragraphs keeping bold, italic and underline. */
export function storyClipboard(story: StorySnapshot, start: number, end: number): ClipboardText {
  let text = '';
  let html = '';
  let offset = 0;
  for (const [index, paragraph] of story.paragraphs.entries()) {
    if (index > 0) {
      if (offset >= start && offset < end) text += '\n';
      offset += 1;
    }
    const paragraphStart = offset;
    let content = '';
    for (const run of paragraph.runs) {
      const part = run.text.slice(Math.max(0, start - offset), Math.max(0, end - offset));
      if (part) {
        text += part;
        content += runHtml(part, run.style);
      }
      offset += run.text.length;
    }
    if (paragraphStart <= end && offset >= start) html += `<p>${content}</p>`;
  }
  return { text, html };
}

/** A shape's whole text: its stories, then its group children's, one per line; hidden shapes are left out. */
export function shapeClipboard(shape: ShapeSnapshot): ClipboardText {
  const parts = shapeStories(shape)
    .map((story) => storyClipboard(story, 0, story.length))
    .filter((part) => part.text.trim());
  return {
    text: parts.map((part) => part.text).join('\n'),
    html: parts.map((part) => part.html).join(''),
  };
}

function shapeStories(shape: ShapeSnapshot): StorySnapshot[] {
  return shape.hidden ? [] : [...shape.textStories, ...shape.children.flatMap(shapeStories)];
}

function runHtml(text: string, style: TextStyleSnapshot): string {
  let html = text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
  if (style.underline && style.underline !== 'none') html = `<u>${html}</u>`;
  if (style.italic) html = `<i>${html}</i>`;
  if (style.bold) html = `<b>${html}</b>`;
  return html;
}
