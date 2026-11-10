import { expect, type Page } from 'playwright/test';

export const slideCanvas = (page: Page) => page.getByTestId('pptx-slide-canvas');

/** Capy's PptxEditor takes slide input once the slide canvas shows. */
export async function inputReady(page: Page) {
  await expect(slideCanvas(page)).toBeVisible({ timeout: 60_000 });
}
