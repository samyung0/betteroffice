import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { Watermark } from '@betteroffice/docx/types/document';
import { WatermarkDialog } from './WatermarkDialog';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a picture watermark is the file prepareImage returns, and its refusal applies nothing', async () => {
  const applied: (Watermark | null)[] = [];
  const errors: Error[] = [];
  const picked: string[] = [];
  let refuse = false;
  const view = render(
    <WatermarkDialog
      isOpen
      onClose={() => {}}
      onApply={(watermark) => applied.push(watermark)}
      current={{ kind: 'picture', dataUrl: '', scale: 1, washout: true }}
      prepareImage={async (file) => {
        picked.push(file.name);
        if (refuse) throw new Error('image too large');
        return new File(['jpeg'], 'logo.jpg', { type: 'image/jpeg' });
      }}
      onError={(error) => errors.push(error)}
    />
  );
  const input = view.container.querySelector('input[type="file"]') as HTMLInputElement;
  const pick = async (name: string) => {
    await act(async () => {
      fireEvent.change(input, { target: { files: [new File(['webp bytes'], name, { type: 'image/webp' })] } });
      await new Promise((done) => setTimeout(done, 0));
    });
  };

  refuse = true;
  await pick('huge.webp');
  expect(errors.map((error) => error.message)).toEqual(['image too large']);
  expect(view.container.querySelector('img')).toBeNull();

  refuse = false;
  await pick('logo.webp');
  const prepared = `data:image/jpeg;base64,${btoa('jpeg')}`;
  await waitFor(() => expect(view.container.querySelector('img')?.getAttribute('src')).toBe(prepared));
  act(() => {
    fireEvent.click(view.getByRole('button', { name: /apply|ok/i }));
  });
  expect(applied).toHaveLength(1);
  expect(applied[0]).toMatchObject({ kind: 'picture', dataUrl: prepared });
  expect(picked).toEqual(['huge.webp', 'logo.webp']);
});
