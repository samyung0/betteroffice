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

test('while a new picture is prepared Apply waits, so the old picture cannot be applied in its place', async () => {
  const applied: (Watermark | null)[] = [];
  let release!: () => void;
  const view = render(
    <WatermarkDialog
      isOpen
      onClose={() => {}}
      onApply={(watermark) => applied.push(watermark)}
      current={{ kind: 'picture', dataUrl: 'data:image/png;base64,b2xk', scale: 1, washout: true }}
      prepareImage={async () => {
        await new Promise<void>((resolve) => {
          release = resolve;
        });
        return new File(['new'], 'new.jpg', { type: 'image/jpeg' });
      }}
    />
  );
  const apply = view.getByRole('button', { name: /apply/i }) as HTMLButtonElement;
  expect(apply.disabled).toBe(false);
  const input = view.container.querySelector('input[type="file"]') as HTMLInputElement;
  await act(async () => {
    fireEvent.change(input, { target: { files: [new File(['photo'], 'photo.webp', { type: 'image/webp' })] } });
    await new Promise((done) => setTimeout(done, 0));
  });
  // Busy: Apply and the picker wait, and the dialog says so.
  expect(apply.disabled).toBe(true);
  expect(input.disabled).toBe(true);
  expect(view.getByRole('status').textContent).toBe('Loading');
  act(() => {
    fireEvent.click(apply);
  });
  expect(applied).toEqual([]);

  await act(async () => {
    release();
    await new Promise((done) => setTimeout(done, 0));
  });
  const prepared = `data:image/jpeg;base64,${btoa('new')}`;
  await waitFor(() => expect(apply.disabled).toBe(false));
  act(() => {
    fireEvent.click(apply);
  });
  expect(applied).toHaveLength(1);
  expect(applied[0]).toMatchObject({ kind: 'picture', dataUrl: prepared });
});
