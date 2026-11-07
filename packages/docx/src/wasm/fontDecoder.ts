import type { FontDecoder } from '@betteroffice/fonts';

type EngineFontDecoder = (
  bytes: Uint8Array,
  codec: string,
  dict: Uint8Array | undefined,
  maxOutput: number | undefined
) => { font: Uint8Array; payload?: Uint8Array };

function copyBytes(bytes: ArrayBuffer | Uint8Array): Uint8Array<ArrayBuffer> {
  return (bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)).slice();
}

// The fonts track builds the packed-font decoder into the edit wasm; until
// then the export is absent and decoding fails explicitly.
export const decodeFontFace: FontDecoder = async (bytes, { codec, dict, maxOutput }) => {
  const { preloadEditWasm } = await import('./edit');
  await preloadEditWasm();
  const glue = (await import('./generated/edit/docx_edit.js')) as unknown as Record<string, unknown>;
  const decode = glue.decodeFontFace as EngineFontDecoder | undefined;
  if (typeof decode !== 'function') {
    throw new Error('decodeFontFace is not available in the embedded edit wasm');
  }
  const face = decode(copyBytes(bytes), codec, dict === undefined ? undefined : copyBytes(dict), maxOutput);
  return {
    font: copyBytes(face.font).buffer,
    ...(face.payload === undefined ? {} : { payload: copyBytes(face.payload).buffer }),
  };
};
