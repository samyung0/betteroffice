import type { FontDecoder } from '@betteroffice/fonts';

function copyBytes(bytes: ArrayBuffer | Uint8Array): Uint8Array<ArrayBuffer> {
  return (bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)).slice();
}

/** Decodes one packed font face with the edit wasm's linked decoder. */
export const decodeFontFace: FontDecoder = async (bytes, { codec, dict, maxOutput }) => {
  const { preloadEditWasm } = await import('./edit');
  await preloadEditWasm();
  const { decodeFontFace: decode } = await import('./generated/edit/docx_edit.js');
  const face = decode(copyBytes(bytes), codec, dict === undefined ? undefined : copyBytes(dict), maxOutput);
  return {
    font: copyBytes(face.font).buffer,
    ...(face.payload === undefined ? {} : { payload: copyBytes(face.payload).buffer }),
  };
};
