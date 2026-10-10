import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import type { ResidentFontRequirement } from '../layout/measure';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { preloadParseWasm } from '../wasm/parse';
import { createYrsSession } from './index';

const root = resolve(import.meta.dir, '../../../..');

describe('resident worker fonts', () => {
  beforeAll(() =>
    Promise.all([
      preloadEditWasm(
        new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
      ),
      preloadOpcWasm(),
      preloadParseWasm(),
    ])
  );

  it('hands the worker the engine’s own copy of each face, registered once', async () => {
    const session = await createYrsSession({ clientId: 72_100 });
    try {
      session.openDocx(new Uint8Array(readFileSync(resolve(root, 'poc/fixtures/feature-rich.docx'))), true);
      const font = new Uint8Array(
        readFileSync(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))
      );
      const original = font.slice();
      const id = session.registerFont(font);
      // A later load of the same face hands over fresh bytes.
      expect(session.registerFont(font.slice())).toBe(id);
      // The engine keeps none of the caller's buffers.
      font.fill(0);
      const request = buildResidentRegionLayoutRequest(session.materializeDocx()!, 24, {});
      const requirements = JSON.parse(
        session.layoutFontRequirementsJson(JSON.stringify(request))
      ) as ResidentFontRequirement[];
      request.measurement = {
        fontChains: Object.fromEntries(requirements.map(({ key }) => [key, [id]])),
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        compat: { noLeading: false, doNotExpandShiftReturn: false },
        authoritativeShaping: true,
      };
      session.layoutDocumentWithRegionsRetainedJson(JSON.stringify(request));
      const snapshot = session.residentWorkerSnapshot()!;
      expect(snapshot.fonts).toHaveLength(1);
      expect(snapshot.fonts[0]).toEqual(original);
      const current = session.residentWorkerSnapshot({ knownFontsRevision: snapshot.fontsRevision })!;
      expect(current.fonts).toEqual([]);
    } finally {
      session.destroy();
    }
  });
});
