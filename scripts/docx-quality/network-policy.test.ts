import { afterEach, describe, expect, mock, test } from 'bun:test';
import { BUNDLED_FONTS, type BundledFontFace } from '../../packages/fonts/src/manifest';
import { checkoutFontAsset, isAllowedFontRequest } from './network-policy.mjs';

// cdn.ts builds a face's URL inside the loader it hands the provider; keep that loader to reach every face.
let loadFace: (face: BundledFontFace) => Promise<ArrayBuffer> = async () => {
  throw new Error('cdn.ts did not create a provider');
};
const { fontProvider } = await import('../../packages/fonts/src/provider');
mock.module('../../packages/fonts/src/provider', () => ({
  fontProvider: (load: typeof loadFace) => {
    loadFace = load;
    return fontProvider(load);
  },
}));
const { createFontProvider } = await import('../../packages/fonts/src/cdn');

const originalFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = originalFetch;
});

describe('checkout font assets', () => {
  test('every manifest face cdn.ts requests is a checkout file with its byte length', async () => {
    const requested: string[] = [];
    globalThis.fetch = (async (input: RequestInfo | URL) => {
      const url = String(input);
      requested.push(url);
      const file = checkoutFontAsset(url);
      if (!file) throw new Error(`no checkout file for ${url}`);
      return new Response(Bun.file(file));
    }) as typeof fetch;
    createFontProvider();
    for (const face of BUNDLED_FONTS) {
      expect((await loadFace(face)).byteLength).toBe(face.byteLength);
    }
    expect(requested).toHaveLength(BUNDLED_FONTS.length);
    for (const url of requested) {
      const other = url.replace(/@[^/@]+\/assets\//, '@0.0.1/assets/');
      expect(other).not.toBe(url);
      expect(checkoutFontAsset(other)).toBeUndefined();
    }
  });
});

function request(url, overrides = {}) {
  return {
    url,
    method: 'GET',
    bodyBytes: 0,
    referer: null,
    cookie: null,
    ...overrides,
  };
}

describe('font network policy', () => {
  test('allows pinned legacy and current Latin binaries', () => {
    for (const version of ['0.1.0', '0.2.0']) {
      expect(
        isAllowedFontRequest(
          request(
            `https://cdn.jsdelivr.net/npm/@betteroffice/fonts@${version}/assets/Carlito-Regular.ttf`,
          ),
        ),
      ).toBe(true);
    }
  });

  test('allows pinned legacy and current CJK binaries', () => {
    for (const version of ['0.1.0', '0.2.0']) {
      expect(
        isAllowedFontRequest(
          request(
            `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-cjk@${version}/assets/NotoSansJP-Regular.otf`,
          ),
        ),
      ).toBe(true);
    }
  });

  test('blocks unpinned versions and query strings', () => {
    const blocked = [
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.3.0/assets/Carlito-Regular.ttf',
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.2.1/assets/Carlito-Regular.ttf',
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@1.0.0/assets/Carlito-Regular.ttf',
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@latest/assets/Carlito-Regular.ttf',
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts-cjk@0.3.0/assets/NotoSansJP-Regular.otf',
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.2.0/assets/Carlito-Regular.ttf?v=1',
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.2.0/assets/Carlito-Regular.ttf?foo=bar',
    ];
    for (const url of blocked) expect(isAllowedFontRequest(request(url))).toBe(false);
  });

  test('blocks other hosts and asset paths', () => {
    const blocked = [
      'https://unpkg.com/@betteroffice/fonts@0.2.0/assets/Carlito-Regular.ttf',
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.2.0/dist/cdn.js',
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.2.0/assets/Carlito-Regular.woff2',
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.2.0/assets/Carlito-Regular.ttf/extra',
      'https://fonts.googleapis.com/css2?family=Carlito',
      'http://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.2.0/assets/Carlito-Regular.ttf',
    ];
    for (const url of blocked) expect(isAllowedFontRequest(request(url))).toBe(false);
  });

  test('blocks unsafe request properties', () => {
    const url =
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.2.0/assets/Carlito-Regular.ttf';
    expect(isAllowedFontRequest(request(url, { method: 'POST' }))).toBe(false);
    expect(isAllowedFontRequest(request(url, { bodyBytes: 12 }))).toBe(false);
    expect(
      isAllowedFontRequest(request(url, { referer: 'http://127.0.0.1:4178/' })),
    ).toBe(false);
    expect(isAllowedFontRequest(request(url, { cookie: 'a=b' }))).toBe(false);
  });
});
