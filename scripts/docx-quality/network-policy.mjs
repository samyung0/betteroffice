import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const packages = new URL('../../packages/', import.meta.url);
const version = (name) =>
  JSON.parse(readFileSync(new URL(`${name}/package.json`, packages), 'utf8')).version.replace(
    /[.*+?^${}()|[\]\\]/g,
    '\\$&',
  );
const fontUrl = new RegExp(
  `^https://cdn\\.jsdelivr\\.net/npm/@betteroffice/(?:fonts@${version('fonts')}/dist/decoder/fontpack\\.wasm|fonts-latin@${version('fonts-latin')}/assets/[A-Za-z0-9-]+\\.ttf\\.zst|fonts-cjk@${version('fonts-cjk')}/assets/[A-Za-z0-9-]+\\.otf\\.zst|fonts-intl@${version('fonts-intl')}/assets/[A-Za-z0-9-]+\\.ttf\\.zst)$`
);

/** Allow only the jsDelivr font binaries the capture harness requests, fetched without credentials. */
export function isAllowedFontRequest(request) {
  return (
    fontUrl.test(request.url) &&
    request.method === 'GET' &&
    request.bodyBytes === 0 &&
    !request.referer &&
    !request.cookie
  );
}

/** The checkout's copy of a pinned font binary (a pack face or the fonts decoder), only when the URL pins the checkout's package version. */
export function checkoutFontAsset(url) {
  const match =
    /^https:\/\/cdn\.jsdelivr\.net\/npm\/@betteroffice\/(fonts|fonts-cjk|fonts-latin|fonts-intl)@([^/]+)\/(assets\/[A-Za-z0-9-]+\.(?:ttf|otf)\.zst|dist\/decoder\/fontpack\.wasm)$/.exec(
      url,
    );
  if (!match || (match[1] === 'fonts') !== match[3].startsWith('dist/')) return undefined;
  const root = new URL(`${match[1]}/`, packages);
  if (JSON.parse(readFileSync(new URL('package.json', root), 'utf8')).version !== match[2])
    return undefined;
  return fileURLToPath(new URL(match[3], root));
}
