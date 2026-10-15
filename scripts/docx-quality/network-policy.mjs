import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const packages = new URL('../../packages/', import.meta.url);

/** Allow only pinned jsDelivr font binaries fetched without credentials. */
export function isAllowedFontRequest(request) {
  const font =
    /^https:\/\/cdn\.jsdelivr\.net\/npm\/@betteroffice\/fonts(?:-cjk)?@(?:0\.1\.0|0\.2\.0)\/assets\/[A-Za-z0-9-]+\.(?:ttf|otf)$/.test(
      request.url,
    );
  return (
    font &&
    request.method === 'GET' &&
    request.bodyBytes === 0 &&
    !request.referer &&
    !request.cookie
  );
}

/** The checkout's copy of a pinned font binary, only when the URL pins the checkout's package version. */
export function checkoutFontAsset(url) {
  const match =
    /^https:\/\/cdn\.jsdelivr\.net\/npm\/@betteroffice\/(fonts|fonts-cjk)@([^/]+)\/assets\/([A-Za-z0-9-]+\.(?:ttf|otf))$/.exec(
      url,
    );
  if (!match) return undefined;
  const root = new URL(`${match[1]}/`, packages);
  if (JSON.parse(readFileSync(new URL('package.json', root), 'utf8')).version !== match[2])
    return undefined;
  return fileURLToPath(new URL(`assets/${match[3]}`, root));
}
