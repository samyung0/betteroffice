# @betteroffice/docx-i18n

UI strings, locale types and translation helpers for the
[`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react) editor. `en` defines
the strings; the other locales override them.

## Install

```bash
bun add @betteroffice/docx-i18n
```

## Usage

Pass a locale to the editor's `i18n` prop:

```tsx
import { zhCN } from '@betteroffice/docx-i18n';
import { DocxEditor } from '@betteroffice/docx-react';

export function Editor({ file }: { file: File }) {
  return <DocxEditor documentBuffer={file} i18n={zhCN} />;
}
```

Override individual strings by spreading a locale. Keys that are missing or `null` in an override fall
back to English.

```ts
const custom = { ...zhCN, formattingBar: { ...zhCN.formattingBar, bold: '加粗' } };
```

## Locales

`en` (source) and `zh-CN`, exported as `zhCN`. Each locale is also its own subpath,
`@betteroffice/docx-i18n/zh-CN`, so an app that picks the locale at runtime can load only the one it
needs. `locales` maps both tags to their strings.

## Helpers

The package exports the types `LocaleStrings`, `PartialLocaleStrings`, `Translations`,
`TranslationKey`, `LocaleCode` and `TFunction`, and the functions `createT` and `deepMerge`.
Outside React, build a typed `t()` directly:

```ts
import { createT, deepMerge, en, zhCN, type LocaleStrings } from '@betteroffice/docx-i18n';

const t = createT(deepMerge(en, zhCN) as LocaleStrings, 'zh-CN');
t('formattingBar.bold'); // '加粗'
```

## Documentation

- [JavaScript guide](https://docs.betteroffice.dev/docs/javascript)
- [`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react)
- [Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/docx-i18n/CHANGELOG.md)
