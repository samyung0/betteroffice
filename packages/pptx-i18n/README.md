# @betteroffice/pptx-i18n

UI strings, locale types and translation helpers for the
[`@betteroffice/pptx-react`](https://www.npmjs.com/package/@betteroffice/pptx-react) editor. `en` defines
the strings; the other locales override them.

## Install

```bash
bun add @betteroffice/pptx-i18n
```

## Usage

Pass a locale to the editor's `i18n` prop:

```tsx
import { zhCN } from '@betteroffice/pptx-i18n';
import { PptxEditor, type PptxAsyncEditorProps } from '@betteroffice/pptx-react';

export function Editor({ file, fonts }: Pick<PptxAsyncEditorProps, 'file' | 'fonts'>) {
  return <PptxEditor file={file} fonts={fonts} i18n={zhCN} />;
}
```

Override individual strings by spreading a locale. Keys that are missing or `null` in an override fall
back to English.

```ts
const custom = { ...zhCN, toolbar: { ...zhCN.toolbar, addSlide: '添加幻灯片' } };
```

## Locales

`en` (source) and `zh-CN`, exported as `zhCN`. Each locale is also its own subpath,
`@betteroffice/pptx-i18n/zh-CN`, so an app that picks the locale at runtime can load only the one it
needs. `locales` maps both tags to their strings.

## Helpers

The package exports the types `LocaleStrings`, `PartialLocaleStrings`, `Translations`,
`TranslationKey`, `LocaleCode` and `TFunction`, and the functions `createT` and `deepMerge`.
Outside React, build a typed `t()` directly:

```ts
import { createT, deepMerge, en, zhCN, type LocaleStrings } from '@betteroffice/pptx-i18n';

const t = createT(deepMerge(en, zhCN) as LocaleStrings, 'zh-CN');
t('toolbar.addSlide'); // '添加幻灯片'
```

## Documentation

- [JavaScript guide](https://docs.betteroffice.dev/docs/javascript)
- [`@betteroffice/pptx-react`](https://www.npmjs.com/package/@betteroffice/pptx-react)
- [Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/pptx-i18n/CHANGELOG.md)
