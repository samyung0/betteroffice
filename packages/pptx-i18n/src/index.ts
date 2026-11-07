/**
 * Shared locale data, types, and runtime helpers for @betteroffice/pptx-react.
 *
 * @packageDocumentation
 * @public
 */

import { editorEn } from './editor';
import { proposalEn } from './proposals';
import { menuEn } from './menus';
import { deepMerge } from './runtime';
import arJson from '../ar.json';
import esJson from '../es.json';
import itJson from '../it.json';
import jaJson from '../ja.json';
import koJson from '../ko.json';
import nlJson from '../nl.json';
import ruJson from '../ru.json';
import ukJson from '../uk.json';
import viJson from '../vi.json';
import zhCNJson from '../zh-CN.json';
import zhTWJson from '../zh-TW.json';

export { createT, deepMerge } from './runtime';
/** @internal */
export { editorEn } from './editor';
/** @internal */
export { proposalEn } from './proposals';
/** @internal */
export { menuEn } from './menus';

export type LocaleStrings = typeof en;
export type LocaleCode =
  | 'en'
  | 'ar'
  | 'es'
  | 'it'
  | 'ja'
  | 'ko'
  | 'nl'
  | 'ru'
  | 'uk'
  | 'vi'
  | 'zh-CN'
  | 'zh-TW';

export const en = /* @__PURE__ */ deepMerge(deepMerge(editorEn, proposalEn), menuEn) as typeof editorEn &
  typeof proposalEn & typeof menuEn;
export const ar: PartialLocaleStrings = arJson;
export const es: PartialLocaleStrings = esJson;
export const it: PartialLocaleStrings = itJson;
export const ja: PartialLocaleStrings = jaJson;
export const ko: PartialLocaleStrings = koJson;
export const nl: PartialLocaleStrings = nlJson;
export const ru: PartialLocaleStrings = ruJson;
export const uk: PartialLocaleStrings = ukJson;
export const vi: PartialLocaleStrings = viJson;
export const zhCN: PartialLocaleStrings = zhCNJson;
export const zhTW: PartialLocaleStrings = zhTWJson;

export const locales: Record<LocaleCode, PartialLocaleStrings> = {
  en,
  ar,
  es,
  it,
  ja,
  ko,
  nl,
  ru,
  uk,
  vi,
  'zh-CN': zhCN,
  'zh-TW': zhTW,
};

export type DeepPartial<T> = {
  [K in keyof T]?: T[K] extends object ? DeepPartial<T[K]> : T[K] | null;
};

export type PartialLocaleStrings = DeepPartial<LocaleStrings> & {
  _lang?: LocaleCode | (string & {});
  editor?: DeepPartial<LocaleStrings['editor']> & { textInput?: string | null; selectionAnnouncement?: string | null };
};

export type Translations = PartialLocaleStrings;

type DotPath<T, Prefix extends string = ''> = {
  [K in keyof T & string]: T[K] extends Record<string, unknown>
    ? DotPath<T[K], `${Prefix}${K}.`>
    : `${Prefix}${K}`;
}[keyof T & string];

export type TranslationKey = DotPath<LocaleStrings>;

export type TFunction = (
  key: TranslationKey,
  vars?: Record<string, string | number>
) => string;
