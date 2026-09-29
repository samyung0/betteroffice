import type { Run } from './run';

export interface RawXml {
  type: 'rawXml';
  xml: string;
  /** The runs it shows in a field that keeps it in place of a tracked change. */
  shown?: Run[];
}

export function isRawXml(value: { type: string }): value is RawXml {
  return value.type === 'rawXml';
}
