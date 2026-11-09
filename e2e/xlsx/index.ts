import { bulkPasteAndFormats } from './bulk-paste-and-formats';
import { editingSession } from './editing-session';
import { formulaChainCascade } from './formula-chain-cascade';
import { proposalsReview } from './proposals-review';
import { saveLoadCycles } from './save-load-cycles';
import { structuralStorm } from './structural-storm';
import { threeEditorsMesh } from './three-editors-mesh';
import { twoEditorsConverge } from './two-editors-converge';
import { incrementalCommits } from './incremental-commits';
import { viewportScrollRender } from './viewport-scroll-render';
import type { XlsxScenario } from './context';

export const scenarios: XlsxScenario[] = [
  editingSession,
  formulaChainCascade,
  bulkPasteAndFormats,
  structuralStorm,
  incrementalCommits,
  viewportScrollRender,
  twoEditorsConverge,
  threeEditorsMesh,
  proposalsReview,
  saveLoadCycles,
];
