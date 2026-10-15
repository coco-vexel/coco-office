/* eslint-disable no-console -- CLI verification output. */
import assert from 'node:assert/strict';
import { fromFlatOpc, toZip, fromZip, toFlatOpc } from '@coco-vexel/coco-office/pack';
import { flatOpcToOnlyOfficeWriteback } from '@coco-vexel/coco-office-oojson';
import { runOojsonSample } from './oojson-sample.mjs';
const result = runOojsonSample('OOJSON package consumer <&> 中文');
const pkg = fromFlatOpc(result.flatOpc);
const reloaded = await fromZip(await toZip(pkg));
const writeback = flatOpcToOnlyOfficeWriteback(toFlatOpc(reloaded));
assert(writeback.bodyJson, 'writeback must include a body document');
assert(JSON.stringify(writeback.bodyJson).includes(result.text));
console.log('OOJSON: JSON/XML round-trip, Core ZIP round-trip and writeback body verified');
