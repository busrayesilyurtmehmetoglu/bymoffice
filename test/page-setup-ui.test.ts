// What the dialog of the page setup hands over and what the bridge does with the page setup:
// only the fields that changed (a value shown rounded is never written back), lengths in the
// unit of the language, references as typed, and the messages setPageSetup, insertPageBreak,
// removePageBreak, resetPageBreaks and insertSectionBreak over a view that is faked here.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Encode, createBridge, type BridgeHost } from '../src/embed/bridge';
import type { OfficeFile } from '../src/index';
import { parsePrintArea, parseTitleBand, printAreaText, titleBandText, type PageSetupState } from '../src/ops/page-setup';
import { changedFields, formatTwips, parseTwips, parseWhole, type PageDialogGroup } from '../src/ui/page-setup-dialog';
import { pageSetupPatchOf } from '../src/ui/sheet-view';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

const STATE: PageSetupState = {
  paperSize: 9,
  paperStated: true,
  orientation: 'portrait',
  margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  fitToPage: false,
  scale: 100,
  fitToWidth: 1,
  fitToHeight: 1,
  horizontalCentered: false,
  verticalCentered: false,
  gridLines: false,
  headings: false,
  printArea: [],
  rowBreaks: [],
  colBreaks: [],
};

test('page setup dialog: lengths in the unit of the language, two decimals; typed back in twips', () => {
  assert.equal(formatTwips(0.7 * 1440, 'cm', ','), '1,78');
  assert.equal(formatTwips(0.75 * 1440, 'in', '.'), '0.75');
  assert.equal(formatTwips(1417, 'cm', ','), '2,5');
  assert.equal(parseTwips('2,5', 'cm'), 1417);
  assert.equal(parseTwips('1.5', 'in'), 2160);
  assert.equal(parseTwips('-1', 'in'), undefined);
  assert.equal(parseTwips('abc', 'cm'), undefined);
  assert.equal(parseWhole(' 80 '), 80);
  assert.equal(parseWhole('8.5'), undefined);
});

test('page setup dialog: only the fields that differ from what they showed are handed over', () => {
  const groups: PageDialogGroup[] = [
    {
      title: 'g',
      fields: [
        { key: 'left', label: 'Left', kind: 'text', value: '1,78' },
        { key: 'top', label: 'Top', kind: 'text', value: '1,91' },
        { key: 'grid', label: 'Grid', kind: 'check', value: false },
        { key: 'fixed', label: 'Fixed', kind: 'text', value: 'x', disabled: true },
      ],
    },
  ];
  assert.deepEqual(changedFields(groups, { left: '1,78', top: ' 1,91 ', grid: false, fixed: 'y' }), {});
  assert.deepEqual(changedFields(groups, { left: '2', top: '1,91', grid: true, fixed: 'y' }), { left: '2', grid: true });
});

test('page setup dialog: the patch of a sheet - untouched rounded margins stay, the scale goes with its mode, references are read', () => {
  const tr = STRINGS_TR.pageSetup;
  assert.deepEqual(pageSetupPatchOf({}, STATE, tr), {});
  assert.deepEqual(pageSetupPatchOf({ left: '2,5' }, STATE, tr), { margins: { left: 0.984252 } });
  assert.deepEqual(pageSetupPatchOf({ left: 'x' }, STATE, tr), { message: tr.badLength, field: 'left' });
  assert.deepEqual(pageSetupPatchOf({ scaling: 'fit' }, STATE, tr), { fitTo: { width: 1, height: 1 } });
  assert.deepEqual(pageSetupPatchOf({ scaling: 'fit', fitTall: '0' }, STATE, tr), { fitTo: { width: 1, height: 0 } });
  assert.deepEqual(pageSetupPatchOf({ scale: '75' }, STATE, tr), { scale: 75 });
  assert.deepEqual(pageSetupPatchOf({ scale: '5' }, STATE, tr), { message: tr.badScale, field: 'scale' });
  // A scale typed while "fit to pages" is chosen says nothing.
  assert.deepEqual(pageSetupPatchOf({ scale: '75' }, { ...STATE, fitToPage: true }, tr), {});
  assert.deepEqual(pageSetupPatchOf({ paper: '1', orientation: 'landscape', gridLines: true }, STATE, tr), { paperSize: 1, orientation: 'landscape', gridLines: true });
  assert.deepEqual(pageSetupPatchOf({ printArea: 'A1:D20;F1:G5' }, STATE, tr), {
    printArea: [
      { top: 0, left: 0, bottom: 19, right: 3 },
      { top: 0, left: 5, bottom: 4, right: 6 },
    ],
  });
  assert.deepEqual(pageSetupPatchOf({ printArea: '' }, STATE, tr), { printArea: null });
  assert.deepEqual(pageSetupPatchOf({ printArea: 'A1:' }, STATE, tr), { message: tr.badArea, field: 'printArea' });
  assert.deepEqual(pageSetupPatchOf({ titleRows: '$1:$2', titleCols: 'A' }, STATE, tr), { titleRows: { first: 0, last: 1 }, titleCols: { first: 0, last: 0 } });
  assert.deepEqual(pageSetupPatchOf({ titleRows: 'A:B' }, STATE, tr), { message: tr.badRows, field: 'titleRows' });
  assert.equal(STRINGS_EN.pageSetup.unit, 'in');
});

test('page setup: references as typed and as shown', () => {
  assert.deepEqual(parsePrintArea("'My, sheet'!$A$1:$B$2,Sheet1!C3"), [
    { top: 0, left: 0, bottom: 1, right: 1 },
    { top: 2, left: 2, bottom: 2, right: 2 },
  ]);
  assert.deepEqual(parsePrintArea('1:3'), [{ top: 0, left: 0, bottom: 2, right: 16383 }]);
  assert.deepEqual(parsePrintArea('B:C'), [{ top: 0, left: 1, bottom: 1048575, right: 2 }]);
  assert.equal(parsePrintArea('XFE1'), undefined);
  assert.equal(parsePrintArea('A0'), undefined);
  assert.equal(parsePrintArea(''), null);
  assert.equal(parseTitleBand('3:1', 'row')!.first, 0);
  assert.equal(parseTitleBand('A1', 'row'), undefined);
  assert.equal(printAreaText([{ top: 0, left: 0, bottom: 19, right: 3 }, { top: 0, left: 0, bottom: 1048575, right: 1 }], ';'), '$A$1:$D$20;$A:$B');
  assert.equal(titleBandText({ first: 0, last: 1 }, 'row'), '$1:$2');
  assert.equal(titleBandText({ first: 2, last: 2 }, 'col'), '$C:$C');
});

const DATA = base64Encode(new Uint8Array([80, 75, 3, 4]));

function sheetBridge() {
  const sent: any[] = [];
  const calls: unknown[][] = [];
  let readOnly = false;
  const view: any = {
    history: { canUndo: false, canRedo: false },
    activeSheet: 0,
    get readOnly() {
      return readOnly;
    },
    isEditing: false,
    workbook: { sheets: [{ name: 'A', state: 'visible', kind: 'worksheet' }, { name: 'B', state: 'visible', kind: 'worksheet' }] },
    commitEditing: () => undefined,
    setReadOnly: (value: boolean) => void (readOnly = value),
    setLocale: () => undefined,
    localizeError: (error: unknown) => `Localized: ${(error as Error).message}`,
    destroy: () => undefined,
    setPageSetup(patch: unknown, sheet: number) {
      calls.push(['setPageSetup', patch, sheet]);
      if ((patch as { scale?: number }).scale === 5) throw new RangeError('The scale is a whole percent from 10 to 400');
      return true;
    },
    insertPageBreak: (at: unknown, sheet: number) => (calls.push(['insertPageBreak', at, sheet]), true),
    removePageBreak: (at: unknown, sheet: number) => (calls.push(['removePageBreak', at, sheet]), false),
    resetPageBreaks: (sheet: number) => (calls.push(['resetPageBreaks', sheet]), true),
    pageSetup: (sheet: number) => ({ sheet, orientation: 'landscape' }),
  };
  const host: BridgeHost = {
    container: {} as HTMLElement,
    version: '9.9.9',
    post(json) {
      sent.push(JSON.parse(json));
      return true;
    },
    open: async () => ({ format: 'xlsx', workbook: async () => view.workbook, save: async () => new Uint8Array([1]) }) as unknown as OfficeFile,
    createView: () => view,
    frame: () => undefined,
    timer: () => undefined,
  };
  return { bridge: createBridge(host), sent, calls, setReadOnly: (value: boolean) => void (readOnly = value) };
}

test('bridge page setup: a sheet - the patch as the message says it, breaks, and what is refused', async () => {
  const f = sheetBridge();
  await f.bridge.receive({ type: 'load', id: 'l', data: DATA });
  f.sent.length = 0;
  await f.bridge.receive({ type: 'setPageSetup', id: 'p', orientation: 'landscape', margins: { left: 1 }, printArea: 'A1:B2', titleRows: null, gridLines: true, index: 1 });
  assert.deepEqual(f.calls.shift(), ['setPageSetup', { orientation: 'landscape', margins: { left: 1 }, gridLines: true, printArea: [{ top: 0, left: 0, bottom: 1, right: 1 }], titleRows: null }, 1]);
  assert.deepEqual(f.sent.filter((message) => message.type === 'done'), [{ type: 'done', id: 'p', command: 'setPageSetup', changed: true, pageSetup: { sheet: 1, orientation: 'landscape' } }]);
  f.sent.length = 0;
  await f.bridge.receive({ type: 'insertPageBreak', id: 'b', row: 4 });
  await f.bridge.receive({ type: 'removePageBreak', id: 'c' });
  await f.bridge.receive({ type: 'resetPageBreaks', id: 'd' });
  assert.deepEqual(f.calls.splice(0), [
    ['insertPageBreak', { row: 4, col: undefined }, 0],
    ['removePageBreak', undefined, 0],
    ['resetPageBreaks', 0],
  ]);
  assert.deepEqual(
    f.sent.filter((message) => message.type === 'done').map((message) => [message.id, message.changed]),
    [
      ['b', true],
      ['c', false],
      ['d', true],
    ],
  );
  f.sent.length = 0;
  await f.bridge.receive({ type: 'setPageSetup', id: 'x', scale: 5 });
  await f.bridge.receive({ type: 'setPageSetup', id: 'y', orientation: 'sideways' });
  await f.bridge.receive({ type: 'setPageSetup', id: 'z', printArea: 'not a range' });
  await f.bridge.receive({ type: 'insertPageBreak', id: 'w', index: 1 });
  await f.bridge.receive({ type: 'insertSectionBreak', id: 'v' });
  await f.bridge.receive({ type: 'setPageSetup', id: 'u', index: 7 });
  assert.deepEqual(
    f.sent.map((message) => [message.id, message.code]),
    [
      ['x', 'page-setup-refused'],
      ['y', 'bad-message'],
      ['z', 'bad-message'],
      ['w', 'bad-message'],
      ['v', 'no-document'],
      ['u', 'no-sheet'],
    ],
  );
  assert.equal(f.sent[0].message, 'Localized: The scale is a whole percent from 10 to 400');
  f.setReadOnly(true);
  f.sent.length = 0;
  await f.bridge.receive({ type: 'setPageSetup', id: 'r', orientation: 'portrait' });
  assert.deepEqual(f.sent, [{ type: 'error', id: 'r', code: 'read-only', message: 'The file is read-only' }]);
});

test('bridge page setup: a document - sections by paper name and twips, section breaks', async () => {
  const sent: any[] = [];
  const calls: unknown[][] = [];
  const pages: any = {
    destroy: () => undefined,
    setSectionSetup: (patch: unknown, scope: string) => (calls.push(['setSectionSetup', patch, scope]), true),
    insertSectionBreak: (kind: string) => (calls.push(['insertSectionBreak', kind]), true),
    get sectionSetup() {
      return { width: 11906, height: 16838, section: 0, sections: 2 };
    },
  };
  const host: BridgeHost = {
    container: {} as HTMLElement,
    version: '9.9.9',
    post(json) {
      sent.push(JSON.parse(json));
      return true;
    },
    open: async () => ({ format: 'docx', kind: 'document', document: async () => ({}), save: async () => new Uint8Array([1]) }) as unknown as OfficeFile,
    createView: () => {
      throw new Error('no sheet');
    },
    createDocumentView: () => pages,
    frame: () => undefined,
    timer: () => undefined,
  };
  const bridge = createBridge(host);
  await bridge.receive({ type: 'load', data: DATA });
  sent.length = 0;
  await bridge.receive({ type: 'setPageSetup', id: 'a', paper: 'a4', orientation: 'landscape', margins: { top: 1440 }, scope: 'document' });
  await bridge.receive({ type: 'insertSectionBreak', id: 'b', kind: 'continuous' });
  await bridge.receive({ type: 'insertSectionBreak', id: 'c' });
  assert.deepEqual(calls, [
    ['setSectionSetup', { paper: 'a4', orientation: 'landscape', margins: { top: 1440 } }, 'document'],
    ['insertSectionBreak', 'continuous'],
    ['insertSectionBreak', 'nextPage'],
  ]);
  assert.deepEqual(sent.filter((message) => message.type === 'done')[0], { type: 'done', id: 'a', command: 'setPageSetup', changed: true, pageSetup: { width: 11906, height: 16838, section: 0, sections: 2 } });
  sent.length = 0;
  await bridge.receive({ type: 'insertSectionBreak', id: 'd', kind: 'sideways' });
  await bridge.receive({ type: 'setPageSetup', id: 'e', scope: 'page' });
  await bridge.receive({ type: 'insertPageBreak', id: 'f' });
  await bridge.receive({ type: 'setPageSetup', id: 'g', margins: { top: 'wide' } });
  assert.deepEqual(
    sent.map((message) => [message.id, message.code]),
    [
      ['d', 'bad-message'],
      ['e', 'bad-message'],
      ['f', 'no-sheet'],
      ['g', 'bad-message'],
    ],
  );
});
