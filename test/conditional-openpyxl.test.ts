// Cross-check of the conditional format reader against openpyxl, an independent
// implementation. Skipped where python3 or openpyxl is not installed.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { ConditionalFormats } from '../src/format/conditional';
import { open } from '../src/index';
import { rangeName, type DifferentialFormat } from '../src/model/sheet/types';
import { buildXlsx, worksheetXml } from './fixtures';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS}">` +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<dxfs count="3">' +
  '<dxf><font><b/><i/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>' +
  '<dxf><font><strike/><u val="double"/></font><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/><bgColor rgb="FF0000FF"/></patternFill></fill>' +
  '<border><left style="thin"><color rgb="FFFF0000"/></left><bottom style="double"><color rgb="FF00FF00"/></bottom></border></dxf>' +
  '<dxf><numFmt numFmtId="164" formatCode="0.00"/></dxf>' +
  '</dxfs></styleSheet>';

const cf = (sqref: string, rules: string) => `<conditionalFormatting sqref="${sqref}">${rules}</conditionalFormatting>`;

const RULES =
  cf(
    'A1:A10',
    '<cfRule type="cellIs" dxfId="0" priority="1" operator="between" stopIfTrue="1"><formula>2</formula><formula>$C$1+1</formula></cfRule>' +
      '<cfRule type="cellIs" dxfId="1" priority="2" operator="notEqual"><formula>"abc"</formula></cfRule>' +
      '<cfRule type="expression" dxfId="2" priority="3"><formula>AND($A1&gt;5,B$1&lt;3)</formula></cfRule>',
  ) +
  cf(
    'B1:B10 D1:D3',
    '<cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule>' +
      '<cfRule type="dataBar" priority="5"><dataBar showValue="0" minLength="5" maxLength="95"><cfvo type="num" val="0"/><cfvo type="percent" val="80"/><color rgb="FF638EC6"/></dataBar></cfRule>' +
      '<cfRule type="iconSet" priority="6"><iconSet iconSet="4Arrows" showValue="0" reverse="1"><cfvo type="percent" val="0"/><cfvo type="percent" val="25"/><cfvo type="num" val="7" gte="0"/><cfvo type="formula" val="$C$2"/></iconSet></cfRule>',
  ) +
  cf(
    'C1:C10',
    '<cfRule type="top10" dxfId="0" priority="7" rank="3" percent="1" bottom="1"/>' +
      '<cfRule type="aboveAverage" dxfId="0" priority="8" aboveAverage="0" equalAverage="1"/>' +
      '<cfRule type="aboveAverage" dxfId="0" priority="9" stdDev="2"/>' +
      '<cfRule type="uniqueValues" dxfId="1" priority="10"/>' +
      '<cfRule type="duplicateValues" dxfId="1" priority="11"/>' +
      '<cfRule type="containsText" dxfId="0" priority="12" operator="containsText" text="ab"><formula>NOT(ISERROR(SEARCH("ab",C1)))</formula></cfRule>' +
      '<cfRule type="notContainsText" dxfId="0" priority="13" operator="notContains" text="ab"><formula>ISERROR(SEARCH("ab",C1))</formula></cfRule>' +
      '<cfRule type="beginsWith" dxfId="0" priority="14" operator="beginsWith" text="ab"><formula>LEFT(C1,2)="ab"</formula></cfRule>' +
      '<cfRule type="endsWith" dxfId="0" priority="15" operator="endsWith" text="ab"><formula>RIGHT(C1,2)="ab"</formula></cfRule>' +
      '<cfRule type="containsBlanks" dxfId="0" priority="16"><formula>LEN(TRIM(C1))=0</formula></cfRule>' +
      '<cfRule type="notContainsBlanks" dxfId="0" priority="17"><formula>LEN(TRIM(C1))&gt;0</formula></cfRule>' +
      '<cfRule type="containsErrors" dxfId="0" priority="18"><formula>ISERROR(C1)</formula></cfRule>' +
      '<cfRule type="notContainsErrors" dxfId="0" priority="19"><formula>NOT(ISERROR(C1))</formula></cfRule>' +
      '<cfRule type="timePeriod" dxfId="0" priority="20" timePeriod="lastWeek"><formula>AND(TODAY()-ROUNDDOWN(C1,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(C1,0)&lt;(WEEKDAY(TODAY())+7))</formula></cfRule>',
  );

// What openpyxl reads, in the shape the comparison below builds from our model.
const PYTHON = `
import json, sys
import openpyxl
wb = openpyxl.load_workbook(sys.argv[1])
ws = wb.active
def color(c):
    return None if c is None or c.rgb is None or not isinstance(c.rgb, str) else c.rgb
def number(x):
    # openpyxl reads "50" as the float 50.0; the file says 50.
    return str(int(x)) if isinstance(x, float) and x == int(x) else str(x)
def cfvo(v):
    return {'type': v.type, 'value': None if v.val is None else number(v.val), 'gte': v.gte is None or bool(v.gte)}
def side(s):
    return None if s is None or s.style is None else {'style': s.style, 'color': color(s.color)}
def dxf(d):
    if d is None:
        return None
    out = {}
    if d.font is not None:
        f = d.font
        out['font'] = {'bold': bool(f.b), 'italic': bool(f.i), 'strike': bool(f.strike), 'underline': f.u, 'color': color(f.color)}
    if d.fill is not None:
        out['fill'] = {'pattern': d.fill.patternType or 'solid', 'bg': color(d.fill.bgColor)}
    if d.border is not None:
        out['border'] = {'left': side(d.border.left), 'right': side(d.border.right), 'top': side(d.border.top), 'bottom': side(d.border.bottom)}
    if d.numFmt is not None:
        out['numFmt'] = d.numFmt.formatCode
    return out
formats = []
for cf in ws.conditional_formatting:
    rules = []
    for r in cf.rules:
        rule = {'type': r.type, 'priority': r.priority, 'stopIfTrue': bool(r.stopIfTrue), 'operator': r.operator,
                'formulas': [str(f) for f in r.formula], 'text': r.text, 'timePeriod': r.timePeriod, 'rank': r.rank,
                'percent': bool(r.percent), 'bottom': bool(r.bottom),
                'aboveAverage': r.aboveAverage is None or bool(r.aboveAverage), 'equalAverage': bool(r.equalAverage),
                'stdDev': r.stdDev, 'dxf': dxf(r.dxf)}
        if r.colorScale is not None:
            rule['colorScale'] = {'values': [cfvo(v) for v in r.colorScale.cfvo], 'colors': [color(c) for c in r.colorScale.color]}
        if r.dataBar is not None:
            b = r.dataBar
            rule['dataBar'] = {'values': [cfvo(v) for v in b.cfvo], 'color': color(b.color), 'showValue': b.showValue is None or bool(b.showValue),
                               'minLength': 10 if b.minLength is None else b.minLength, 'maxLength': 90 if b.maxLength is None else b.maxLength}
        if r.iconSet is not None:
            s = r.iconSet
            rule['iconSet'] = {'set': s.iconSet or '3TrafficLights1', 'values': [cfvo(v) for v in s.cfvo],
                               'showValue': s.showValue is None or bool(s.showValue), 'reverse': bool(s.reverse)}
        rules.append(rule)
    formats.append({'sqref': str(cf.sqref), 'rules': rules})
print(json.dumps(formats))
`;

const orNull = <T>(value: T | undefined): T | null => (value === undefined ? null : value);

function dxfView(format: DifferentialFormat | undefined): unknown {
  if (!format) return null;
  const out: Record<string, unknown> = {};
  if (format.font) {
    const font = format.font;
    out.font = { bold: font.bold === true, italic: font.italic === true, strike: font.strike === true, underline: orNull(font.underline), color: orNull(font.color && font.color.rgb) };
  }
  if (format.fill && format.fill.type === 'pattern') out.fill = { pattern: format.fill.pattern, bg: orNull(format.fill.bg && format.fill.bg.rgb) };
  if (format.border) {
    const side = (edge: { style: string; color?: { rgb?: string } } | undefined) => (edge ? { style: edge.style, color: orNull(edge.color && edge.color.rgb) } : null);
    out.border = { left: side(format.border.left), right: side(format.border.right), top: side(format.border.top), bottom: side(format.border.bottom) };
  }
  if (format.numFmt) out.numFmt = format.numFmt.code;
  return out;
}

test('conditional: the rules of a synthetic workbook read the same as openpyxl reads them', async (t) => {
  const probe = spawnSync('python3', ['-c', 'import openpyxl; print(openpyxl.__version__)'], { encoding: 'utf8' });
  if (probe.status !== 0) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const rows = Array.from({ length: 10 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}"><v>${r}</v></c><c r="B${r + 1}"><v>${r * 2}</v></c><c r="C${r + 1}"><v>${r % 3}</v></c></row>`).join('');
  const bytes = await buildXlsx({ sheets: [{ name: 'Rules', xml: worksheetXml(`<sheetData>${rows}</sheetData>${RULES}`) }], styles: STYLES });

  // A scratch folder inside the repository (ignored by git), removed afterwards.
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-cf-openpyxl-${process.pid}`);
  mkdirSync(folder, { recursive: true });
  let theirs: any[];
  try {
    const file = join(folder, 'rules.xlsx');
    writeFileSync(file, bytes);
    const run = spawnSync('python3', ['-c', PYTHON, file], { encoding: 'utf8' });
    assert.equal(run.status, 0, run.stderr);
    theirs = JSON.parse(run.stdout);
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }

  const workbook = await (await open(bytes)).workbook();
  const dxfs = workbook.styles.dxfFormats!;
  const ours = workbook.sheets[0].conditionalFormats!.map((format) => ({
    sqref: format.ranges.map(rangeName).join(' '),
    rules: format.rules.map((rule) => {
      const view: Record<string, unknown> = {
        type: rule.type,
        priority: rule.priority,
        stopIfTrue: rule.stopIfTrue,
        operator: orNull(rule.operator),
        formulas: rule.formulas,
        text: orNull(rule.text),
        timePeriod: orNull(rule.timePeriod),
        rank: orNull(rule.rank),
        percent: rule.percent,
        bottom: rule.bottom,
        aboveAverage: rule.aboveAverage,
        equalAverage: rule.equalAverage,
        stdDev: orNull(rule.stdDev),
        dxf: dxfView(rule.dxfId === undefined ? undefined : dxfs[rule.dxfId]),
      };
      const values = (list: { type: string; value?: string; gte: boolean }[]) => list.map((value) => ({ type: value.type, value: orNull(value.value), gte: value.gte }));
      if (rule.colorScale) view.colorScale = { values: values(rule.colorScale.values), colors: rule.colorScale.colors.map((color) => orNull(color.rgb)) };
      if (rule.dataBar) {
        const bar = rule.dataBar;
        view.dataBar = { values: values([bar.min, bar.max]), color: orNull(bar.color && bar.color.rgb), showValue: bar.showValue, minLength: bar.minLength, maxLength: bar.maxLength };
      }
      if (rule.iconSet) view.iconSet = { set: rule.iconSet.set, values: values(rule.iconSet.values), showValue: rule.iconSet.showValue, reverse: rule.iconSet.reverse };
      return view;
    }),
  }));

  assert.equal(ours.length, 3);
  assert.equal(ours.reduce((count, format) => count + format.rules.length, 0), 20);
  // openpyxl keeps one entry per range list, in an order of its own.
  const bySqref = (a: { sqref: string }, b: { sqref: string }) => (a.sqref < b.sqref ? -1 : a.sqref > b.sqref ? 1 : 0);
  assert.deepEqual(ours.sort(bySqref), theirs.sort(bySqref));
});

// A workbook written by openpyxl: another producer's way of laying the rules and their formats out.
const PYTHON_WRITE = `
import sys
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule, IconSetRule, Rule
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
wb = Workbook()
ws = wb.active
for r in range(1, 11):
    ws.cell(row=r, column=1, value=r)
    ws.cell(row=r, column=2, value=r * 10)
    ws.cell(row=r, column=3, value=11 - r)
    ws.cell(row=r, column=4, value='ok' if r % 2 else 'late')
    ws.cell(row=r, column=5, value=r % 4)
red = PatternFill(start_color='FFFFC7CE', end_color='FFFFC7CE', fill_type='solid')
ws.conditional_formatting.add('A1:A10', CellIsRule(operator='greaterThan', formula=['7'], fill=red, font=Font(bold=True, color='FF9C0006')))
ws.conditional_formatting.add('A1:A10', CellIsRule(operator='between', formula=['2', '3'], stopIfTrue=True, border=Border(left=Side(style='thin', color='FF0000FF'))))
ws.conditional_formatting.add('B1:B10', ColorScaleRule(start_type='min', start_color='FFF8696B', mid_type='percentile', mid_value=50, mid_color='FFFFEB84', end_type='max', end_color='FF63BE7B'))
ws.conditional_formatting.add('C1:C10', DataBarRule(start_type='num', start_value=0, end_type='max', color='FF638EC6', showValue=False))
ws.conditional_formatting.add('C1:C10', IconSetRule('3Arrows', 'percent', [0, 33, 67], showValue=None, reverse=True))
ws.conditional_formatting.add('A1:E10', FormulaRule(formula=['$D1="late"'], font=Font(italic=True)))
ws.conditional_formatting.add('E1:E10', Rule(type='duplicateValues', dxf=DifferentialStyle(fill=PatternFill(bgColor='FF00FF00'))))
ws.conditional_formatting.add('E1:E10', Rule(type='top10', rank=1, dxf=DifferentialStyle(font=Font(strike=True))))
wb.save(sys.argv[1])
`;

test('conditional: a workbook written by openpyxl is read and evaluated', async (t) => {
  const probe = spawnSync('python3', ['-c', 'import openpyxl'], { encoding: 'utf8' });
  if (probe.status !== 0) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-cf-openpyxl-w${process.pid}`);
  mkdirSync(folder, { recursive: true });
  let bytes: Uint8Array;
  try {
    const file = join(folder, 'written.xlsx');
    const run = spawnSync('python3', ['-c', PYTHON_WRITE, file], { encoding: 'utf8' });
    assert.equal(run.status, 0, run.stderr);
    bytes = new Uint8Array(readFileSync(file));
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }

  const workbook = await (await open(bytes)).workbook();
  const sheet = workbook.sheets[0];
  const rules = sheet.conditionalFormats!.reduce((all, format) => all.concat(format.rules), [] as NonNullable<typeof sheet.conditionalFormats>[number]['rules']);
  assert.deepEqual(rules.map((rule) => rule.type).sort(), ['cellIs', 'cellIs', 'colorScale', 'dataBar', 'duplicateValues', 'expression', 'iconSet', 'top10']);
  // Priorities are 1..8 in the order the rules were added.
  assert.deepEqual(rules.map((rule) => rule.priority).sort((a, b) => a - b), [1, 2, 3, 4, 5, 6, 7, 8]);
  for (const rule of rules) if (rule.dxfId !== undefined) assert.ok(workbook.styles.dxfFormats![rule.dxfId], `dxf ${rule.dxfId} of ${rule.type}`);

  // No engine is needed except for the one formula rule.
  const { CalcEngine } = await import('../src/calc/engine');
  const engine = new CalcEngine(workbook);
  const formats = new ConditionalFormats(workbook, sheet, { engine: () => engine });
  const at = (row: number, col: number) => formats.resolve(row - 1, col);

  // A: above 7 is pink, bold and dark red; 2..3 has the left border and stops. Even rows are "late": italic.
  assert.deepEqual(at(1, 0), undefined);
  assert.deepEqual(at(8, 0)!.style, { fill: '#FFC7CE', bold: true, color: '#9C0006', italic: true });
  assert.deepEqual(at(9, 0)!.style, { fill: '#FFC7CE', bold: true, color: '#9C0006' });
  assert.deepEqual(at(2, 0)!.style, { left: { style: 'thin', color: { rgb: 'FF0000FF' } } });
  assert.deepEqual(at(4, 0)!.style, { italic: true });
  // B: the three colours at the minimum, the median and the maximum.
  assert.equal(at(1, 1)!.fill, '#F8696B');
  assert.equal(at(10, 1)!.fill, '#63BE7B');
  // 50 is eight ninths of the way from the minimum (10) to the median (55).
  assert.equal(at(5, 1)!.fill, '#FEDD81');
  // C: 10 down to 1 - the bar from 0 to the maximum, no value shown; arrows reversed, so the largest value points down.
  assert.deepEqual(at(1, 2)!.bar, { from: 0, to: 0.9, color: '#638EC6', gradient: true, negative: false, rightToLeft: false });
  assert.equal(at(1, 2)!.hideValue, true);
  assert.deepEqual(at(1, 2)!.icon, { set: '3Arrows', index: 0 });
  assert.deepEqual(at(10, 2)!.icon, { set: '3Arrows', index: 2 });
  // E: 1, 2, 3, 0, 1, 2, 3, 0, 1, 2 - every value repeats (green); the 3s are the top value (struck through).
  assert.deepEqual(at(3, 4)!.style, { fill: '#00FF00', strike: true });
  assert.deepEqual(at(1, 4)!.style, { fill: '#00FF00' });
  assert.deepEqual(at(2, 4)!.style, { italic: true, fill: '#00FF00' });
});
