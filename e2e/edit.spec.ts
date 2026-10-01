import { XLSX_PARTS, readPartsInNode, xlsx } from './fixtures/files';
import { expect, gotoHarness, test, type Flavour } from './helpers';

const EDITED = '/xl/worksheets/sheet1.xml';
const NEW_CONTENT =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>İşlem ğüşöç €</t></is></c></row></sheetData></worksheet>';

for (const flavour of ['script', 'esm'] as Flavour[]) {
  test(`edit round trip (${flavour}): writePart changes one part and leaves the rest byte-identical`, async ({ page }) => {
    const input = await xlsx();
    const original = await readPartsInNode(input);
    await gotoHarness(page, flavour);

    const result = await page.evaluate(
      async ({ bytes, part, content }) => {
        const api = (window as any).__bym;
        const file = await api.open(new Uint8Array(bytes));
        file.package.writePart(part, new TextEncoder().encode(content));
        const saved: Uint8Array = await file.save();

        const reopened = await api.open(saved);
        const parts: Record<string, number[]> = {};
        for (const name of reopened.package.partNames()) parts[name] = Array.from(await reopened.package.readPart(name));
        return {
          saved: Array.from(saved),
          format: reopened.format,
          names: reopened.package.partNames(),
          parts,
          editedText: new TextDecoder().decode(await reopened.package.readPart(part)),
        };
      },
      { bytes: Array.from(input), part: EDITED, content: NEW_CONTENT },
    );

    expect(result.format).toBe('xlsx');
    expect(result.names).toEqual(XLSX_PARTS);
    expect(result.editedText).toBe(NEW_CONTENT);
    expect(result.parts[EDITED]).not.toEqual(original[EDITED]);
    for (const name of XLSX_PARTS.filter((part) => part !== EDITED)) {
      expect(result.parts[name], `${name} must be untouched`).toEqual(original[name]);
    }

    // Same verdict when the saved bytes are opened outside the browser.
    const inNode = await readPartsInNode(Uint8Array.from(result.saved));
    expect(new TextDecoder().decode(Uint8Array.from(inNode[EDITED]))).toBe(NEW_CONTENT);
    for (const name of XLSX_PARTS.filter((part) => part !== EDITED)) expect(inNode[name]).toEqual(original[name]);
  });
}
