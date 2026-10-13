// Pictures and shapes of an RTF file: the bytes of a picture from hex digits
// or from "\bin", its size and cropping, shapes with their place, their
// picture and the text of their text box.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { documentPictures } from '../src/model/doc/pictures';
import { readRtf } from '../src/formats/rtf/document';
import { DocLayout } from '../src/layout/doc/layout';
import { paragraphText, type DocModel, type DrawingItem, type Paragraph } from '../src/model/doc/types';
import { fakeMeasurer } from './doc-layout-fixtures';
import { bytesOf, doc, drawingsOf, fakePng, hexOf, itemsOf, paragraphs, paragraphsOf, read, texts } from './rtf-fixtures';

const PNG = fakePng(40, 20, 16);
const TWIP = 635;

const firstDrawing = (model: DocModel): DrawingItem => drawingsOf(paragraphs(model)[0])[0];
const mediaOf = (model: DocModel, drawing: DrawingItem): Uint8Array => model.media!.get(drawing.image!.target!)!.bytes;

test('a picture given as hex digits is an inline drawing with its bytes in the media of the model', () => {
  const model = read(doc(String.raw`a{\pict\pngblip\picw40\pich20\picwgoal600\pichgoal300 ` + hexOf(PNG) + String.raw`}b\par`));
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'ab');
  const drawing = firstDrawing(model);
  assert.equal(drawing.placement, 'inline');
  assert.equal(drawing.kind, 'picture');
  assert.equal(drawing.width, 600 * TWIP);
  assert.equal(drawing.height, 300 * TWIP);
  assert.equal(model.media!.size, 1);
  const media = model.media!.get(drawing.image!.target!)!;
  assert.equal(media.type, 'image/png');
  assert.deepEqual(Array.from(media.bytes), Array.from(PNG));
  assert.deepEqual(drawing.images, [{ image: drawing.image }]);
  assert.match(drawing.image!.target!, /\.png$/);
});

test('the view finds the picture through the model', async () => {
  const model = read(doc(String.raw`{\pict\pngblip\picwgoal600\pichgoal300 ` + hexOf(PNG) + String.raw`}\par`));
  const source = documentPictures(model)!;
  assert.deepEqual(Array.from(await source.readPart(firstDrawing(model).image!.target!)), Array.from(PNG));
});

test('hex digits may be split by line ends, spaces and groups, in either case', () => {
  const hex = hexOf(PNG).toUpperCase();
  const split = hex.slice(0, 7) + '\r\n' + hex.slice(7, 30) + ' \n ' + hex.slice(30, 41) + '{\\*\\blipuid 00112233445566778899aabbccddeeff}' + hex.slice(41);
  const model = read(doc(String.raw`{\pict\pngblip\picwgoal600\pichgoal300\bliptag-123{\*\picprop\shplid1025{\sp{\sn shapeType}{\sv 75}}}` + split + String.raw`}\par`));
  assert.deepEqual(Array.from(mediaOf(model, firstDrawing(model))), Array.from(PNG));
});

test('a picture cut into thousands of pieces is put together again', () => {
  const large = fakePng(64, 64, 40000);
  const hex = hexOf(large);
  // A group after every third digit, so that pieces end in the middle of a byte.
  let source = '';
  for (let i = 0; i < hex.length; i += 3) source += hex.slice(i, i + 3) + '{}';
  const model = read(doc(String.raw`{\pict\pngblip\picwgoal600\pichgoal300 ` + source + String.raw`}\par`));
  const bytes = mediaOf(model, firstDrawing(model));
  assert.equal(bytes.length, large.length);
  assert.deepEqual(Array.from(bytes), Array.from(large));
});

test('a picture given as bytes after "\\bin"', () => {
  let binary = '';
  for (let i = 0; i < PNG.length; i++) binary += String.fromCharCode(PNG[i]);
  // The bytes hold braces and backslashes that are not RTF.
  const data = binary + '}{\\par';
  const source = '{\\rtf1 a{\\pict\\pngblip\\picwgoal600\\pichgoal300\\bin' + data.length + ' ' + data + '}b\\par}';
  const model = readRtf(bytesOf(source));
  assert.deepEqual(texts(model), ['ab']);
  const bytes = mediaOf(model, firstDrawing(model));
  assert.equal(bytes.length, data.length);
  assert.deepEqual(Array.from(bytes.subarray(0, PNG.length)), Array.from(PNG));
});

test('the kinds of pictures', () => {
  const kind = (words: string, bytes: Uint8Array = PNG) => {
    const model = read(doc('{\\pict' + words + '\\picwgoal600\\pichgoal300 ' + hexOf(bytes) + '}\\par'));
    const drawing = firstDrawing(model);
    return drawing.image ? model.media!.get(drawing.image.target!)!.type + ' ' + drawing.image.target!.replace(/^.*\./, '') : 'none';
  };
  assert.equal(kind('\\jpegblip'), 'image/jpeg jpg');
  assert.equal(kind('\\emfblip'), 'image/x-emf emf');
  assert.equal(kind('\\wmetafile8'), 'image/x-wmf wmf');
  assert.equal(kind('\\macpict'), 'image/pict pict');
  // A kind nothing can show, and a picture that says nothing about its kind: a drawing of its size without an image.
  assert.equal(kind('\\wbitmap0'), 'none');
  assert.equal(kind(''), 'none');
});

test('a device independent bitmap gets the file header it lacks', () => {
  const dib = new Uint8Array(40 + 4);
  dib[0] = 40;
  dib[4] = 1;
  dib[8] = 1;
  dib[12] = 1;
  dib[14] = 32;
  const model = read(doc('{\\pict\\dibitmap0\\picwgoal15\\pichgoal15 ' + hexOf(dib) + '}\\par'));
  const media = model.media!.get(firstDrawing(model).image!.target!)!;
  assert.equal(media.type, 'image/bmp');
  assert.equal(media.bytes.length, 14 + dib.length);
  assert.equal(media.bytes[0], 0x42);
  assert.equal(media.bytes[1], 0x4d);
  assert.equal(media.bytes[10], 54);
});

test('size: the goal times the scale, less what is cropped', () => {
  const model = read(doc(String.raw`{\pict\pngblip\picwgoal4000\pichgoal2000\picscalex50\picscaley25\piccropl400\piccropr600\piccropt200 ` + hexOf(PNG) + String.raw`}\par`));
  const drawing = firstDrawing(model);
  assert.equal(drawing.width, ((4000 - 400 - 600) * 50 * TWIP) / 100);
  assert.equal(drawing.height, ((2000 - 200) * 25 * TWIP) / 100);
  assert.deepEqual(drawing.crop, { left: 10000, right: 15000, top: 10000 });
  assert.deepEqual(drawing.images, [{ image: drawing.image, crop: drawing.crop }]);
});

test('without a goal the size is that of the picture: pixels of a bitmap, hundredths of a millimetre of a metafile', () => {
  const bitmap = firstDrawing(read(doc(String.raw`{\pict\pngblip\picw100\pich50 ` + hexOf(PNG) + String.raw`}\par`)));
  assert.equal(bitmap.width, 100 * 15 * TWIP);
  assert.equal(bitmap.height, 50 * 15 * TWIP);
  const metafile = firstDrawing(read(doc(String.raw`{\pict\wmetafile8\picw2540\pich1270 ` + hexOf(PNG) + String.raw`}\par`)));
  assert.equal(metafile.width, 1440 * TWIP);
  assert.equal(metafile.height, 720 * TWIP);
  // A bitmap that states no size at all says it in its own header.
  const header = firstDrawing(read(doc(String.raw`{\pict\pngblip ` + hexOf(PNG) + String.raw`}\par`)));
  assert.equal(header.width, 40 * 15 * TWIP);
  assert.equal(header.height, 20 * 15 * TWIP);
});

test('a size no picture has is cut to what Word allows, and one axis gone wrong follows the other', () => {
  const huge = firstDrawing(read(doc(String.raw`{\pict\pngblip\picwgoal2000000000\pichgoal2000000000\picscalex100000\picscaley100000 ` + hexOf(PNG) + String.raw`}\par`)));
  assert.equal(huge.width, 31680 * TWIP);
  assert.equal(huge.height, 31680 * TWIP);
  const skewed = firstDrawing(read(doc(String.raw`{\pict\pngblip\picwgoal4000\pichgoal2000\picscalex100000\picscaley100 ` + hexOf(PNG) + String.raw`}\par`)));
  assert.equal(skewed.height, 2000 * TWIP);
  assert.equal(skewed.width, 4000 * TWIP);
  const negative = firstDrawing(read(doc(String.raw`{\pict\pngblip\picwgoal-500\pichgoal-500\picscalex-5 ` + hexOf(PNG) + String.raw`}\par`)));
  assert.ok(negative.width! >= 0 && negative.height! >= 0);
});

test('the same picture twice is kept once', () => {
  const picture = String.raw`{\pict\pngblip\picwgoal600\pichgoal300 ` + hexOf(PNG) + '}';
  const other = String.raw`{\pict\pngblip\picwgoal600\pichgoal300 ` + hexOf(fakePng(41, 20, 16)) + '}';
  const model = read(doc(picture + picture + other + String.raw`\par{\header\pard ` + picture + String.raw`\par}`));
  assert.equal(model.media!.size, 2);
  const [a, b, c] = drawingsOf(paragraphs(model)[0]);
  assert.equal(a.image!.target, b.image!.target);
  assert.notEqual(a.image!.target, c.image!.target);
  const inHeader = drawingsOf(paragraphsOf(model.headers.get('hdr:0:default')!.blocks)[0])[0];
  assert.equal(inHeader.image!.target, a.image!.target);
});

test('of the two forms Word writes a picture in, the one for new readers is read', () => {
  const body = String.raw`{\*\shppict{\pict\pngblip\picwgoal600\pichgoal300 ` + hexOf(PNG) + String.raw`}}{\nonshppict{\pict\wmetafile8\picwgoal600\pichgoal300 010009000003}}\par`;
  const model = read(doc(body));
  assert.equal(drawingsOf(paragraphs(model)[0]).length, 1);
  assert.equal(model.media!.size, 1);
  assert.equal(model.media!.get(firstDrawing(model).image!.target!)!.type, 'image/png');
});

test('a picture without bytes and without a size is nothing; one with a size is a drawing without an image', () => {
  const empty = read(doc(String.raw`a{\pict\pngblip}b\par`));
  assert.equal(drawingsOf(paragraphs(empty)[0]).length, 0);
  assert.equal(empty.media, undefined);
  const sized = firstDrawing(read(doc(String.raw`{\pict\pngblip\picwgoal600\pichgoal300}\par`)));
  assert.equal(sized.image, undefined);
  assert.equal(sized.width, 600 * TWIP);
});

test('an odd number of hex digits loses the last one, and other characters are passed over', () => {
  const model = read(doc(String.raw`{\pict\pngblip\picwgoal600\pichgoal300 ` + hexOf(PNG) + String.raw`zz-7}\par`));
  assert.deepEqual(Array.from(mediaOf(model, firstDrawing(model))), Array.from(PNG));
});

// A shape as Word writes it: where it sits, its properties, the text of its text box, and what old readers show instead.
const shape = (inner: string, words = String.raw`\shpleft1000\shptop2000\shpright4000\shpbottom3500\shpfhdr0\shpbxcolumn\shpbypara\shpwr3\shpwrk0\shpfblwtxt0\shpz0\shplid1026`) =>
  String.raw`{\shp{\*\shpinst` + words + inner + String.raw`}{\shprslt{\*\do\dobxcolumn\dobypara\dodhgt8192\dptxbx{\dptxbxtext\pard\plain FALLBACK\par}\dpx1000\dpy2000\dpxsize3000\dpysize1500}}}`;

test('a text box: an anchored drawing with the blocks of its text', () => {
  const inner = String.raw`{\sp{\sn shapeType}{\sv 202}}{\sp{\sn dxTextLeft}{\sv 91440}}{\sp{\sn dyTextTop}{\sv 45720}}{\sp{\sn wzName}{\sv Text Box 2}}{\shptxt\pard\plain\qc Inside\par second {\b bold}\par}`;
  const model = read(doc('before' + shape(inner) + String.raw`after\par`));
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'beforeafter');
  const drawing = drawingsOf(p)[0];
  assert.equal(drawing.placement, 'anchor');
  assert.equal(drawing.kind, 'shape');
  assert.equal(drawing.name, 'Text Box 2');
  assert.equal(drawing.width, 3000 * TWIP);
  assert.equal(drawing.height, 1500 * TWIP);
  assert.deepEqual(drawing.anchor, {
    horizontal: { relativeFrom: 'column', offset: 1000 * TWIP },
    vertical: { relativeFrom: 'paragraph', offset: 2000 * TWIP },
    wrap: 'none',
  });
  assert.equal(drawing.textBoxes!.length, 1);
  const inside = paragraphsOf(drawing.textBoxes![0]);
  assert.deepEqual(inside.map(paragraphText), ['Inside', 'second bold']);
  assert.equal(inside[0].props.align, 'center');
  // The insets are where layout looks for them.
  const body = drawing.xml.children[0];
  assert.ok(body.type === 'element' && body.attrs.get('lIns') === '91440' && body.attrs.get('tIns') === '45720');
  // What old readers show instead is not shown a second time.
  assert.ok(JSON.stringify(model.body.blocks).indexOf('FALLBACK') < 0);
});

test('where a shape sits: what it is measured from, alignments, wrapping, behind the text', () => {
  const words = String.raw`\shpleft-200\shptop300\shpright2800\shpbottom1300\shpbxpage\shpbxignore\shpbymargin\shpbyignore\shpwr2\shpwrk1\shpfblwtxt1`;
  const inner = String.raw`{\sp{\sn posrelh}{\sv 0}}{\sp{\sn posh}{\sv 2}}{\sp{\sn posrelv}{\sv 1}}{\sp{\sn dxWrapDistLeft}{\sv 114300}}{\sp{\sn rotation}{\sv 5898240}}{\sp{\sn fFlipH}{\sv 1}}{\sp{\sn wzDescription}{\sv A shape}}`;
  const drawing = firstDrawing(read(doc(shape(inner, words) + String.raw`\par`)));
  assert.deepEqual(drawing.anchor, {
    horizontal: { relativeFrom: 'margin', align: 'center' },
    vertical: { relativeFrom: 'page', offset: 300 * TWIP },
    wrap: 'square',
    wrapText: 'left',
    behindDoc: true,
    distL: 114300,
  });
  assert.deepEqual(drawing.transform, { rotation: 90 * 60000, flipH: true });
  assert.equal(drawing.description, 'A shape');
  const front = firstDrawing(read(doc(shape(String.raw`{\sp{\sn fBehindDocument}{\sv 0}}`, words) + String.raw`\par`)));
  assert.equal(front.anchor!.behindDoc, undefined);
});

test('the picture of a shape is the value of its property "pib"', () => {
  const inner = String.raw`{\sp{\sn shapeType}{\sv 75}}{\sp{\sn pib}{\sv {\pict\pngblip\picwgoal600\pichgoal300\piccropl60 ` + hexOf(PNG) + String.raw`}}}`;
  const model = read(doc(shape(inner) + String.raw`\par`));
  const drawing = firstDrawing(model);
  assert.equal(drawing.kind, 'picture');
  assert.equal(drawing.placement, 'anchor');
  assert.deepEqual(Array.from(mediaOf(model, drawing)), Array.from(PNG));
  assert.deepEqual(drawing.crop, { left: 10000 });
  assert.equal(drawingsOf(paragraphs(model)[0]).length, 1);
});

test('a group of shapes shows the pictures and text boxes of the shapes in it', () => {
  const picture = String.raw`{\shp{\*\shpinst{\sp{\sn pib}{\sv {\pict\pngblip\picwgoal600\pichgoal300 ` + hexOf(PNG) + String.raw`}}}}}`;
  const box = String.raw`{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt\pard\plain In group\par}}}`;
  const group = String.raw`{\shpgrp{\*\shpinst\shpleft0\shptop0\shpright5000\shpbottom4000\shpwr3` + picture + box + String.raw`}}`;
  const model = read(doc(group + String.raw`\par`));
  const drawings = drawingsOf(paragraphs(model)[0]);
  assert.equal(drawings.length, 1);
  assert.equal(drawings[0].kind, 'group');
  assert.equal(drawings[0].images!.length, 1);
  assert.deepEqual(paragraphsOf(drawings[0].textBoxes![0]).map(paragraphText), ['In group']);
});

test('text boxes in text boxes are read to a depth, and a shape in a text box belongs to that text', () => {
  const nest = (depth: number): string =>
    String.raw`{\shp{\*\shpinst\shpleft0\shptop0\shpright1000\shpbottom1000{\shptxt\pard\plain L` + depth + (depth < 8 ? nest(depth + 1) : '') + String.raw`\par}}}`;
  const model = read(doc(nest(1) + String.raw`\par`));
  let drawing: DrawingItem | undefined = firstDrawing(model);
  let depth = 0;
  // Each round goes one text box further in.
  while (drawing && drawing.textBoxes && depth < 20) {
    depth++;
    const inside: Paragraph = paragraphsOf(drawing.textBoxes[0])[0];
    assert.equal(paragraphText(inside), 'L' + depth);
    drawing = drawingsOf(inside)[0];
  }
  assert.equal(depth, 4);
});

test('a text box of Word 95 standing alone is read; its lines and curves are not drawn', () => {
  const box = String.raw`{\*\do\dobxpage\dobypara\dodhgt8192\dptxbx\dptxlrtb{\dptxbxtext\pard\plain Old box\par}\dpx1000\dpy2000\dpxsize3000\dpysize1500\dpfillfgcr255\dpfillfgcg255\dpfillfgcb255\dpfillbgcr255\dpfillpat1}`;
  const line = String.raw`{\*\do\dobxcolumn\dobypara\dodhgt8193\dpline\dpptx0\dppty0\dpptx100\dppty100\dpx0\dpy0\dpxsize100\dpysize100\dplinew15\dplinecor0}`;
  const model = read(doc('a' + box + line + String.raw`b\par`));
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'ab');
  const drawings = drawingsOf(p);
  assert.equal(drawings.length, 1);
  assert.deepEqual(paragraphsOf(drawings[0].textBoxes![0]).map(paragraphText), ['Old box']);
  assert.equal(drawings[0].width, 3000 * TWIP);
  assert.equal(drawings[0].anchor!.horizontal!.relativeFrom, 'page');
  assert.equal(drawings[0].anchor!.horizontal!.offset, 1000 * TWIP);
  assert.equal(model.unknown.size, 0);
});

test('pictures and text boxes lay out', () => {
  const inner = String.raw`{\sp{\sn shapeType}{\sv 202}}{\shptxt\pard\plain Boxed\par}`;
  const model = read(doc(String.raw`a{\pict\pngblip\picwgoal600\pichgoal300 ` + hexOf(PNG) + '}' + shape(inner) + String.raw`\par`));
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  const items = layout.pages[0].items;
  const image = items.find((item) => item.type === 'image');
  if (!image || image.type !== 'image') throw new Error('the page has no picture');
  assert.equal(Math.round(image.width), 40);
  assert.equal(Math.round(image.height), 20);
  assert.equal(image.part, firstDrawing(model).image!.target);
  assert.ok(items.some((item) => item.type === 'text' && item.text === 'Boxed'));
});

test('a picture inside a field result, a table cell and a note', () => {
  const picture = String.raw`{\pict\pngblip\picwgoal600\pichgoal300 ` + hexOf(PNG) + '}';
  const body = String.raw`{\field{\*\fldinst INCLUDEPICTURE "x.png"}{\fldrslt ` + picture + String.raw`}}\par\trowd\cellx3000\pard\intbl ` + picture + String.raw`\cell\row\pard x\chftn{\footnote\pard\plain ` + picture + String.raw`}\par`;
  const model = read(doc(body));
  const all = paragraphs(model).map((p) => itemsOf(p).filter((item) => item.type === 'drawing').length);
  assert.deepEqual(all, [1, 1, 0]);
  assert.equal(drawingsOf(paragraphsOf(model.footnotes!.items.get(1)!.blocks)[0]).length, 1);
  assert.equal(model.media!.size, 1);
});

// An embedded object (a chart of Microsoft Graph, a sheet of Excel) is "{\object ... {\*\objdata ...}{\result ...}}":
// what stands in its result is what is shown. The data of the object is never looked into.
test('an embedded chart shows the picture of its result: a bitmap as it is, a metafile as a frame without an image', () => {
  const object = (result: string) =>
    doc(String.raw`a{\object\objemb\objw3000\objh2000{\*\objclass MSGraph.Chart.8}{\*\objdata 0105000002000000080000004d53477261706800d0cf11e0a1b11ae1}{\result ` + result + String.raw`}}b\par`);
  const bitmap = read(object(String.raw`{\pict\pngblip\picwgoal3000\pichgoal2000 ` + hexOf(PNG) + '}'));
  assert.equal(paragraphText(paragraphs(bitmap)[0]), 'ab');
  const shown = firstDrawing(bitmap);
  assert.deepEqual([shown.kind, shown.placement, shown.width, shown.height], ['picture', 'inline', 3000 * TWIP, 2000 * TWIP]);
  assert.deepEqual(Array.from(mediaOf(bitmap, shown)), Array.from(PNG));
  assert.equal(bitmap.media!.size, 1, 'the data of the object is no picture of the document');

  // A metafile (what Word writes for nearly every object) keeps its place and its size; nothing is drawn in it.
  const metafile = read(object(String.raw`{\pict\wmetafile8\picw5292\pich3528\picwgoal3000\pichgoal2000 010009000003}`));
  const frame = firstDrawing(metafile);
  assert.deepEqual([frame.kind, frame.width, frame.height], ['picture', 3000 * TWIP, 2000 * TWIP]);

  // An object without a result shows nothing.
  const bare = read(doc(String.raw`a{\object\objemb{\*\objclass Excel.Chart.8}{\*\objdata 0105000002000000}}b\par`));
  assert.deepEqual([paragraphText(paragraphs(bare)[0]), drawingsOf(paragraphs(bare)[0]).length], ['ab', 0]);
});
