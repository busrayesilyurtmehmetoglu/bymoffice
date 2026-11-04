// Runs the example's controller against the real bridge page in the platform's
// WebView (WKWebView on macOS and iOS, the system WebView on Android):
// flutter test integration_test -d macos   (or -d <an Android or iOS device>)
import 'dart:typed_data';

import 'package:bym_office_example/bym_office_view.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:integration_test/integration_test.dart';

void main() {
  final binding = IntegrationTestWidgetsFlutterBinding.ensureInitialized();
  binding.framePolicy = LiveTestWidgetsFlutterBindingFramePolicy.fullyLive;

  Future<BymOfficeController> show(WidgetTester tester) async {
    final controller = BymOfficeController();
    await tester.pumpWidget(MaterialApp(home: Scaffold(body: SizedBox(width: 900, height: 600, child: BymOfficeView(controller: controller)))));
    final version = await tester.runAsync(() => controller.ready.timeout(const Duration(seconds: 30)));
    debugPrint('bridge page ready: $version');
    return controller;
  }

  bool isZip(Uint8List bytes) => bytes.length > 4 && bytes[0] == 0x50 && bytes[1] == 0x4b;

  testWidgets('a blank workbook: a sheet is added, the file is saved and opens again with both sheets', (tester) async {
    final editor = await show(tester);
    final blank = (await tester.runAsync(() => editor.loadBlank('xlsx', locale: 'en')))!;
    expect(blank.isDocument, isFalse);
    expect(blank.sheets, hasLength(1));

    final at = (await tester.runAsync(() => editor.insertSheet(name: 'Second')))!;
    expect(at, 1);
    final saved = (await tester.runAsync(() => editor.saveFile()))!;
    expect(saved.format, isNull, reason: 'format comes only with a converted file');
    expect(isZip(saved.bytes), isTrue);

    final again = (await tester.runAsync(() => editor.load(saved.bytes, name: 'book.xlsx')))!;
    expect(again.sheets.map((sheet) => sheet.name), <String>[blank.sheets.first.name, 'Second']);
    editor.dispose();
  });

  testWidgets('a blank document opens as a document and saves as docx', (tester) async {
    final editor = await show(tester);
    final document = (await tester.runAsync(() => editor.loadBlank('docx', locale: 'tr', toolbar: true)))!;
    expect(document.isDocument, isTrue);
    expect(document.format, 'docx');
    final bytes = (await tester.runAsync(() => editor.save()))!;
    expect(isZip(bytes), isTrue);
    final reopened = (await tester.runAsync(() => editor.load(bytes, name: 'yazi.docx')))!;
    expect(reopened.isDocument, isTrue);
    editor.dispose();
  });

  testWidgets('a file that is not an Office file is refused with an error code, and the open one stays', (tester) async {
    final editor = await show(tester);
    await tester.runAsync(() => editor.loadBlank('xlsx'));
    final Object? error = await tester.runAsync<Object?>(() async {
      try {
        await editor.load(Uint8List.fromList(List<int>.filled(64, 7)), name: 'x.xlsx');
        return null;
      } catch (e) {
        return e;
      }
    });
    expect(error, isA<BymOfficeException>());
    expect(<String>['unsupported', 'corrupt', 'load-failed'], contains((error as BymOfficeException).code));
    final bytes = (await tester.runAsync(() => editor.save()))!;
    expect(isZip(bytes), isTrue);
    editor.dispose();
  });

  testWidgets('an RTF file opens converted to docx and saves as docx', (tester) async {
    final editor = await show(tester);
    final rtf = Uint8List.fromList(r'{\rtf1\ansi{\fonttbl{\f0 Times New Roman;}}\f0\fs24 Merhaba d\u252?nya\par}'.codeUnits);
    final opened = (await tester.runAsync(() => editor.load(rtf, name: 'eski.rtf', convert: true)))!;
    expect(opened.isDocument, isTrue);
    expect(opened.readOnly, isFalse);
    expect(opened.format, 'docx');
    expect(opened.convertedFrom, 'rtf');
    final saved = (await tester.runAsync(() => editor.saveFile()))!;
    expect(saved.format, 'docx');
    expect(saved.convertedFrom, 'rtf');
    expect(isZip(saved.bytes), isTrue);
    editor.dispose();
  });
}
