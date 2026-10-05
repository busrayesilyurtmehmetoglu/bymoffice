// A reusable widget that shows the BYM Office editor in a WebView and talks to
// it through the bridge contract (docs/webview-koprusu.md).
//
// NOTE: written against webview_flutter 4.x (WebViewController,
// addJavaScriptChannel, loadFlutterAsset, runJavaScript) but never compiled or
// run. Lines marked UNVERIFIED are the ones most likely to need a correction.

import 'dart:async';
import 'dart:convert';

import 'package:flutter/foundation.dart';
import 'package:flutter/widgets.dart';
import 'package:webview_flutter/webview_flutter.dart';

/// One sheet of the opened workbook, as reported by `loaded`.
class BymOfficeSheet {
  const BymOfficeSheet({required this.name, required this.state, required this.kind});

  final String name;

  /// "visible", "hidden" or "veryHidden".
  final String state;

  /// "worksheet" or "chartsheet".
  final String kind;
}

/// The answer to a `load`.
class BymOfficeDocument {
  const BymOfficeDocument({
    required this.format,
    required this.sheets,
    required this.activeSheet,
    this.isDocument = false,
    this.readOnly = false,
    this.convertedFrom,
    this.warnings = const <String>[],
  });

  /// The format of the file that is shown and saved: "docx", "xlsx", "doc",
  /// "xls" or "rtf". After a conversion this is the new format.
  final String format;

  /// Empty for a document.
  final List<BymOfficeSheet> sheets;
  final int activeSheet;

  /// True for a Word document (docx, doc, RTF), false for a workbook.
  final bool isDocument;

  /// True when the file is in a format that cannot be written (doc, RTF, xls
  /// that were not converted): it stays read-only, whatever setReadOnly asks.
  final bool readOnly;

  /// The format the file had before `convert` turned it into [format]; null
  /// when nothing was converted.
  final String? convertedFrom;

  /// What a conversion could not carry over, in English; for developers.
  final List<String> warnings;
}

/// The answer to a `save`.
class BymOfficeSavedFile {
  const BymOfficeSavedFile({required this.bytes, this.format, this.convertedFrom});

  final Uint8List bytes;

  /// Set, with [convertedFrom], only for a file that was converted when it was
  /// loaded: the bytes are a file of this format, not of the one that was given.
  final String? format;
  final String? convertedFrom;
}

/// A `changed` message: the workbook or the document was edited (or an edit
/// was undone or redone).
class BymOfficeChange {
  const BymOfficeChange({required this.sheets, required this.canUndo, required this.canRedo});

  /// The sheets that changed; always empty for a document.
  final List<int> sheets;
  final bool canUndo;
  final bool canRedo;
}

/// A `selection` message. [range] and [active] are cell names such as "A1:B2" and "A1".
class BymOfficeSelection {
  const BymOfficeSelection({required this.sheet, required this.range, required this.active});

  final int sheet;
  final String range;
  final String active;
}

/// An `error` message. [code] is one of: unsupported, corrupt, load-failed,
/// save-failed, bad-message, not-loaded, print-unavailable, print-failed,
/// read-only, no-sheet, sheet-refused (and "timeout" and "closed", which only
/// this wrapper produces). For sheet-refused the [message] is the reason in
/// the language of the editor, fit to be shown to the user.
class BymOfficeException implements Exception {
  const BymOfficeException(this.code, this.message);

  final String code;
  final String message;

  @override
  String toString() => 'BymOfficeException($code): $message';
}

/// A `sheet` message: the user moved to another sheet tab.
class BymOfficeSheetChange {
  const BymOfficeSheetChange({required this.index, required this.name});

  final int index;
  final String name;
}

/// A `sheets` message: the list of sheets changed (a sheet was inserted,
/// deleted or renamed, by a command, by the user or by undo). [sheets] is the
/// whole list as it is now, in the form [BymOfficeDocument.sheets] has it.
///
/// NOTE: added with the sheet commands below; like the rest of this file it
/// was never compiled.
class BymOfficeSheetList {
  const BymOfficeSheetList({required this.sheets, required this.activeSheet});

  final List<BymOfficeSheet> sheets;

  /// Index of the sheet that is shown.
  final int activeSheet;
}

/// Drives one editor. Create it once, hand it to a [BymOfficeView] and dispose
/// it with the widget that owns it.
class BymOfficeController {
  BymOfficeController({
    this.assetKey = 'assets/bymoffice.webview.html',
    this.requestTimeout = const Duration(seconds: 60),
  }) {
    webView = WebViewController()
      ..setJavaScriptMode(JavaScriptMode.unrestricted)
      // The name is part of the contract: the page looks for window.BymOfficeHost.postMessage.
      ..addJavaScriptChannel(_channel, onMessageReceived: (JavaScriptMessage message) => _onMessage(message.message))
      // Should `ready` have been sent before the channel could take it, `hello` asks for it again.
      ..setNavigationDelegate(NavigationDelegate(onPageFinished: (String url) => _post(<String, dynamic>{'type': 'hello'})))
      ..loadFlutterAsset(assetKey);
  }

  static const String _channel = 'BymOfficeHost';

  /// Asset key of the bridge page as declared in pubspec.yaml.
  final String assetKey;

  /// How long a request waits for its answer, the page becoming ready included.
  final Duration requestTimeout;

  late final WebViewController webView;

  final Completer<String> _ready = Completer<String>();
  final Map<String, Completer<Map<String, dynamic>>> _pending = <String, Completer<Map<String, dynamic>>>{};
  final StreamController<BymOfficeChange> _changes = StreamController<BymOfficeChange>.broadcast();
  final StreamController<BymOfficeSelection> _selections = StreamController<BymOfficeSelection>.broadcast();
  final StreamController<BymOfficeSheetChange> _sheets = StreamController<BymOfficeSheetChange>.broadcast();
  final StreamController<BymOfficeSheetList> _sheetLists = StreamController<BymOfficeSheetList>.broadcast();
  final StreamController<BymOfficeException> _errors = StreamController<BymOfficeException>.broadcast();
  final StreamController<String> _links = StreamController<String>.broadcast();
  int _nextId = 0;
  bool _disposed = false;

  /// Completes with the version of the bridge page once it can take messages.
  Future<String> get ready => _ready.future;

  Stream<BymOfficeChange> get onChanged => _changes.stream;
  Stream<BymOfficeSelection> get onSelection => _selections.stream;
  Stream<BymOfficeSheetChange> get onSheetChanged => _sheets.stream;

  /// The list of sheets changed; never sent for a document.
  Stream<BymOfficeSheetList> get onSheetsChanged => _sheetLists.stream;

  /// A link in a document was activated; the page never opens it itself.
  Stream<String> get onLink => _links.stream;

  /// Every `error` message, including those that also fail a request.
  Stream<BymOfficeException> get onError => _errors.stream;

  /// Opens a file. If it cannot be opened, the file that was open stays open.
  ///
  /// [convert] opens doc, RTF and xls files for editing by converting them to
  /// docx and xlsx: the answer then names the new format and the old one, and
  /// what [save] returns is a file of the new format. [toolbar] shows a
  /// document with the formatting toolbar above its pages.
  Future<BymOfficeDocument> load(
    Uint8List bytes, {
    String? name,
    String? locale,
    bool? readOnly,
    int? sheet,
    bool? convert,
    bool? toolbar,
  }) async {
    _checkAlive();
    // Base64 of a large file takes a while; keep it off the UI thread.
    final String data = await compute(base64Encode, bytes);
    return _load(<String, dynamic>{'data': data}, name: name, locale: locale, readOnly: readOnly, sheet: sheet, convert: convert, toolbar: toolbar);
  }

  /// Starts with a new, empty file instead of opening one. [kind] is "docx"
  /// or "xlsx"; [locale] also chooses the paper size and the name of the sheet.
  Future<BymOfficeDocument> loadBlank(String kind, {String? name, String? locale, bool? readOnly, bool? toolbar}) {
    _checkAlive();
    return _load(<String, dynamic>{'blank': kind}, name: name, locale: locale, readOnly: readOnly, toolbar: toolbar);
  }

  Future<BymOfficeDocument> _load(
    Map<String, dynamic> source, {
    String? name,
    String? locale,
    bool? readOnly,
    int? sheet,
    bool? convert,
    bool? toolbar,
  }) async {
    final Map<String, dynamic> answer = await _request(<String, dynamic>{
      'type': 'load',
      ...source,
      if (name != null) 'name': name,
      if (locale != null) 'locale': locale,
      if (readOnly != null) 'readOnly': readOnly,
      if (sheet != null) 'sheet': sheet,
      if (convert != null) 'convert': convert,
      if (toolbar != null) 'toolbar': toolbar,
    });
    return BymOfficeDocument(
      format: answer['format'] as String,
      activeSheet: (answer['activeSheet'] as num).toInt(),
      sheets: (answer['sheets'] as List<dynamic>)
          .map((dynamic entry) => BymOfficeSheet(
                name: entry['name'] as String,
                state: entry['state'] as String,
                kind: entry['kind'] as String,
              ))
          .toList(),
      isDocument: answer['kind'] == 'document',
      readOnly: answer['readOnly'] == true,
      convertedFrom: answer['convertedFrom'] as String?,
      // UNVERIFIED: the cast of a decoded JSON list of strings.
      warnings: (answer['warnings'] as List<dynamic>? ?? const <dynamic>[]).map((dynamic entry) => entry.toString()).toList(),
    );
  }

  /// The file with every change made in the editor.
  Future<Uint8List> save() async => (await saveFile()).bytes;

  /// As [save], with the format of the bytes when the file was converted on loading.
  Future<BymOfficeSavedFile> saveFile() async {
    final Map<String, dynamic> answer = await _request(<String, dynamic>{'type': 'save'});
    return BymOfficeSavedFile(
      bytes: await compute(base64Decode, answer['data'] as String),
      format: answer['format'] as String?,
      convertedFrom: answer['convertedFrom'] as String?,
    );
  }

  // These complete when the page confirms the command with `done`.

  /// Completes with what holds afterwards: true when the file is read-only.
  /// A doc, RTF or xls file that was not converted stays read-only whatever
  /// is asked, so `setReadOnly(false)` completes with true for it.
  Future<bool> setReadOnly(bool readOnly) async {
    final Map<String, dynamic> answer = await _request(<String, dynamic>{'type': 'setReadOnly', 'readOnly': readOnly});
    return readOnly || answer['readOnly'] == true;
  }

  Future<void> setLocale(String locale) => _request(<String, dynamic>{'type': 'setLocale', 'locale': locale});
  Future<void> undo() => _request(<String, dynamic>{'type': 'undo'});
  Future<void> redo() => _request(<String, dynamic>{'type': 'redo'});
  Future<void> close() => _request(<String, dynamic>{'type': 'close'});

  // Sheets of a workbook. Each is one undo step; the new list arrives on
  // [onSheetsChanged]. They fail with a BymOfficeException whose code is
  // read-only, no-sheet (no sheet has the index; a document is open),
  // sheet-refused (the editor's reason is the message) or bad-message.

  /// Adds an empty sheet at [at] (after the last one when null) and shows it.
  /// Without a [name] it gets the next default name of the editor's language.
  /// Completes with the index of the new sheet.
  Future<int> insertSheet({int? at, String? name}) async {
    final Map<String, dynamic> answer = await _request(<String, dynamic>{
      'type': 'insertSheet',
      if (at != null) 'at': at,
      if (name != null) 'name': name,
    });
    return (answer['index'] as num).toInt();
  }

  Future<void> renameSheet(int index, String name) => _request(<String, dynamic>{'type': 'renameSheet', 'index': index, 'name': name});
  Future<void> deleteSheet(int index) => _request(<String, dynamic>{'type': 'deleteSheet', 'index': index});

  /// Closes the file in the page, detaches from the WebView and fails whatever
  /// is still waiting. Requests made afterwards fail at once.
  void dispose() {
    if (_disposed) return;
    _disposed = true;
    // Best effort: the page may already be gone.
    unawaited(_post(<String, dynamic>{'type': 'close'}).catchError((Object _) {}));
    unawaited(webView.removeJavaScriptChannel(_channel).catchError((Object _) {}));
    const BymOfficeException closed = BymOfficeException('closed', 'The editor was disposed');
    for (final Completer<Map<String, dynamic>> completer in _pending.values) {
      completer.completeError(closed);
    }
    _pending.clear();
    _changes.close();
    _selections.close();
    _sheets.close();
    _sheetLists.close();
    _errors.close();
    _links.close();
  }

  void _checkAlive() {
    if (_disposed) throw const BymOfficeException('closed', 'The editor was disposed');
  }

  /// Sends a message and waits for the answer that carries the same id
  /// (`loaded`, `saved`, `done` or `error`). The time limit covers the wait
  /// for the page to be ready as well.
  Future<Map<String, dynamic>> _request(Map<String, dynamic> message) {
    _checkAlive();
    final String id = 'f${_nextId++}';
    final Completer<Map<String, dynamic>> completer = Completer<Map<String, dynamic>>();
    // dispose() may fail this completer when nobody listens to it any more.
    completer.future.ignore();
    _pending[id] = completer;

    Future<Map<String, dynamic>> run() async {
      await ready;
      _checkAlive();
      await _post(<String, dynamic>{...message, 'id': id});
      return completer.future;
    }

    return run().timeout(requestTimeout, onTimeout: () {
      throw const BymOfficeException('timeout', 'The editor did not answer in time');
    }).whenComplete(() => _pending.remove(id));
  }

  /// Hands one message to the page, without waiting for `ready`.
  Future<void> _post(Map<String, dynamic> message) {
    // The message travels as JSON text. Encoding that text once more gives a
    // JavaScript string literal, except for the two line separators: JSON
    // allows them raw, older JavaScript engines do not, so they are escaped.
    final String literal = jsonEncode(jsonEncode(message))
        .replaceAll(String.fromCharCode(0x2028), r'\u2028')
        .replaceAll(String.fromCharCode(0x2029), r'\u2029');
    // UNVERIFIED: very large scripts (tens of MB of base64) are slow to pass to
    // the WebView and may hit platform limits; see README.md.
    return webView.runJavaScript('window.BymOfficeBridge.receive($literal);');
  }

  void _onMessage(String text) {
    if (_disposed) return;
    final Object? decoded;
    try {
      decoded = jsonDecode(text);
    } on FormatException {
      return;
    }
    if (decoded is! Map<String, dynamic>) return;
    final Map<String, dynamic> message = decoded;
    final Object? id = message['id'];
    switch (message['type']) {
      case 'ready':
        // Sent again after every `hello`; only the first one counts.
        if (!_ready.isCompleted) _ready.complete(message['version'] as String? ?? '');
      case 'loaded':
      case 'saved':
      case 'done':
        if (id is String) _pending[id]?.complete(message);
      case 'changed':
        _changes.add(BymOfficeChange(
          sheets: (message['sheets'] as List<dynamic>).map((dynamic index) => (index as num).toInt()).toList(),
          canUndo: message['canUndo'] == true,
          canRedo: message['canRedo'] == true,
        ));
      case 'selection':
        _selections.add(BymOfficeSelection(
          sheet: (message['sheet'] as num).toInt(),
          range: message['range'] as String,
          active: message['active'] as String,
        ));
      case 'sheet':
        _sheets.add(BymOfficeSheetChange(index: (message['index'] as num).toInt(), name: message['name'] as String? ?? ''));
      case 'sheets':
        _sheetLists.add(BymOfficeSheetList(
          activeSheet: (message['activeSheet'] as num).toInt(),
          sheets: (message['sheets'] as List<dynamic>)
              .map((dynamic entry) => BymOfficeSheet(
                    name: entry['name'] as String,
                    state: entry['state'] as String,
                    kind: entry['kind'] as String,
                  ))
              .toList(),
        ));
      case 'link':
        final Object? url = message['url'];
        if (url is String) _links.add(url);
      case 'error':
        final BymOfficeException error = BymOfficeException(message['code'] as String? ?? 'unknown', message['message'] as String? ?? '');
        _errors.add(error);
        if (id is String) _pending[id]?.completeError(error);
    }
  }
}

/// Shows the editor of a [BymOfficeController]. Give it a bounded size.
class BymOfficeView extends StatelessWidget {
  const BymOfficeView({super.key, required this.controller});

  final BymOfficeController controller;

  @override
  Widget build(BuildContext context) => WebViewWidget(controller: controller.webView);
}
