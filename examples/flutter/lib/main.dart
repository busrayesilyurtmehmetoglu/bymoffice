// Example application: pick a Word or Excel file, edit it, save it.
//
// NOTE: never compiled or run; see README.md. The file_picker calls are
// UNVERIFIED against the package's current API.

import 'dart:async';
import 'dart:io';

import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';

import 'bym_office_view.dart';

void main() => runApp(const ExampleApp());

class ExampleApp extends StatelessWidget {
  const ExampleApp({super.key});

  @override
  Widget build(BuildContext context) => const MaterialApp(title: 'BYM Office', home: EditorPage());
}

class EditorPage extends StatefulWidget {
  const EditorPage({super.key});

  @override
  State<EditorPage> createState() => _EditorPageState();
}

class _EditorPageState extends State<EditorPage> {
  final BymOfficeController _editor = BymOfficeController();
  late final StreamSubscription<BymOfficeChange> _changes;
  late final StreamSubscription<BymOfficeException> _errors;
  String? _name;
  bool _dirty = false;
  bool _canUndo = false;
  bool _canRedo = false;

  @override
  void initState() {
    super.initState();
    _changes = _editor.onChanged.listen((BymOfficeChange change) {
      setState(() {
        _dirty = true;
        _canUndo = change.canUndo;
        _canRedo = change.canRedo;
      });
    });
    _errors = _editor.onError.listen((BymOfficeException error) => _show('${error.code}: ${error.message}'));
  }

  @override
  void dispose() {
    _changes.cancel();
    _errors.cancel();
    _editor.dispose();
    super.dispose();
  }

  void _show(String text) {
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(text)));
  }

  Future<void> _open() async {
    final FilePickerResult? picked = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: <String>['xlsx', 'docx', 'xls', 'doc', 'rtf'],
      withData: true,
    );
    final Uint8List? bytes = picked?.files.single.bytes;
    if (picked == null || bytes == null) return;
    try {
      // Old formats are converted so that they can be edited; documents get the formatting toolbar.
      final BymOfficeDocument opened = await _editor.load(bytes, name: picked.files.single.name, locale: 'tr', convert: true, toolbar: true);
      String name = picked.files.single.name;
      if (opened.convertedFrom != null) {
        // "report.doc" is now a docx: only the extension of the old format is replaced.
        name = '${name.replaceFirst(RegExp(r'\.(doc|rtf|xls)$', caseSensitive: false), '')}.${opened.format}';
        if (opened.warnings.isNotEmpty) _show('${opened.format} biçimine çevrildi; bazı içerik taşınamadı');
      }
      setState(() {
        _name = name;
        _dirty = false;
        _canUndo = false;
        _canRedo = false;
      });
    } on BymOfficeException {
      // Already shown through onError.
    }
  }

  Future<void> _save() async {
    try {
      final Uint8List bytes = await _editor.save();
      // UNVERIFIED: on mobile file_picker writes `bytes` itself; on desktop it
      // returns the chosen path and the application writes the file.
      final String? path = await FilePicker.platform.saveFile(fileName: _name ?? 'workbook.xlsx', bytes: bytes);
      if (path == null) return;
      if (Platform.isMacOS || Platform.isWindows || Platform.isLinux) await File(path).writeAsBytes(bytes);
      setState(() => _dirty = false);
      _show('Kaydedildi');
    } on BymOfficeException {
      // Already shown through onError.
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('${_name ?? 'BYM Office'}${_dirty ? ' *' : ''}'),
        actions: <Widget>[
          // A failure is shown through onError; the future itself needs no handler.
          IconButton(icon: const Icon(Icons.undo), onPressed: _canUndo ? () => unawaited(_editor.undo().catchError((Object _) {})) : null),
          IconButton(icon: const Icon(Icons.redo), onPressed: _canRedo ? () => unawaited(_editor.redo().catchError((Object _) {})) : null),
          IconButton(icon: const Icon(Icons.folder_open), onPressed: _open),
          IconButton(icon: const Icon(Icons.save), onPressed: _name == null ? null : _save),
        ],
      ),
      // The page handles the safe area itself; the keyboard resizes the body.
      body: BymOfficeView(controller: _editor),
    );
  }
}
