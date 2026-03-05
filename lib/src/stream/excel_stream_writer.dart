import 'dart:io';

import 'package:archive/archive_io.dart';
import 'package:excel/excel.dart';

import 'stream_shared_strings.dart';
import 'stream_sheet.dart';
import 'stream_styles.dart';
import 'stream_xml_templates.dart';

/// A streaming Excel writer that uses O(1) memory per row.
///
/// Unlike [Excel], which holds all cell data in memory, this writer
/// serializes each row to disk immediately via [StreamSheet.appendRow].
/// The final .xlsx file is assembled from temporary files when [close]
/// or [closeToFile] is called.
///
/// ```dart
/// import 'package:excel/excel_streaming.dart';
///
/// final writer = ExcelStreamWriter();
/// final sheet = writer.addSheet('Data');
/// for (final row in bigData) {
///   sheet.appendRow([TextCellValue('hello'), IntCellValue(42)]);
/// }
/// writer.closeToFile('output.xlsx');
/// ```
class ExcelStreamWriter {
  final List<StreamSheet> _sheets = [];
  final SharedStringCollector _sharedStrings = SharedStringCollector();
  final StreamStyleManager _styles = StreamStyleManager();
  Directory? _tempDir;
  bool _closed = false;

  /// Registers a [CellStyle] and returns its index.
  int registerStyle(CellStyle style) => _styles.getOrRegister(style);

  /// Adds a new sheet with the given [name] and returns it.
  StreamSheet addSheet(String name) {
    if (_closed) throw StateError('Writer is already closed');
    _tempDir ??= Directory.systemTemp.createTempSync('excel_stream_');

    final idx = _sheets.length + 1;
    final file = File('${_tempDir!.path}/sheet$idx.xml');
    final raf = file.openSync(mode: FileMode.write);

    final sheet = StreamSheet.internal(name, raf, _sharedStrings, _styles);
    _sheets.add(sheet);
    return sheet;
  }

  /// Assembles the .xlsx ZIP and writes it directly to [path].
  ///
  /// Uses streaming ZIP encoding — sheet files are read from disk via
  /// InputFileStream and compressed on-the-fly, avoiding loading the
  /// entire uncompressed XML into memory.
  void closeToFile(String path) {
    if (_closed) throw StateError('Writer is already closed');
    _closed = true;

    for (final sheet in _sheets) {
      sheet.finalize();
    }

    final sheetNames = _sheets.map((s) => s.name).toList();

    // Write small scaffolding files to temp dir
    _writeTempFile('Content_Types.xml',
        StreamXmlTemplates.contentTypes(sheetNames));
    _writeTempFile('rels.xml', StreamXmlTemplates.topLevelRels());
    _writeTempFile('workbook_rels.xml',
        StreamXmlTemplates.workbookRels(sheetNames));
    _writeTempFile('workbook.xml', StreamXmlTemplates.workbook(sheetNames));
    _writeTempFile('styles.xml', _styles.buildStylesXmlBytes());
    _writeTempFile('sharedStrings.xml', _sharedStrings.buildXmlBytes());

    // Build ZIP using ZipFileEncoder — streams files from disk
    final encoder = ZipFileEncoder();
    encoder.create(path);

    // Add small files
    _addStreamingFile(encoder, 'Content_Types.xml', '[Content_Types].xml');
    _addStreamingFile(encoder, 'rels.xml', '_rels/.rels');
    _addStreamingFile(
        encoder, 'workbook_rels.xml', 'xl/_rels/workbook.xml.rels');
    _addStreamingFile(encoder, 'workbook.xml', 'xl/workbook.xml');
    _addStreamingFile(encoder, 'styles.xml', 'xl/styles.xml');
    _addStreamingFile(encoder, 'sharedStrings.xml', 'xl/sharedStrings.xml');

    // Add sheet files — streamed from disk via InputFileStream
    for (var i = 0; i < _sheets.length; i++) {
      final tempPath = '${_tempDir!.path}/sheet${i + 1}.xml';
      final fileStream = InputFileStream(tempPath);
      final archiveFile = ArchiveFile.stream(
        'xl/worksheets/sheet${i + 1}.xml',
        File(tempPath).lengthSync(),
        fileStream,
      );
      encoder.addArchiveFile(archiveFile);
      fileStream.closeSync();
    }

    encoder.closeSync();
    _cleanup();
  }

  /// Assembles the .xlsx ZIP and returns it as bytes.
  ///
  /// Note: for large files, prefer [closeToFile] which uses streaming ZIP
  /// and avoids holding the entire ZIP in memory.
  List<int> close() {
    if (_closed) throw StateError('Writer is already closed');

    final tmpZipFile = File(
        '${Directory.systemTemp.path}/excel_stream_out_${DateTime.now().microsecondsSinceEpoch}.xlsx');
    try {
      closeToFile(tmpZipFile.path);
      return tmpZipFile.readAsBytesSync();
    } finally {
      if (tmpZipFile.existsSync()) tmpZipFile.deleteSync();
    }
  }

  /// Cleans up temp files if [close]/[closeToFile] was not called.
  void dispose() {
    if (!_closed) {
      _closed = true;
      for (final sheet in _sheets) {
        try {
          sheet.finalize();
        } catch (_) {}
      }
    }
    _cleanup();
  }

  void _writeTempFile(String tempName, List<int> data) {
    File('${_tempDir!.path}/$tempName').writeAsBytesSync(data);
  }

  void _addStreamingFile(
      ZipFileEncoder encoder, String tempName, String archiveName) {
    final tempPath = '${_tempDir!.path}/$tempName';
    final fileStream = InputFileStream(tempPath);
    final archiveFile = ArchiveFile.stream(
      archiveName,
      File(tempPath).lengthSync(),
      fileStream,
    );
    encoder.addArchiveFile(archiveFile);
    fileStream.closeSync();
  }

  void _cleanup() {
    if (_tempDir != null && _tempDir!.existsSync()) {
      try {
        _tempDir!.deleteSync(recursive: true);
      } catch (_) {}
      _tempDir = null;
    }
  }
}
