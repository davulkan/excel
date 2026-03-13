import 'dart:convert';
import 'dart:io';
import 'dart:typed_data';

import '../../excel.dart';
import 'stream_shared_strings.dart';
import 'stream_styles.dart';

/// Pre-computed column letter cache. Avoids repeated _numericToLetters calls.
final List<String> _colLetterCache = List.generate(16384, (i) {
  // Replicate _numericToLetters(i + 1) inline
  var number = i + 1;
  var letters = '';
  while (number != 0) {
    var remainder = number % 26;
    if (remainder == 0) remainder = 26;
    letters = String.fromCharCode(64 + remainder) + letters;
    number = (number - 1) ~/ 26;
  }
  return letters;
});

/// A single sheet in the streaming writer.
///
/// Rows are appended one at a time via [appendRow]. Each row is immediately
/// serialized to XML and flushed to a temporary file in batches, so memory
/// usage stays constant regardless of how many rows are written.
class StreamSheet {
  final String name;
  final RandomAccessFile _file;
  final SharedStringCollector _strings;
  final StreamStyleManager _styles;

  /// Reusable write buffer — flushed when full or on finalize.
  Uint8List _writeBuf = Uint8List(65536); // 64 KB
  int _writePos = 0;

  /// Reusable StringBuffer for building one row's XML.
  final StringBuffer _buf = StringBuffer();

  int _rowIndex = 0;
  int _maxCol = 0;
  bool _finalized = false;

  StreamSheet.internal(this.name, this._file, this._strings, this._styles) {
    _writeRaw('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetData>');
  }

  /// Writes raw string to the buffered output.
  void _writeRaw(String s) {
    final bytes = utf8.encode(s);
    _writeBytes(bytes);
  }

  /// Appends bytes to the write buffer, flushing when needed.
  void _writeBytes(List<int> bytes) {
    var offset = 0;
    while (offset < bytes.length) {
      final remaining = _writeBuf.length - _writePos;
      final chunk = bytes.length - offset;
      if (chunk <= remaining) {
        _writeBuf.setRange(_writePos, _writePos + chunk, bytes, offset);
        _writePos += chunk;
        offset += chunk;
      } else {
        // Fill buffer and flush
        _writeBuf.setRange(_writePos, _writePos + remaining, bytes, offset);
        _writePos += remaining;
        offset += remaining;
        _flush();
      }
    }
  }

  void _flush() {
    if (_writePos > 0) {
      _file.writeFromSync(_writeBuf, 0, _writePos);
      _writePos = 0;
    }
  }

  /// Appends a single row of cell values.
  ///
  /// [values] — cell values for each column (null for empty cells).
  /// [styles] — optional per-cell styles. Must match [values] length.
  void appendRow(List<CellValue?> values, {List<CellStyle?>? styles}) {
    if (_finalized) {
      throw StateError('Cannot append to a finalized sheet');
    }
    if (styles != null && styles.length != values.length) {
      throw ArgumentError(
          'styles length (${styles.length}) must match values length (${values.length})');
    }
    if (values.length > _maxCol) _maxCol = values.length;

    _buf.clear();
    final rowNum = _rowIndex + 1;
    _buf.write('<row r="');
    _buf.write(rowNum);
    _buf.write('">');

    for (var col = 0; col < values.length; col++) {
      _writeCellXml(col, rowNum, values[col], styles?[col]);
    }

    _buf.write('</row>');

    // Encode and buffer the row
    _writeBytes(utf8.encode(_buf.toString()));
    _rowIndex++;
  }

  void _writeCellXml(
      int col, int rowNum, CellValue? value, CellStyle? cellStyle) {
    // Use cached column letters + row number for cell reference
    final colLetter = col < _colLetterCache.length
        ? _colLetterCache[col]
        : getCellId(col, 0).replaceAll('1', '');

    final int styleIndex;
    if (cellStyle != null) {
      styleIndex = _styles.getOrRegister(cellStyle);
    } else {
      styleIndex = _styles.defaultForValue(value);
    }

    if (value == null) {
      if (styleIndex > 0) {
        _buf.write('<c r="$colLetter$rowNum" s="$styleIndex"/>');
      }
      return;
    }

    final numFormat = cellStyle?.numberFormat ?? NumFormat.defaultFor(value);

    // Common cell opening: <c r="AB12"
    _buf.write('<c r="$colLetter$rowNum"');

    switch (value) {
      case TextCellValue():
        final ssIdx = _strings.getOrAdd(value.toString());
        _buf.write(' t="s"');
        if (styleIndex > 0) _buf.write(' s="$styleIndex"');
        _buf.write('><v>$ssIdx</v></c>');

      case FormulaCellValue():
        if (styleIndex > 0) _buf.write(' s="$styleIndex"');
        _buf.write('><f>${escapeXmlText(value.formula)}</f><v></v></c>');

      case IntCellValue():
        if (styleIndex > 0) _buf.write(' s="$styleIndex"');
        final v = switch (numFormat) {
          NumericNumFormat() => numFormat.writeInt(value),
          _ => value.value.toString(),
        };
        _buf.write('><v>$v</v></c>');

      case DoubleCellValue():
        if (styleIndex > 0) _buf.write(' s="$styleIndex"');
        final v = switch (numFormat) {
          NumericNumFormat() => numFormat.writeDouble(value),
          _ => value.value.toString(),
        };
        _buf.write('><v>$v</v></c>');

      case BoolCellValue():
        _buf.write(' t="b"');
        if (styleIndex > 0) _buf.write(' s="$styleIndex"');
        _buf.write('><v>${value.value ? '1' : '0'}</v></c>');

      case DateTimeCellValue():
        if (styleIndex > 0) _buf.write(' s="$styleIndex"');
        if (numFormat is! DateTimeNumFormat) {
          throw Exception('$numFormat does not work for ${value.runtimeType}');
        }
        final v = numFormat.writeDateTime(value);
        _buf.write('><v>$v</v></c>');

      case DateCellValue():
        if (styleIndex > 0) _buf.write(' s="$styleIndex"');
        if (numFormat is! DateTimeNumFormat) {
          throw Exception('$numFormat does not work for ${value.runtimeType}');
        }
        final v = numFormat.writeDate(value);
        _buf.write('><v>$v</v></c>');

      case TimeCellValue():
        if (styleIndex > 0) _buf.write(' s="$styleIndex"');
        if (numFormat is! TimeNumFormat) {
          throw Exception('$numFormat does not work for ${value.runtimeType}');
        }
        final v = numFormat.writeTime(value);
        _buf.write('><v>$v</v></c>');
    }
  }

  /// Writes closing XML tags. Called by ExcelStreamWriter.close().
  void finalize() {
    if (_finalized) return;
    _finalized = true;
    _writeRaw('</sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75"'
        ' bottom="0.75" header="0.3" footer="0.3"/>'
        '</worksheet>');
    _flush();
    _file.closeSync();
  }

  /// Number of rows written so far.
  int get rowCount => _rowIndex;

  /// Maximum number of columns in any row.
  int get maxColumns => _maxCol;
}
