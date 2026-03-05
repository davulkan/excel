/// Streaming Excel writer for O(1) memory row encoding.
///
/// Requires `dart:io` — only available on server/CLI platforms.
///
/// ```dart
/// import 'package:excel/excel_streaming.dart';
///
/// final writer = ExcelStreamWriter();
/// final sheet = writer.addSheet('Data');
/// for (final row in bigData) {
///   sheet.appendRow([TextCellValue('hello'), IntCellValue(42)]);
/// }
/// final bytes = writer.close();
/// ```
library excel_streaming;

export 'excel.dart';
export 'src/stream/excel_stream_writer.dart';
export 'src/stream/stream_sheet.dart';
