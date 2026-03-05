import 'dart:io';
import 'package:excel/excel_streaming.dart';

/// Measures RSS memory of the current process via /proc/self/status on Linux.
int getRssKb() {
  final lines = File('/proc/self/status').readAsLinesSync();
  for (final line in lines) {
    if (line.startsWith('VmRSS:')) {
      final parts = line.split(RegExp(r'\s+'));
      return int.parse(parts[1]);
    }
  }
  return -1;
}

String formatMb(int kb) => '${(kb / 1024).toStringAsFixed(1)} MB';

void printMemory(String label) {
  final rss = getRssKb();
  print('[$label] RSS: ${formatMb(rss)}');
}

void main(List<String> args) {
  final rowCount = args.isNotEmpty ? int.parse(args[0]) : 100000;
  final colCount = args.length > 1 ? int.parse(args[1]) : 20;

  print('');
  print('============================================');
  print('  Streaming Writer vs Classic Benchmark');
  print('============================================');
  print('Rows: $rowCount, Columns: $colCount');
  print('Total cells: ${rowCount * colCount}');
  print('');

  // ── Streaming writer ──────────────────────────
  print('--- ExcelStreamWriter ---');
  printMemory('STREAM START');

  final swStream = Stopwatch()..start();
  final writer = ExcelStreamWriter();
  final streamSheet = writer.addSheet('Data');

  // Header
  streamSheet.appendRow(
      List.generate(colCount, (i) => TextCellValue('Column_$i')));

  // Data rows
  for (var r = 0; r < rowCount; r++) {
    final row = List<CellValue?>.generate(colCount, (c) {
      if (c == 0) return IntCellValue(r);
      if (c == 1) return TextCellValue('Row_$r text data here for testing');
      if (c == 2) return DoubleCellValue(r * 1.5);
      if (c == 3) return BoolCellValue(r % 2 == 0);
      return IntCellValue(r * colCount + c);
    });
    streamSheet.appendRow(row);

    if ((r + 1) % 25000 == 0) {
      printMemory('STREAM AFTER ${r + 1} ROWS');
    }
  }

  final streamFillTime = swStream.elapsedMilliseconds;
  printMemory('STREAM ALL ROWS FILLED');
  print('Stream fill time: ${streamFillTime}ms');

  swStream.reset();
  swStream.start();
  writer.closeToFile('stream_benchmark_output.xlsx');
  final streamEncodeTime = swStream.elapsedMilliseconds;

  final streamFile = File('stream_benchmark_output.xlsx');
  final streamSize = streamFile.lengthSync();
  printMemory('STREAM AFTER CLOSE');
  print('Stream encode time: ${streamEncodeTime}ms');
  print(
      'Stream file size: ${(streamSize / 1024 / 1024).toStringAsFixed(1)} MB');
  print('');

  printMemory('BEFORE CLASSIC');

  // ── Classic Excel writer ──────────────────────
  print('--- Classic Excel ---');
  printMemory('CLASSIC START');

  final swClassic = Stopwatch()..start();
  final excel = Excel.createExcel();
  final classicSheet = excel['Data'];

  // Header
  classicSheet
      .appendRow(List.generate(colCount, (i) => TextCellValue('Column_$i')));

  // Data rows
  for (var r = 0; r < rowCount; r++) {
    final row = List<CellValue>.generate(colCount, (c) {
      if (c == 0) return IntCellValue(r);
      if (c == 1) return TextCellValue('Row_$r text data here for testing');
      if (c == 2) return DoubleCellValue(r * 1.5);
      if (c == 3) return BoolCellValue(r % 2 == 0);
      return IntCellValue(r * colCount + c);
    });
    classicSheet.appendRow(row);

    if ((r + 1) % 25000 == 0) {
      printMemory('CLASSIC AFTER ${r + 1} ROWS');
    }
  }

  final classicFillTime = swClassic.elapsedMilliseconds;
  printMemory('CLASSIC ALL ROWS FILLED');
  print('Classic fill time: ${classicFillTime}ms');

  excel.delete('Sheet1');

  swClassic.reset();
  swClassic.start();
  final bytes = excel.encode();
  final classicEncodeTime = swClassic.elapsedMilliseconds;

  printMemory('CLASSIC AFTER ENCODE');
  print('Classic encode time: ${classicEncodeTime}ms');

  if (bytes != null) {
    File('classic_benchmark_output.xlsx').writeAsBytesSync(bytes);
    print(
        'Classic file size: ${(bytes.length / 1024 / 1024).toStringAsFixed(1)} MB');
  }

  printMemory('CLASSIC FINAL');

  // ── Summary ──────────────────────────────────
  print('');
  print('============================================');
  print('  Summary: $rowCount rows x $colCount cols');
  print('============================================');
  print('                  Stream     Classic');
  print(
      'Fill time:     ${streamFillTime.toString().padLeft(7)}ms  ${classicFillTime.toString().padLeft(7)}ms');
  print(
      'Encode time:   ${streamEncodeTime.toString().padLeft(7)}ms  ${classicEncodeTime.toString().padLeft(7)}ms');
  print(
      'Total time:    ${(streamFillTime + streamEncodeTime).toString().padLeft(7)}ms  ${(classicFillTime + classicEncodeTime).toString().padLeft(7)}ms');
  print(
      'File size:     ${(streamSize / 1024 / 1024).toStringAsFixed(1).padLeft(6)} MB  ${bytes != null ? (bytes.length / 1024 / 1024).toStringAsFixed(1).padLeft(6) : '  N/A'} MB');

  // Cleanup
  streamFile.deleteSync();
  File('classic_benchmark_output.xlsx').deleteSync();
}
