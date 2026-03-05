import 'dart:io';
import 'package:excel/excel.dart';

/// Measures RSS memory of the current process via /proc/self/status on Linux.
int getRssKb() {
  final lines = File('/proc/self/status').readAsLinesSync();
  for (final line in lines) {
    if (line.startsWith('VmRSS:')) {
      final parts = line.split(RegExp(r'\s+'));
      return int.parse(parts[1]); // in kB
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
  // Configuration
  final rowCount = args.isNotEmpty ? int.parse(args[0]) : 100000;
  final colCount = args.length > 1 ? int.parse(args[1]) : 20;

  print('=== Excel Memory Benchmark ===');
  print('Rows: $rowCount, Columns: $colCount');
  print('Total cells: ${rowCount * colCount}');
  print('');

  printMemory('START');

  // Step 1: Create Excel object
  final sw = Stopwatch()..start();
  final excel = Excel.createExcel();
  final sheet = excel['Data'];
  printMemory('AFTER CREATE');

  // Step 2: Write header row
  final header = List.generate(colCount, (i) => TextCellValue('Column_$i'));
  sheet.appendRow(header);

  // Step 3: Fill data rows
  print('');
  print('Filling $rowCount rows...');
  for (var r = 0; r < rowCount; r++) {
    final row = List<CellValue>.generate(colCount, (c) {
      if (c == 0) return IntCellValue(r);
      if (c == 1) return TextCellValue('Row_$r text data here for testing');
      if (c == 2) return DoubleCellValue(r * 1.5);
      if (c == 3) return BoolCellValue(r % 2 == 0);
      return IntCellValue(r * colCount + c);
    });
    sheet.appendRow(row);

    if ((r + 1) % 25000 == 0) {
      printMemory('AFTER ${r + 1} ROWS');
    }
  }

  final fillTime = sw.elapsedMilliseconds;
  printMemory('ALL ROWS FILLED');
  print('Fill time: ${fillTime}ms');
  print('');

  // Step 4: Delete default Sheet1
  excel.delete('Sheet1');

  // Step 5: Encode to bytes
  print('Encoding to bytes...');
  sw.reset();
  sw.start();
  final bytes = excel.encode();
  final encodeTime = sw.elapsedMilliseconds;

  printMemory('AFTER ENCODE');
  print('Encode time: ${encodeTime}ms');

  if (bytes != null) {
    print('Output size: ${(bytes.length / 1024 / 1024).toStringAsFixed(1)} MB');

    // Step 6: Save to file
    final outFile = File('benchmark_output.xlsx');
    outFile.writeAsBytesSync(bytes);
    print('Saved to: ${outFile.path}');
  }

  printMemory('FINAL');
  print('');
  print('=== Summary ===');
  print('Rows: $rowCount x Cols: $colCount = ${rowCount * colCount} cells');
  print('Fill time: ${fillTime}ms');
  print('Encode time: ${encodeTime}ms');
  if (bytes != null) {
    print('File size: ${(bytes.length / 1024 / 1024).toStringAsFixed(1)} MB');
  }
}
