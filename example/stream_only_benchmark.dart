import 'dart:io';
import 'package:excel/excel_streaming.dart';

/// Measures RSS memory of the current process
int getRssKb() => ProcessInfo.currentRss ~/ 1024;

String formatMb(int kb) => '${(kb / 1024).toStringAsFixed(1)} MB';

void main(List<String> args) {
  final rowCount = args.isNotEmpty ? int.parse(args[0]) : 10000000;
  final colCount = args.length > 1 ? int.parse(args[1]) : 20;

  print('');
  print('============================================');
  print('  ExcelStreamWriter — Large Scale Benchmark');
  print('============================================');
  print('Rows: $rowCount, Columns: $colCount');
  print('Total cells: ${rowCount * colCount}');
  print('');

  final rssStart = getRssKb();
  print('[START] RSS: ${formatMb(rssStart)}');

  final sw = Stopwatch()..start();
  final writer = ExcelStreamWriter();
  final sheet = writer.addSheet('Data');

  sheet.appendRow(List.generate(colCount, (i) => TextCellValue('Column_$i')));

  int peakFillRss = rssStart;
  for (var r = 0; r < rowCount; r++) {
    final row = List<CellValue?>.generate(colCount, (c) {
      if (c == 0) return IntCellValue(r);
      if (c == 1) return TextCellValue('Row_$r text data here for testing');
      if (c == 2) return DoubleCellValue(r * 1.5);
      if (c == 3) return BoolCellValue(r % 2 == 0);
      return IntCellValue(r * colCount + c);
    });
    sheet.appendRow(row);

    if ((r + 1) % 1000000 == 0) {
      final rss = getRssKb();
      if (rss > peakFillRss) peakFillRss = rss;
      print('[AFTER ${((r + 1) / 1000000).toStringAsFixed(0)}M ROWS] '
          'RSS: ${formatMb(rss)}  '
          'elapsed: ${(sw.elapsedMilliseconds / 1000).toStringAsFixed(1)}s');
    }
  }

  final fillTime = sw.elapsedMilliseconds;
  final rssFill = getRssKb();
  if (rssFill > peakFillRss) peakFillRss = rssFill;
  print('');
  print('[ALL ROWS FILLED] RSS: ${formatMb(rssFill)}');
  print('Fill time: ${(fillTime / 1000).toStringAsFixed(1)}s');
  print('');

  print('Closing (ZIP encode)...');
  sw.reset();
  sw.start();
  writer.closeToFile('stream_200m_output.xlsx');
  final encodeTime = sw.elapsedMilliseconds;

  final rssEncode = getRssKb();
  final fileSize = File('stream_200m_output.xlsx').lengthSync();

  print('[AFTER CLOSE] RSS: ${formatMb(rssEncode)}');
  print('Encode time: ${(encodeTime / 1000).toStringAsFixed(1)}s');
  print('');

  print('============================================');
  print('  Summary');
  print('============================================');
  print(
      'Cells:          ${rowCount * colCount} (${(rowCount * colCount / 1000000).toStringAsFixed(0)}M)');
  print('Fill time:      ${(fillTime / 1000).toStringAsFixed(1)}s');
  print('Encode time:    ${(encodeTime / 1000).toStringAsFixed(1)}s');
  print(
      'Total time:     ${((fillTime + encodeTime) / 1000).toStringAsFixed(1)}s');
  print('Peak fill RSS:  ${formatMb(peakFillRss)}');
  print('Peak encode RSS:${formatMb(rssEncode)}');
  print('File size:      ${(fileSize / 1024 / 1024).toStringAsFixed(1)} MB');
  print(
      'RAM/file ratio: ${(rssEncode / 1024 / (fileSize / 1024 / 1024)).toStringAsFixed(1)}x');

  File('stream_200m_output.xlsx').deleteSync();
}
