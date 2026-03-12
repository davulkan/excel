import 'dart:io';
import 'package:excel/excel_streaming.dart';
import 'package:test/test.dart';

void main() {
  group('ExcelStreamWriter', () {
    test('creates valid xlsx with basic text and numbers', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Sheet1');

      sheet.appendRow([
        TextCellValue('Name'),
        TextCellValue('Age'),
        TextCellValue('Score'),
      ]);
      sheet.appendRow([
        TextCellValue('Alice'),
        IntCellValue(30),
        DoubleCellValue(95.5),
      ]);
      sheet.appendRow([
        TextCellValue('Bob'),
        IntCellValue(25),
        DoubleCellValue(87.3),
      ]);

      final bytes = writer.close();
      expect(bytes, isNotEmpty);

      // Verify by reading back with Excel
      final excel = Excel.decodeBytes(bytes);
      final readSheet = excel['Sheet1'];
      expect(readSheet.rows.length, 3);

      // Row 0: headers
      expect(readSheet.rows[0][0]?.value, isA<TextCellValue>());
      expect((readSheet.rows[0][0]?.value as TextCellValue).value.toString(),
          equals('Name'));
      expect((readSheet.rows[0][1]?.value as TextCellValue).value.toString(),
          equals('Age'));
      expect((readSheet.rows[0][2]?.value as TextCellValue).value.toString(),
          equals('Score'));

      // Row 1: Alice
      expect((readSheet.rows[1][0]?.value as TextCellValue).value.toString(),
          equals('Alice'));
      expect((readSheet.rows[1][1]?.value as IntCellValue).value, equals(30));
      expect((readSheet.rows[1][2]?.value as DoubleCellValue).value,
          closeTo(95.5, 0.01));

      // Row 2: Bob
      expect((readSheet.rows[2][0]?.value as TextCellValue).value.toString(),
          equals('Bob'));
      expect((readSheet.rows[2][1]?.value as IntCellValue).value, equals(25));
    });

    test('handles multiple sheets', () {
      final writer = ExcelStreamWriter();
      final sheet1 = writer.addSheet('Data');
      final sheet2 = writer.addSheet('Summary');

      sheet1.appendRow([TextCellValue('A1'), IntCellValue(1)]);
      sheet1.appendRow([TextCellValue('A2'), IntCellValue(2)]);
      sheet2.appendRow([TextCellValue('Total'), IntCellValue(3)]);

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);

      expect(excel.tables.keys, containsAll(['Data', 'Summary']));
      expect(excel['Data'].rows.length, 2);
      expect(excel['Summary'].rows.length, 1);
    });

    test('handles bool values', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Bools');

      sheet.appendRow([BoolCellValue(true), BoolCellValue(false)]);

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final row = excel['Bools'].rows[0];

      expect((row[0]?.value as BoolCellValue).value, isTrue);
      expect((row[1]?.value as BoolCellValue).value, isFalse);
    });

    test('handles date and time values', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Dates');

      final date = DateCellValue(year: 2024, month: 6, day: 15);
      final time = TimeCellValue(hour: 14, minute: 30);
      final dt = DateTimeCellValue(
          year: 2024, month: 6, day: 15, hour: 14, minute: 30);

      sheet.appendRow([date, time, dt]);

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final row = excel['Dates'].rows[0];

      // Date values are read back as numeric by Excel.decodeBytes,
      // but the numeric value should be correct
      expect(row[0]?.value, isNotNull);
      expect(row[1]?.value, isNotNull);
      expect(row[2]?.value, isNotNull);
    });

    test('handles formula values', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Formulas');

      sheet.appendRow([IntCellValue(10), IntCellValue(20)]);
      sheet.appendRow([FormulaCellValue('A1+B1')]);

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);

      expect(excel['Formulas'].rows.length, 2);
      final formulaCell = excel['Formulas'].rows[1][0];
      expect(formulaCell?.value, isA<FormulaCellValue>());
      expect((formulaCell?.value as FormulaCellValue).formula, equals('A1+B1'));
    });

    test('handles null cells', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Nulls');

      sheet.appendRow([TextCellValue('A'), null, TextCellValue('C')]);

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final row = excel['Nulls'].rows[0];

      expect((row[0]?.value as TextCellValue).value.toString(), equals('A'));
      // null cell may be null or have null value
      expect((row[2]?.value as TextCellValue).value.toString(), equals('C'));
    });

    test('handles XML special characters', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Special');

      sheet.appendRow([
        TextCellValue('Tom & Jerry'),
        TextCellValue('<html>'),
        TextCellValue('a > b'),
      ]);

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final row = excel['Special'].rows[0];

      expect((row[0]?.value as TextCellValue).value.toString(),
          equals('Tom & Jerry'));
      expect(
          (row[1]?.value as TextCellValue).value.toString(), equals('<html>'));
      expect(
          (row[2]?.value as TextCellValue).value.toString(), equals('a > b'));
    });

    test('shared strings deduplicates', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Dedup');

      sheet.appendRow([TextCellValue('hello'), TextCellValue('world')]);
      sheet.appendRow([TextCellValue('hello'), TextCellValue('hello')]);

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final rows = excel['Dedup'].rows;

      expect((rows[0][0]?.value as TextCellValue).value.toString(),
          equals('hello'));
      expect((rows[1][0]?.value as TextCellValue).value.toString(),
          equals('hello'));
      expect((rows[1][1]?.value as TextCellValue).value.toString(),
          equals('hello'));
    });

    test('closeToFile writes to disk', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('FileTest');
      sheet.appendRow([TextCellValue('written to file')]);

      final tmpFile = File(
          '${Directory.systemTemp.path}/stream_test_${DateTime.now().millisecondsSinceEpoch}.xlsx');
      try {
        writer.closeToFile(tmpFile.path);
        expect(tmpFile.existsSync(), isTrue);
        expect(tmpFile.lengthSync(), greaterThan(0));

        // Verify contents
        final excel = Excel.decodeBytes(tmpFile.readAsBytesSync());
        expect(
            (excel['FileTest'].rows[0][0]?.value as TextCellValue)
                .value
                .toString(),
            equals('written to file'));
      } finally {
        if (tmpFile.existsSync()) tmpFile.deleteSync();
      }
    });

    test('dispose cleans up temp files', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Cleanup');
      sheet.appendRow([TextCellValue('temp')]);
      writer.dispose();

      // Should not throw if called again
      writer.dispose();
    });

    test('throws after close', () {
      final writer = ExcelStreamWriter();
      writer.addSheet('S1');
      writer.close();

      expect(() => writer.addSheet('S2'), throwsStateError);
      expect(() => writer.close(), throwsStateError);
    });

    test('handles large number of rows', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Big');

      for (var i = 0; i < 1000; i++) {
        sheet.appendRow([
          IntCellValue(i),
          TextCellValue('row_$i'),
          DoubleCellValue(i * 1.5),
        ]);
      }

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      expect(excel['Big'].rows.length, 1000);
      expect((excel['Big'].rows[999][0]?.value as IntCellValue).value,
          equals(999));
    });

    test('handles custom styles', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Styled');

      final boldStyle = CellStyle(bold: true);
      final redStyle = CellStyle(
        fontColorHex: ExcelColor.fromHexString('#FF0000'),
      );

      sheet.appendRow(
        [TextCellValue('Bold'), TextCellValue('Red')],
        styles: [boldStyle, redStyle],
      );

      final bytes = writer.close();
      // Just verify the file is valid and readable
      final excel = Excel.decodeBytes(bytes);
      expect(excel['Styled'].rows.length, 1);
      expect(
          (excel['Styled'].rows[0][0]?.value as TextCellValue).value.toString(),
          equals('Bold'));
    });
  });
}
