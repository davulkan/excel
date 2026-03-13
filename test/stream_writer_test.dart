import 'dart:convert';
import 'dart:io';
import 'package:archive/archive.dart';
import 'package:excel/excel_streaming.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

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

      expect(row[0]?.value, equals(DateCellValue(year: 2024, month: 6, day: 15)));
      expect(row[1]?.value, equals(TimeCellValue(hour: 14, minute: 30)));
      expect(row[2]?.value, equals(DateTimeCellValue(
          year: 2024, month: 6, day: 15, hour: 14, minute: 30)));
      expect(row[0]?.cellStyle?.numberFormat, equals(NumFormat.defaultDate));
      expect(row[1]?.cellStyle?.numberFormat, equals(NumFormat.defaultTime));
      expect(row[2]?.cellStyle?.numberFormat, equals(NumFormat.defaultDateTime));
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
      final excel = Excel.decodeBytes(bytes);
      final row = excel['Styled'].rows[0];

      expect((row[0]?.value as TextCellValue).value.toString(), equals('Bold'));
      expect(row[0]?.cellStyle?.isBold, isTrue);
      expect((row[1]?.value as TextCellValue).value.toString(), equals('Red'));
      expect(row[1]?.cellStyle?.fontColor.colorHex, equals('FFFF0000'));
    });

    test('number format preserved with explicit style', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('NumFmt');

      sheet.appendRow(
        [IntCellValue(42), DoubleCellValue(15.99)],
        styles: [
          CellStyle(numberFormat: NumFormat.standard_11),
          CellStyle(numberFormat: NumFormat.defaultFloat),
        ],
      );

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final row = excel['NumFmt'].rows[0];

      expect(row[0]?.value, equals(IntCellValue(42)));
      expect(row[0]?.cellStyle?.numberFormat, equals(NumFormat.standard_11));
      expect(row[1]?.value, equals(DoubleCellValue(15.99)));
      expect(row[1]?.cellStyle?.numberFormat, equals(NumFormat.defaultFloat));
    });

    test('custom number format round-trip', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('CustomFmt');

      final format1 = CustomNumericNumFormat(formatCode: r'0.00%');
      final format2 = CustomNumericNumFormat(formatCode: r'#,##0.00');

      sheet.appendRow(
        [DoubleCellValue(0.15), DoubleCellValue(123456.789)],
        styles: [
          CellStyle(numberFormat: format1),
          CellStyle(numberFormat: format2),
        ],
      );

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final row = excel['CustomFmt'].rows[0];

      expect(row[0]?.value, equals(DoubleCellValue(0.15)));
      expect(row[0]?.cellStyle?.numberFormat, equals(format1));
      expect(row[1]?.value, equals(DoubleCellValue(123456.789)));
      expect(row[1]?.cellStyle?.numberFormat, equals(format2));
    });

    test('border styles round-trip', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Borders');

      final thinBorder = Border(borderStyle: BorderStyle.Thin);
      final mediumRedBorder = Border(
        borderStyle: BorderStyle.Medium,
        borderColorHex: 'FFFF0000'.excelColor,
      );

      sheet.appendRow(
        [TextCellValue('A'), TextCellValue('B')],
        styles: [
          CellStyle()
            ..leftBorder = thinBorder
            ..rightBorder = thinBorder
            ..topBorder = thinBorder
            ..bottomBorder = thinBorder,
          CellStyle()
            ..leftBorder = mediumRedBorder
            ..bottomBorder = mediumRedBorder,
        ],
      );

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final row = excel['Borders'].rows[0];

      expect(row[0]?.cellStyle?.leftBorder, equals(thinBorder));
      expect(row[0]?.cellStyle?.rightBorder, equals(thinBorder));
      expect(row[0]?.cellStyle?.topBorder, equals(thinBorder));
      expect(row[0]?.cellStyle?.bottomBorder, equals(thinBorder));
      expect(row[1]?.cellStyle?.leftBorder, equals(mediumRedBorder));
      expect(row[1]?.cellStyle?.bottomBorder, equals(mediumRedBorder));
    });

    test('background color round-trip', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('BgColor');

      sheet.appendRow(
        [TextCellValue('Yellow')],
        styles: [CellStyle()..backgroundColor = ExcelColor.fromHexString('#FFFF00')],
      );

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final row = excel['BgColor'].rows[0];

      expect(row[0]?.cellStyle?.backgroundColor.colorHex, equals('FFFFFF00'));
    });

    test('font properties round-trip', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Fonts');

      sheet.appendRow(
        [TextCellValue('Italic'), TextCellValue('Big')],
        styles: [
          CellStyle(italic: true),
          CellStyle(fontSize: 16, fontFamily: 'Arial'),
        ],
      );

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final row = excel['Fonts'].rows[0];

      expect(row[0]?.cellStyle?.isItalic, isTrue);
      expect(row[1]?.cellStyle?.fontSize, equals(16));
      expect(row[1]?.cellStyle?.fontFamily, equals('Arial'));
    });

    test('appendRow throws on styles length mismatch', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('S');

      expect(
        () => sheet.appendRow(
          [TextCellValue('a'), TextCellValue('b')],
          styles: [null],
        ),
        throwsArgumentError,
      );
      writer.dispose();
    });

    test('appendRow throws after sheet finalized by close', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('S');
      sheet.appendRow([TextCellValue('x')]);
      writer.close();

      expect(
        () => sheet.appendRow([TextCellValue('y')]),
        throwsStateError,
      );
    });

    test('StreamSheet rowCount and maxColumns', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Metrics');

      expect(sheet.rowCount, equals(0));
      expect(sheet.maxColumns, equals(0));

      sheet.appendRow([TextCellValue('a'), TextCellValue('b'), TextCellValue('c')]);
      expect(sheet.rowCount, equals(1));
      expect(sheet.maxColumns, equals(3));

      sheet.appendRow([TextCellValue('x')]);
      expect(sheet.rowCount, equals(2));
      expect(sheet.maxColumns, equals(3));

      writer.dispose();
    });

    test('empty sheet produces valid xlsx', () {
      final writer = ExcelStreamWriter();
      writer.addSheet('Empty');

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);

      expect(excel.tables.keys, contains('Empty'));
      expect(excel['Empty'].rows, isEmpty);
    });

    test('alignment styles round-trip', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Align');

      sheet.appendRow(
        [
          TextCellValue('Right'),
          TextCellValue('Top'),
          TextCellValue('Wrapped'),
          TextCellValue('Rotated'),
        ],
        styles: [
          CellStyle()..horizontalAlignment = HorizontalAlign.Right,
          CellStyle()..verticalAlignment = VerticalAlign.Top,
          CellStyle()..wrap = TextWrapping.WrapText,
          CellStyle()..rotation = 45,
        ],
      );

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final row = excel['Align'].rows[0];

      expect(row[0]?.cellStyle?.horizontalAlignment,
          equals(HorizontalAlign.Right));
      expect(row[1]?.cellStyle?.verticalAlignment, equals(VerticalAlign.Top));
      expect(row[2]?.cellStyle?.wrap, equals(TextWrapping.WrapText));
      expect(row[3]?.cellStyle?.rotation, equals(45));
    });

    test('shared strings XML uniqueCount matches actual unique strings', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('SSTest');

      // 2 unique strings: 'hello', 'world'
      // 'hello' is referenced 3 times, 'world' once
      sheet.appendRow([TextCellValue('hello'), TextCellValue('world')]);
      sheet.appendRow([TextCellValue('hello'), TextCellValue('hello')]);

      final bytes = writer.close();

      final archive = ZipDecoder().decodeBytes(bytes);
      final ssFile = archive.findFile('xl/sharedStrings.xml')!;
      final doc = XmlDocument.parse(utf8.decode(ssFile.content));
      final sst = doc.findAllElements('sst').first;

      expect(sst.getAttributeNode('uniqueCount')!.value, equals('2'));
    });

    test('empty sheet alongside sheet with data', () {
      final writer = ExcelStreamWriter();
      writer.addSheet('Empty');
      final data = writer.addSheet('Data');
      data.appendRow([TextCellValue('value'), IntCellValue(42)]);

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);

      expect(excel.tables.keys, containsAll(['Empty', 'Data']));
      expect(excel['Empty'].rows, isEmpty);
      expect(excel['Data'].rows.length, equals(1));
      expect(
          (excel['Data'].rows[0][0]?.value as TextCellValue).value.toString(),
          equals('value'));
      expect((excel['Data'].rows[0][1]?.value as IntCellValue).value,
          equals(42));
    });

    test('handles wide rows with many columns', () {
      const colCount = 500;
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('Wide');

      sheet.appendRow(
          List.generate(colCount, (i) => TextCellValue('col$i')));
      sheet.appendRow(
          List.generate(colCount, (i) => IntCellValue(i)));

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);
      final rows = excel['Wide'].rows;

      expect(rows.length, equals(2));
      expect(rows[0].length, equals(colCount));
      expect(
          (rows[0][0]?.value as TextCellValue).value.toString(), equals('col0'));
      expect(
          (rows[0][colCount - 1]?.value as TextCellValue).value.toString(),
          equals('col${colCount - 1}'));
      expect((rows[1][colCount - 1]?.value as IntCellValue).value,
          equals(colCount - 1));
    });

    test('unicode in sheet names and cell values', () {
      final writer = ExcelStreamWriter();
      final sheet = writer.addSheet('データ');

      sheet.appendRow([
        TextCellValue('日本語テスト'),
        TextCellValue('Привет мир'),
        TextCellValue('中文测试'),
        TextCellValue('العربية'),
      ]);

      final bytes = writer.close();
      final excel = Excel.decodeBytes(bytes);

      expect(excel.tables.keys, contains('データ'));
      final row = excel['データ'].rows[0];
      expect(
          (row[0]?.value as TextCellValue).value.toString(), equals('日本語テスト'));
      expect(
          (row[1]?.value as TextCellValue).value.toString(), equals('Привет мир'));
      expect(
          (row[2]?.value as TextCellValue).value.toString(), equals('中文测试'));
      expect(
          (row[3]?.value as TextCellValue).value.toString(), equals('العربية'));
    });
  });
}
