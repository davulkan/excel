import 'dart:convert';

import 'stream_shared_strings.dart' show escapeXmlForStream;

/// Generates minimal XML scaffolding files for the streaming writer.
class StreamXmlTemplates {
  /// [Content_Types].xml
  static List<int> contentTypes(List<String> sheetNames) {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">');
    buf.write(
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    buf.write(
        '<Default Extension="xml" ContentType="application/xml"/>');
    buf.write(
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>');
    for (var i = 0; i < sheetNames.length; i++) {
      buf.write(
          '<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
    }
    buf.write(
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>');
    buf.write(
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>');
    buf.write('</Types>');
    return utf8.encode(buf.toString());
  }

  /// _rels/.rels
  static List<int> topLevelRels() {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
    buf.write(
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>');
    buf.write('</Relationships>');
    return utf8.encode(buf.toString());
  }

  /// xl/_rels/workbook.xml.rels
  static List<int> workbookRels(List<String> sheetNames) {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
    for (var i = 0; i < sheetNames.length; i++) {
      buf.write(
          '<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>');
    }
    final ssId = sheetNames.length + 1;
    buf.write(
        '<Relationship Id="rId$ssId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>');
    final stId = sheetNames.length + 2;
    buf.write(
        '<Relationship Id="rId$stId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>');
    buf.write('</Relationships>');
    return utf8.encode(buf.toString());
  }

  /// xl/workbook.xml
  static List<int> workbook(List<String> sheetNames) {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write(
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
    buf.write('<sheets>');
    for (var i = 0; i < sheetNames.length; i++) {
      buf.write(
          '<sheet name="${escapeXmlForStream(sheetNames[i])}" sheetId="${i + 1}" r:id="rId${i + 1}"/>');
    }
    buf.write('</sheets>');
    buf.write('</workbook>');
    return utf8.encode(buf.toString());
  }
}
