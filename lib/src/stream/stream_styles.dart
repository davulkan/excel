import 'dart:convert';

import 'package:excel/excel.dart';

import 'stream_shared_strings.dart' show escapeXmlForStream;

/// Manages styles for the streaming writer.
///
/// Pre-populates default styles for each value type. Users can register
/// custom styles via [getOrRegister]. Generates styles.xml at close time.
class StreamStyleManager {
  final Map<CellStyle, int> _cache = {};
  final List<CellStyle> _list = [];

  late final int _defaultGeneralIdx;
  late final int _defaultDateIdx;
  late final int _defaultTimeIdx;
  late final int _defaultDateTimeIdx;

  StreamStyleManager() {
    _defaultGeneralIdx = _addStyle(CellStyle());
    _defaultDateIdx =
        _addStyle(CellStyle(numberFormat: NumFormat.defaultDate));
    _defaultTimeIdx =
        _addStyle(CellStyle(numberFormat: NumFormat.defaultTime));
    _defaultDateTimeIdx =
        _addStyle(CellStyle(numberFormat: NumFormat.defaultDateTime));
  }

  int _addStyle(CellStyle style) {
    final idx = _list.length;
    _list.add(style);
    _cache[style] = idx;
    return idx;
  }

  /// Returns the style index for [style], registering it if new.
  int getOrRegister(CellStyle style) {
    final existing = _cache[style];
    if (existing != null) return existing;
    return _addStyle(style);
  }

  /// Returns the default style index for a given cell value type.
  int defaultForValue(CellValue? value) => switch (value) {
        null ||
        TextCellValue() ||
        FormulaCellValue() ||
        IntCellValue() ||
        DoubleCellValue() ||
        BoolCellValue() =>
          _defaultGeneralIdx,
        DateCellValue() => _defaultDateIdx,
        TimeCellValue() => _defaultTimeIdx,
        DateTimeCellValue() => _defaultDateTimeIdx,
      };

  /// Generates the complete xl/styles.xml as UTF-8 bytes.
  List<int> buildStylesXmlBytes() {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write(
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' mc:Ignorable="x14ac x16r2 xr"'
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
        ' xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"'
        ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">');

    // Collect unique components
    final fonts = <_FontKey>[];
    final fontIndex = <_FontKey, int>{};
    final fills = <String>[];
    final fillIndex = <String, int>{};
    final borders = <_BorderKey>[];
    final borderIndex = <_BorderKey, int>{};
    final numFmtEntries = <int, String>{};

    // Default fill entries required by Excel
    fills.add('none');
    fillIndex['none'] = 0;
    fills.add('gray125');
    fillIndex['gray125'] = 1;

    // Default font
    final defaultFont = _FontKey.fromStyle(CellStyle());
    fonts.add(defaultFont);
    fontIndex[defaultFont] = 0;

    // Default border
    final defaultBorder = _BorderKey.fromStyle(CellStyle());
    borders.add(defaultBorder);
    borderIndex[defaultBorder] = 0;

    for (final style in _list) {
      final fk = _FontKey.fromStyle(style);
      if (!fontIndex.containsKey(fk)) {
        fontIndex[fk] = fonts.length;
        fonts.add(fk);
      }

      final bg = style.backgroundColor.colorHex;
      if (!fillIndex.containsKey(bg)) {
        fillIndex[bg] = fills.length;
        fills.add(bg);
      }

      final bk = _BorderKey.fromStyle(style);
      if (!borderIndex.containsKey(bk)) {
        borderIndex[bk] = borders.length;
        borders.add(bk);
      }

      final nf = style.numberFormat;
      if (nf is CustomNumFormat) {
        if (!numFmtEntries.values.contains(nf.formatCode)) {
          numFmtEntries[164 + numFmtEntries.length] = nf.formatCode;
        }
      }
    }

    // numFmts
    if (numFmtEntries.isNotEmpty) {
      buf.write('<numFmts count="${numFmtEntries.length}">');
      for (final e in numFmtEntries.entries) {
        buf.write(
            '<numFmt numFmtId="${e.key}" formatCode="${escapeXmlForStream(e.value)}"/>');
      }
      buf.write('</numFmts>');
    }

    // fonts
    buf.write('<fonts count="${fonts.length}" x14ac:knownFonts="1">');
    for (final f in fonts) {
      buf.write('<font>');
      if (f.bold) buf.write('<b/>');
      if (f.italic) buf.write('<i/>');
      if (f.underline == Underline.Single) {
        buf.write('<u/>');
      } else if (f.underline == Underline.Double) {
        buf.write('<u val="double"/>');
      }
      buf.write('<sz val="${f.fontSize ?? 11}"/>');
      if (f.fontColorHex != 'FF000000') {
        buf.write('<color rgb="${f.fontColorHex}"/>');
      } else {
        buf.write('<color theme="1"/>');
      }
      buf.write(
          '<name val="${escapeXmlForStream(f.fontFamily ?? 'Calibri')}"/>');
      if (f.fontScheme != FontScheme.Unset) {
        final scheme =
            f.fontScheme == FontScheme.Major ? 'major' : 'minor';
        buf.write('<scheme val="$scheme"/>');
      }
      buf.write('</font>');
    }
    buf.write('</fonts>');

    // fills
    buf.write('<fills count="${fills.length}">');
    for (final fill in fills) {
      if (fill == 'none' || fill == 'gray125' || fill == 'lightGray') {
        buf.write('<fill><patternFill patternType="$fill"/></fill>');
      } else {
        buf.write(
            '<fill><patternFill patternType="solid">'
            '<fgColor rgb="$fill"/><bgColor rgb="$fill"/>'
            '</patternFill></fill>');
      }
    }
    buf.write('</fills>');

    // borders
    buf.write('<borders count="${borders.length}">');
    for (final b in borders) {
      buf.write('<border');
      if (b.diagonalBorderDown) buf.write(' diagonalDown="1"');
      if (b.diagonalBorderUp) buf.write(' diagonalUp="1"');
      buf.write('>');
      _writeBorderSide(buf, 'left', b.leftBorder);
      _writeBorderSide(buf, 'right', b.rightBorder);
      _writeBorderSide(buf, 'top', b.topBorder);
      _writeBorderSide(buf, 'bottom', b.bottomBorder);
      _writeBorderSide(buf, 'diagonal', b.diagonalBorder);
      buf.write('</border>');
    }
    buf.write('</borders>');

    // cellStyleXfs
    buf.write('<cellStyleXfs count="1">');
    buf.write('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>');
    buf.write('</cellStyleXfs>');

    // cellXfs
    buf.write('<cellXfs count="${_list.length}">');
    for (final style in _list) {
      final fk = _FontKey.fromStyle(style);
      final bg = style.backgroundColor.colorHex;
      final bk = _BorderKey.fromStyle(style);

      final nf = style.numberFormat;
      final int numFmtId;
      if (nf is StandardNumFormat) {
        numFmtId = nf.numFmtId;
      } else if (nf is CustomNumFormat) {
        numFmtId = numFmtEntries.entries
            .firstWhere((e) => e.value == nf.formatCode)
            .key;
      } else {
        numFmtId = 0;
      }

      final fIdx = fontIndex[fk] ?? 0;
      final fiIdx = fillIndex[bg] ?? 0;
      final bIdx = borderIndex[bk] ?? 0;

      buf.write('<xf numFmtId="$numFmtId" fontId="$fIdx"'
          ' fillId="$fiIdx" borderId="$bIdx" xfId="0"');

      if (fiIdx > 0) buf.write(' applyFill="1"');
      if (fIdx > 0) buf.write(' applyFont="1"');
      if (numFmtId > 0) buf.write(' applyNumberFormat="1"');
      if (bIdx > 0) buf.write(' applyBorder="1"');

      final ha = style.horizontalAlignment;
      final va = style.verticalAlignment;
      final tw = style.wrap;
      final rot = style.rotation;

      if (ha != HorizontalAlign.Left ||
          va != VerticalAlign.Bottom ||
          tw != null ||
          rot != 0) {
        buf.write(' applyAlignment="1"><alignment');
        if (ha != HorizontalAlign.Left) {
          buf.write(
              ' horizontal="${ha == HorizontalAlign.Right ? 'right' : 'center'}"');
        }
        if (va != VerticalAlign.Bottom) {
          buf.write(
              ' vertical="${va == VerticalAlign.Top ? 'top' : 'center'}"');
        }
        if (tw != null) {
          buf.write(
              ' ${tw == TextWrapping.Clip ? 'shrinkToFit' : 'wrapText'}="1"');
        }
        if (rot != 0) buf.write(' textRotation="$rot"');
        buf.write('/></xf>');
      } else {
        buf.write('/>');
      }
    }
    buf.write('</cellXfs>');

    // cellStyles
    buf.write('<cellStyles count="1">');
    buf.write('<cellStyle name="Normal" xfId="0" builtinId="0"/>');
    buf.write('</cellStyles>');

    buf.write('<dxfs count="0"/>');
    buf.write('<tableStyles count="0" defaultTableStyle="TableStyleMedium2"'
        ' defaultPivotStyle="PivotStyleLight16"/>');

    buf.write('</styleSheet>');
    return utf8.encode(buf.toString());
  }

  void _writeBorderSide(StringBuffer buf, String name, Border border) {
    final style = border.borderStyle;
    final color = border.borderColorHex;
    if (style == null && color == null) {
      buf.write('<$name/>');
    } else {
      buf.write('<$name');
      if (style != null) buf.write(' style="${style.style}"');
      if (color != null) {
        buf.write('><color rgb="$color"/></$name>');
      } else {
        buf.write('/>');
      }
    }
  }
}

/// Lightweight font key for deduplication (avoids depending on private _FontStyle).
class _FontKey {
  final bool bold;
  final bool italic;
  final String fontColorHex;
  final Underline underline;
  final int? fontSize;
  final String? fontFamily;
  final FontScheme fontScheme;

  _FontKey({
    required this.bold,
    required this.italic,
    required this.fontColorHex,
    required this.underline,
    required this.fontSize,
    required this.fontFamily,
    required this.fontScheme,
  });

  factory _FontKey.fromStyle(CellStyle s) => _FontKey(
        bold: s.isBold,
        italic: s.isItalic,
        fontColorHex: s.fontColor.colorHex,
        underline: s.underline,
        fontSize: s.fontSize,
        fontFamily: s.fontFamily,
        fontScheme: s.fontScheme,
      );

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is _FontKey &&
          bold == other.bold &&
          italic == other.italic &&
          fontColorHex == other.fontColorHex &&
          underline == other.underline &&
          fontSize == other.fontSize &&
          fontFamily == other.fontFamily &&
          fontScheme == other.fontScheme;

  @override
  int get hashCode => Object.hash(
      bold, italic, fontColorHex, underline, fontSize, fontFamily, fontScheme);
}

/// Lightweight border key for deduplication.
class _BorderKey {
  final Border leftBorder;
  final Border rightBorder;
  final Border topBorder;
  final Border bottomBorder;
  final Border diagonalBorder;
  final bool diagonalBorderUp;
  final bool diagonalBorderDown;

  _BorderKey({
    required this.leftBorder,
    required this.rightBorder,
    required this.topBorder,
    required this.bottomBorder,
    required this.diagonalBorder,
    required this.diagonalBorderUp,
    required this.diagonalBorderDown,
  });

  factory _BorderKey.fromStyle(CellStyle s) => _BorderKey(
        leftBorder: s.leftBorder,
        rightBorder: s.rightBorder,
        topBorder: s.topBorder,
        bottomBorder: s.bottomBorder,
        diagonalBorder: s.diagonalBorder,
        diagonalBorderUp: s.diagonalBorderUp,
        diagonalBorderDown: s.diagonalBorderDown,
      );

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is _BorderKey &&
          leftBorder == other.leftBorder &&
          rightBorder == other.rightBorder &&
          topBorder == other.topBorder &&
          bottomBorder == other.bottomBorder &&
          diagonalBorder == other.diagonalBorder &&
          diagonalBorderUp == other.diagonalBorderUp &&
          diagonalBorderDown == other.diagonalBorderDown;

  @override
  int get hashCode => Object.hash(leftBorder, rightBorder, topBorder,
      bottomBorder, diagonalBorder, diagonalBorderUp, diagonalBorderDown);
}
