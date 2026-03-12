part of excel;

class Save {
  final Excel _excel;
  late Map<String, ArchiveFile> _archiveFiles;
  late List<CellStyle> _innerCellStyle;
  final Parser parser;
  final String? creator;
  final String? description;

  /// Dictionaries for cell style position lookups instead of O(n) search in styles list.
  final _innerStylePosMap = HashMap<CellStyle, int>();
  final _upperStylePosMap = HashMap<CellStyle, int>();

  /// Tracks which XML file keys use streaming serialization (sheet files).
  final Set<String> _xmlFilesForSheets = {};

  Save._(this._excel, this.parser, {this.creator, this.description}) {
    _archiveFiles = <String, ArchiveFile>{};
    _innerCellStyle = <CellStyle>[];
  }

  void _addNewColumn(XmlElement columns, int min, int max, double width) {
    columns.children.add(XmlElement(XmlName('col'), [
      XmlAttribute(XmlName('min'), (min + 1).toString()),
      XmlAttribute(XmlName('max'), (max + 1).toString()),
      XmlAttribute(XmlName('width'), width.toStringAsFixed(2)),
      XmlAttribute(XmlName('bestFit'), "1"),
      XmlAttribute(XmlName('customWidth'), "1"),
    ], []));
  }

  double _calcAutoFitColumnWidth(Sheet sheet, int column) {
    var maxNumOfCharacters = 0;
    sheet._sheetData.forEach((key, value) {
      if (value.containsKey(column) &&
          value[column]!.value is! FormulaCellValue) {
        maxNumOfCharacters =
            max(value[column]!.value.toString().length, maxNumOfCharacters);
      }
    });

    return ((maxNumOfCharacters * 7.0 + 9.0) / 7.0 * 256).truncate() / 256;
  }

  /// Registers a shared string for a TextCellValue and returns its index.
  /// Does NOT build XmlElement DOM — uses the existing SharedStrings API.
  int _registerSharedString(TextCellValue val) {
    SharedString? sharedString = _excel._sharedStrings.tryFind(val.value);
    if (sharedString != null) {
      _excel._sharedStrings.add(sharedString, val.value);
    } else {
      sharedString = _excel._sharedStrings.addFromString(val.value);
    }
    return _excel._sharedStrings.indexOf(sharedString);
  }

  /// Pre-registers all shared strings for a sheet without building DOM.
  void _registerSharedStringsForSheet(Sheet sheet) {
    sheet._sheetData.forEach((_, columnMap) {
      columnMap.forEach((_, data) {
        if (data.value is TextCellValue) {
          _registerSharedString(data.value as TextCellValue);
        }
      });
    });
  }

  /// Computes the style index (the 's' attribute value) for a cell.
  /// Returns -1 if no style should be applied.
  int _computeStyleIndex(
      String sheetName, int columnIndex, int rowIndex, CellValue? value) {
    CellStyle? cellStyle = _excel
        ._sheetMap[sheetName]?._sheetData[rowIndex]?[columnIndex]?.cellStyle;

    // When styles are being tracked, synthesize a minimal CellStyle for cells
    // without explicit style but with values that need a specific numFormat.
    if (_excel._styleChanges && cellStyle == null && value != null) {
      final numFmt = NumFormat.defaultFor(value);
      if (numFmt != NumFormat.standard_0) {
        cellStyle = CellStyle(numberFormat: numFmt);
      }
    }

    if (_excel._styleChanges && cellStyle != null) {
      int upperLevelPos = _upperStylePosMap[cellStyle] ?? -1;
      if (upperLevelPos == -1) {
        int lowerLevelPos = _innerStylePosMap[cellStyle] ?? -1;
        if (lowerLevelPos != -1) {
          upperLevelPos = lowerLevelPos + _excel._cellStyleList.length;
        } else {
          upperLevelPos = 0;
        }
      }
      return upperLevelPos;
    }

    String rC = getCellId(columnIndex, rowIndex);
    if (_excel._cellStyleReferenced.containsKey(sheetName) &&
        _excel._cellStyleReferenced[sheetName]!.containsKey(rC)) {
      return _excel._cellStyleReferenced[sheetName]![rC]!;
    }

    return -1;
  }

  /// Writes a single <c> element for a cell directly to the StringBuffer without tons of XmlNode dart objects.
  void _writeCellXml(StringBuffer buf, String sheetName, int columnIndex,
      int rowIndex, Data data) {
    final CellValue? value = data.value;

    String rC = getCellId(columnIndex, rowIndex);
    int styleIndex =
        _computeStyleIndex(sheetName, columnIndex, rowIndex, value);

    // Handle null values: write an empty <c> element if there's a style, otherwise skip
    if (value == null) {
      if (styleIndex >= 0) buf.write('<c r="$rC" s="$styleIndex"/>');
      return;
    }

    final NumFormat numFormat =
        data.cellStyle?.numberFormat ?? NumFormat.defaultFor(value);

    switch (value) {
      case TextCellValue():
        int ssIndex = _excel._sharedStrings
            .indexOf(_excel._sharedStrings.tryFind(value.value)!);
        buf.write('<c r="$rC" t="s"');
        if (styleIndex >= 0) buf.write(' s="$styleIndex"');
        buf.write('><v>$ssIndex</v></c>');

      case FormulaCellValue():
        buf.write('<c r="$rC"');
        if (styleIndex >= 0) buf.write(' s="$styleIndex"');

        buf.write('><f>');
        buf.write(_escapeXmlText(value.formula));
        buf.write('</f><v></v></c>');

      case IntCellValue():
        final String v = switch (numFormat) {
          NumericNumFormat() => numFormat.writeInt(value),
          _ => value.value.toString(),
        };
        buf.write('<c r="$rC"');
        if (styleIndex >= 0) buf.write(' s="$styleIndex"');
        buf.write('><v>$v</v></c>');

      case DoubleCellValue():
        final String v = switch (numFormat) {
          NumericNumFormat() => numFormat.writeDouble(value),
          _ => value.value.toString(),
        };
        buf.write('<c r="$rC"');
        if (styleIndex >= 0) buf.write(' s="$styleIndex"');
        buf.write('><v>$v</v></c>');

      case BoolCellValue():
        buf.write('<c r="$rC" t="b"');
        if (styleIndex >= 0) buf.write(' s="$styleIndex"');
        buf.write('><v>${value.value ? '1' : '0'}</v></c>');

      case DateTimeCellValue():
        final String v = switch (numFormat) {
          DateTimeNumFormat() => numFormat.writeDateTime(value),
          _ => throw Exception(
              '$numFormat does not work for ${value.runtimeType}'),
        };
        buf.write('<c r="$rC"');
        if (styleIndex >= 0) buf.write(' s="$styleIndex"');
        buf.write('><v>$v</v></c>');

      case DateCellValue():
        final String v = switch (numFormat) {
          DateTimeNumFormat() => numFormat.writeDate(value),
          _ => throw Exception(
              '$numFormat does not work for ${value.runtimeType}'),
        };
        buf.write('<c r="$rC"');
        if (styleIndex >= 0) buf.write(' s="$styleIndex"');
        buf.write('><v>$v</v></c>');

      case TimeCellValue():
        final String v = switch (numFormat) {
          TimeNumFormat() => numFormat.writeTime(value),
          _ => throw Exception(
              '$numFormat does not work for ${value.runtimeType}'),
        };
        buf.write('<c r="$rC"');
        if (styleIndex >= 0) buf.write(' s="$styleIndex"');
        buf.write('><v>$v</v></c>');
    }
  }

  /// Writes the full <sheetData>...</sheetData> block to the StringBuffer without tons of XmlNode dart objects.
  void _buildSheetDataXml(StringBuffer buf, Sheet sheet) {
    buf.write('<sheetData>');

    final customHeights = sheet.getRowHeights;
    final sortedRows = sheet._sheetData.keys.toList()..sort();

    for (final rowIndex in sortedRows) {
      final columnMap = sheet._sheetData[rowIndex];
      if (columnMap == null || columnMap.isEmpty) continue;

      double? height = customHeights[rowIndex];
      int? level = sheet.getRowLevel(rowIndex);

      buf.write('<row r="rowIndex + 1"');
      if (height != null) {
        buf.write(' ht="${height.toStringAsFixed(2)}" customHeight="1"');
      }
      if (level != null) buf.write(' outlineLevel="$level"');
      buf.write('>');

      final sortedCols = columnMap.keys.toList()..sort();
      for (final columnIndex in sortedCols) {
        final data = columnMap[columnIndex];
        if (data == null) continue;

        _writeCellXml(buf, sheet.sheetName, columnIndex, rowIndex, data);
      }

      buf.write('</row>');
    }

    buf.write('</sheetData>');
  }

  /// Serializes a complete sheet XML using DOM for everything except sheetData,
  /// which is streamed via StringBuffer for performance.
  String _serializeSheetXml(String xmlFileKey, Sheet sheet) {
    final XmlDocument xmlDoc = _excel._xmlFiles[xmlFileKey]!;

    // Find sheetData element and clear its children (we'll replace this empty element later)
    final sheetDataElement = xmlDoc.findAllElements('sheetData').first;
    sheetDataElement.children.clear();

    // Serialize the DOM to string — sheetData will be empty: <sheetData/> or <sheetData></sheetData>
    String xmlString = xmlDoc.toString();

    final sheetDataBuf = StringBuffer();
    _buildSheetDataXml(sheetDataBuf, sheet);

    // Replace the empty sheetData tag with content from our sheetDataBuf
    // Handle both self-closing and open/close forms
    return xmlString.replaceFirst(
      RegExp(r'<sheetData\s*/>|<sheetData>\s*</sheetData>'),
      sheetDataBuf.toString(),
    );
  }

  /// Writing Font Color in [xl/styles.xml] from the Cells of the sheets.

  void _processStylesFile() {
    _innerCellStyle = <CellStyle>[];
    _innerStylePosMap.clear();
    _upperStylePosMap.clear();
    List<String> innerPatternFill = <String>[];
    List<_FontStyle> innerFontStyle = <_FontStyle>[];
    List<_BorderSet> innerBorderSet = <_BorderSet>[];

    // Build _upperStyleCache from existing _cellStyleList
    for (int i = 0; i < _excel._cellStyleList.length; i++) {
      _upperStylePosMap[_excel._cellStyleList[i]] = i;
    }

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      sheetObject._sheetData.forEach((_, columnMap) {
        columnMap.forEach((_, dataObject) {
          CellStyle? cs = dataObject.cellStyle;
          // Synthesize style for cells with values that need a non-General numFormat
          if (cs == null && dataObject.value != null) {
            final numFmt = NumFormat.defaultFor(dataObject.value);
            if (numFmt != NumFormat.standard_0) {
              cs = CellStyle(numberFormat: numFmt);
            }
          }
          if (cs != null) {
            if (!_upperStylePosMap.containsKey(cs) &&
                !_innerStylePosMap.containsKey(cs)) {
              int idx = _innerCellStyle.length;
              _innerCellStyle.add(cs);
              _innerStylePosMap[cs] = idx;
            }
          }
        });
      });
    });

    _innerCellStyle.forEach((cellStyle) {
      _FontStyle _fs = _FontStyle(
          bold: cellStyle.isBold,
          italic: cellStyle.isItalic,
          fontColorHex: cellStyle.fontColor,
          underline: cellStyle.underline,
          fontSize: cellStyle.fontSize,
          fontFamily: cellStyle.fontFamily,
          fontScheme: cellStyle.fontScheme);

      /// If `-1` is returned then it indicates that `_fontStyle` is not present in the `_fs`
      if (_fontStyleIndex(_excel._fontStyleList, _fs) == -1 &&
          _fontStyleIndex(innerFontStyle, _fs) == -1) {
        innerFontStyle.add(_fs);
      }

      /// Filling the inner usable extra list of background color
      String backgroundColor = cellStyle.backgroundColor.colorHex;
      if (!_excel._patternFill.contains(backgroundColor) &&
          !innerPatternFill.contains(backgroundColor)) {
        innerPatternFill.add(backgroundColor);
      }

      final _bs = _createBorderSetFromCellStyle(cellStyle);
      if (!_excel._borderSetList.contains(_bs) &&
          !innerBorderSet.contains(_bs)) {
        innerBorderSet.add(_bs);
      }
    });

    XmlElement fonts =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('fonts').first;

    var fontAttribute = fonts.getAttributeNode('count');
    if (fontAttribute != null) {
      fontAttribute.value =
          '${_excel._fontStyleList.length + innerFontStyle.length}';
    } else {
      fonts.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._fontStyleList.length + innerFontStyle.length}'));
    }

    innerFontStyle.forEach((fontStyleElement) {
      fonts.children.add(XmlElement(XmlName('font'), [], [
        /// putting color
        if (fontStyleElement._fontColorHex != null &&
            fontStyleElement._fontColorHex!.colorHex != "FF000000")
          XmlElement(XmlName('color'), [
            XmlAttribute(
                XmlName('rgb'), fontStyleElement._fontColorHex!.colorHex)
          ], []),

        /// putting bold
        if (fontStyleElement.isBold) XmlElement(XmlName('b'), [], []),

        /// putting italic
        if (fontStyleElement.isItalic) XmlElement(XmlName('i'), [], []),

        /// putting single underline
        if (fontStyleElement.underline != Underline.None &&
            fontStyleElement.underline == Underline.Single)
          XmlElement(XmlName('u'), [], []),

        /// putting double underline
        if (fontStyleElement.underline != Underline.None &&
            fontStyleElement.underline != Underline.Single &&
            fontStyleElement.underline == Underline.Double)
          XmlElement(
              XmlName('u'), [XmlAttribute(XmlName('val'), 'double')], []),

        /// putting fontFamily
        if (fontStyleElement.fontFamily != null &&
            fontStyleElement.fontFamily!.toLowerCase().toString() != 'null' &&
            fontStyleElement.fontFamily != '' &&
            fontStyleElement.fontFamily!.isNotEmpty)
          XmlElement(XmlName('name'), [
            XmlAttribute(XmlName('val'), fontStyleElement.fontFamily.toString())
          ], []),

        /// putting fontScheme
        if (fontStyleElement.fontScheme != FontScheme.Unset)
          XmlElement(XmlName('scheme'), [
            XmlAttribute(
                XmlName('val'),
                switch (fontStyleElement.fontScheme) {
                  FontScheme.Major => "major",
                  _ => "minor"
                })
          ], []),

        /// putting fontSize
        if (fontStyleElement.fontSize != null &&
            fontStyleElement.fontSize.toString().isNotEmpty)
          XmlElement(XmlName('sz'), [
            XmlAttribute(XmlName('val'), fontStyleElement.fontSize.toString())
          ], []),
      ]));
    });

    XmlElement fills =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('fills').first;

    var fillAttribute = fills.getAttributeNode('count');

    if (fillAttribute != null) {
      fillAttribute.value =
          '${_excel._patternFill.length + innerPatternFill.length}';
    } else {
      fills.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._patternFill.length + innerPatternFill.length}'));
    }

    innerPatternFill.forEach((color) {
      if (color.length >= 2) {
        if (color.substring(0, 2).toUpperCase() == 'FF') {
          fills.children.add(XmlElement(XmlName('fill'), [], [
            XmlElement(XmlName('patternFill'), [
              XmlAttribute(XmlName('patternType'), 'solid')
            ], [
              XmlElement(XmlName('fgColor'),
                  [XmlAttribute(XmlName('rgb'), color)], []),
              XmlElement(
                  XmlName('bgColor'), [XmlAttribute(XmlName('rgb'), color)], [])
            ])
          ]));
        } else if (color == "none" ||
            color == "gray125" ||
            color == "lightGray") {
          fills.children.add(XmlElement(XmlName('fill'), [], [
            XmlElement(XmlName('patternFill'),
                [XmlAttribute(XmlName('patternType'), color)], [])
          ]));
        }
      } else {
        _damagedExcel(
            text:
                "Corrupted Styles Found. Can't process further, Open up issue in github.");
      }
    });

    XmlElement borders =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('borders').first;
    var borderAttribute = borders.getAttributeNode('count');

    if (borderAttribute != null) {
      borderAttribute.value =
          '${_excel._borderSetList.length + innerBorderSet.length}';
    } else {
      borders.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._borderSetList.length + innerBorderSet.length}'));
    }

    innerBorderSet.forEach((border) {
      var borderElement = XmlElement(XmlName('border'));
      if (border.diagonalBorderDown) {
        borderElement.attributes
            .add(XmlAttribute(XmlName('diagonalDown'), '1'));
      }
      if (border.diagonalBorderUp) {
        borderElement.attributes.add(XmlAttribute(XmlName('diagonalUp'), '1'));
      }
      final Map<String, Border> borderMap = {
        'left': border.leftBorder,
        'right': border.rightBorder,
        'top': border.topBorder,
        'bottom': border.bottomBorder,
        'diagonal': border.diagonalBorder,
      };
      for (var key in borderMap.keys) {
        final borderValue = borderMap[key]!;

        final element = XmlElement(XmlName(key));
        final style = borderValue.borderStyle;
        if (style != null) {
          element.attributes.add(XmlAttribute(XmlName('style'), style.style));
        }
        final color = borderValue.borderColorHex;
        if (color != null) {
          element.children.add(XmlElement(
              XmlName('color'), [XmlAttribute(XmlName('rgb'), color)]));
        }
        borderElement.children.add(element);
      }

      borders.children.add(borderElement);
    });

    final styleSheet = _excel._xmlFiles['xl/styles.xml']!;

    XmlElement celx = styleSheet.findAllElements('cellXfs').first;
    var cellAttribute = celx.getAttributeNode('count');

    if (cellAttribute != null) {
      cellAttribute.value =
          '${_excel._cellStyleList.length + _innerCellStyle.length}';
    } else {
      celx.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._cellStyleList.length + _innerCellStyle.length}'));
    }

    _innerCellStyle.forEach((cellStyle) {
      String backgroundColor = cellStyle.backgroundColor.colorHex;

      _FontStyle _fs = _FontStyle(
          bold: cellStyle.isBold,
          italic: cellStyle.isItalic,
          fontColorHex: cellStyle.fontColor,
          underline: cellStyle.underline,
          fontSize: cellStyle.fontSize,
          fontFamily: cellStyle.fontFamily);

      HorizontalAlign horizontalAlign = cellStyle.horizontalAlignment;
      VerticalAlign verticalAlign = cellStyle.verticalAlignment;
      int rotation = cellStyle.rotation;
      TextWrapping? textWrapping = cellStyle.wrap;
      int backgroundIndex = innerPatternFill.indexOf(backgroundColor),
          fontIndex = _fontStyleIndex(innerFontStyle, _fs);
      _BorderSet _bs = _createBorderSetFromCellStyle(cellStyle);
      int borderIndex = innerBorderSet.indexOf(_bs);

      final numberFormat = cellStyle.numberFormat;
      final int numFmtId = switch (numberFormat) {
        StandardNumFormat() => numberFormat.numFmtId,
        CustomNumFormat() => _excel._numFormats.findOrAdd(numberFormat),
      };

      var attributes = <XmlAttribute>[
        XmlAttribute(XmlName('borderId'),
            '${borderIndex == -1 ? 0 : borderIndex + _excel._borderSetList.length}'),
        XmlAttribute(XmlName('fillId'),
            '${backgroundIndex == -1 ? 0 : backgroundIndex + _excel._patternFill.length}'),
        XmlAttribute(XmlName('fontId'),
            '${fontIndex == -1 ? 0 : fontIndex + _excel._fontStyleList.length}'),
        XmlAttribute(XmlName('numFmtId'), numFmtId.toString()),
        XmlAttribute(XmlName('xfId'), '0'),
      ];

      if ((_excel._patternFill.contains(backgroundColor) ||
              innerPatternFill.contains(backgroundColor)) &&
          backgroundColor != "none" &&
          backgroundColor != "gray125" &&
          backgroundColor.toLowerCase() != "lightgray") {
        attributes.add(XmlAttribute(XmlName('applyFill'), '1'));
      }

      if (_fontStyleIndex(_excel._fontStyleList, _fs) != -1 &&
          _fontStyleIndex(innerFontStyle, _fs) != -1) {
        attributes.add(XmlAttribute(XmlName('applyFont'), '1'));
      }

      var children = <XmlElement>[];

      if (horizontalAlign != HorizontalAlign.Left ||
          textWrapping != null ||
          verticalAlign != VerticalAlign.Bottom ||
          rotation != 0) {
        attributes.add(XmlAttribute(XmlName('applyAlignment'), '1'));
        var childAttributes = <XmlAttribute>[];

        if (textWrapping != null) {
          childAttributes.add(XmlAttribute(
              XmlName(textWrapping == TextWrapping.Clip
                  ? 'shrinkToFit'
                  : 'wrapText'),
              '1'));
        }

        if (verticalAlign != VerticalAlign.Bottom) {
          String ver = verticalAlign == VerticalAlign.Top ? 'top' : 'center';
          childAttributes.add(XmlAttribute(XmlName('vertical'), '$ver'));
        }

        if (horizontalAlign != HorizontalAlign.Left) {
          String hor =
              horizontalAlign == HorizontalAlign.Right ? 'right' : 'center';
          childAttributes.add(XmlAttribute(XmlName('horizontal'), '$hor'));
        }
        if (rotation != 0) {
          childAttributes
              .add(XmlAttribute(XmlName('textRotation'), '$rotation'));
        }

        children.add(XmlElement(XmlName('alignment'), childAttributes, []));
      }

      celx.children.add(XmlElement(XmlName('xf'), attributes, children));
    });

    final customNumberFormats = _excel._numFormats._map.entries
        .map<MapEntry<int, CustomNumFormat>?>((e) {
          final format = e.value;
          if (format is! CustomNumFormat) {
            return null;
          }
          return MapEntry<int, CustomNumFormat>(e.key, format);
        })
        .whereNotNull()
        .sorted((a, b) => a.key.compareTo(b.key));
    if (customNumberFormats.isNotEmpty) {
      var numFmtsElement = styleSheet
          .findAllElements('numFmts')
          .whereType<XmlElement>()
          .firstOrNull;
      int count;
      if (numFmtsElement == null) {
        numFmtsElement = XmlElement(XmlName('numFmts'));

        ///FIX: if no default numFormats were added in styles.xml - customNumFormats were added in wrong place,
        styleSheet
            .findElements('styleSheet')
            .first
            .children
            .insert(0, numFmtsElement);
      }
      count = int.parse(numFmtsElement.getAttribute('count') ?? '0');

      for (var numFormat in customNumberFormats) {
        final numFmtIdString = numFormat.key.toString();
        final formatCode = numFormat.value.formatCode;
        var numFmtElement = numFmtsElement.children
            .whereType<XmlElement>()
            .firstWhereOrNull((node) =>
                node.name.local == 'numFmt' &&
                node.getAttribute('numFmtId') == numFmtIdString);
        if (numFmtElement == null) {
          numFmtElement = XmlElement(
              XmlName('numFmt'),
              [
                XmlAttribute(XmlName('numFmtId'), numFmtIdString),
                XmlAttribute(XmlName('formatCode'), formatCode),
              ],
              [],
              true);
          numFmtsElement.children.add(numFmtElement);
          count++;
        } else if ((numFmtElement.getAttribute('formatCode') ?? '') !=
            formatCode) {
          numFmtElement.setAttribute('formatCode', formatCode);
        }
      }

      numFmtsElement.setAttribute('count', count.toString());
    }
  }

  void _saveCore() {
    final XmlBuilder builder = XmlBuilder();
    builder.processing(
        'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('cp:coreProperties', nest: () {
      builder.attribute('xmlns:cp',
          'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
      builder.attribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
      builder.attribute('xmlns:dcterms', 'http://purl.org/dc/terms/');
      builder.attribute('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
      builder.attribute(
          'xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');

      builder.element('cp:keywords', nest: creator);

      builder.element('dc:description', nest: description);
    });
    _excel._xmlFiles['docProps/core.xml'] = builder.buildDocument();
  }

  void _saveTopLevelRelation() {
    final XmlBuilder builder = XmlBuilder();

    builder.processing('xml', 'version="1.0"');
    builder.element('Relationships', nest: () {
      builder.attribute('xmlns',
          'http://schemas.openxmlformats.org/package/2006/relationships');

      builder.element('Relationship', nest: () {
        builder.attribute('Id', 'rId1');
        builder.attribute('Type',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument');
        builder.attribute('Target', 'xl/workbook.xml');
      });

      builder.element('Relationship', nest: () {
        builder.attribute('Id', 'rId2');
        builder.attribute('Type',
            'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties');
        builder.attribute('Target', 'docProps/core.xml');
      });
    });
    _excel._xmlFiles['_rels/.rels'] = builder.buildDocument();
  }

  _addCoreProps() {
    var builder = XmlBuilder();
    builder.element('Override', nest: () {
      builder.attribute('PartName', '/docProps/core.xml');
      builder.attribute('ContentType',
          'application/vnd.openxmlformats-package.core-properties+xml');
    });
    _excel._xmlFiles['[Content_Types].xml']
        ?.findAllElements('Types')
        .first
        .children
        .add(builder.buildFragment());

    _saveCore();
    _saveTopLevelRelation();
  }

  List<int>? _save() {
    if (_excel._styleChanges) {
      _processStylesFile();
    }
    _setSheetElements();
    if (_excel._defaultSheet != null) {
      _setDefaultSheet(_excel._defaultSheet);
    }
    _setSharedStrings();

    if (_excel._mergeChanges) {
      _setMerge();
    }

    if (_excel._rtlChanges) {
      _setRTL();
    }
    if (creator != null && description != null) {
      _addCoreProps();
    }

    for (var xmlFile in _excel._xmlFiles.keys) {
      String xml;
      if (_xmlFilesForSheets.contains(xmlFile)) {
        // Optimized serialization for sheet files
        final sheetName = _excel._xmlSheetId.entries
            .firstWhere((e) => e.value == xmlFile)
            .key;

        xml = _serializeSheetXml(xmlFile, _excel._sheetMap[sheetName]!);
      } else {
        xml = _excel._xmlFiles[xmlFile].toString();
      }
      var content = utf8.encode(xml);
      if (xmlFile == 'docProps/core.xml') {
        _excel._archive
            .addFile(ArchiveFile('docProps/core.xml', content.length, content));
      }
      _archiveFiles[xmlFile] = ArchiveFile(xmlFile, content.length, content);
    }
    return ZipEncoder()
        .encode(_buildOutputArchive(_excel._archive, _archiveFiles));
  }

  void _setColumns(Sheet sheetObject, XmlDocument xmlFile) {
    final columnElements = xmlFile.findAllElements('cols');

    if (sheetObject.getColumnWidths.isEmpty &&
        sheetObject.getColumnAutoFits.isEmpty) {
      if (columnElements.isEmpty) {
        return;
      }

      final columns = columnElements.first;
      final worksheet = xmlFile.findAllElements('worksheet').first;
      worksheet.children.remove(columns);
      return;
    }

    if (columnElements.isEmpty) {
      final worksheet = xmlFile.findAllElements('worksheet').first;
      final sheetData = xmlFile.findAllElements('sheetData').first;
      final index = worksheet.children.indexOf(sheetData);

      worksheet.children.insert(index, XmlElement(XmlName('cols'), [], []));
    }

    var columns = columnElements.first;

    if (columns.children.isNotEmpty) {
      columns.children.clear();
    }

    final autoFits = sheetObject.getColumnAutoFits;
    final customWidths = sheetObject.getColumnWidths;

    final columnCount = max(
        autoFits.isEmpty ? 0 : autoFits.keys.reduce(max) + 1,
        customWidths.isEmpty ? 0 : customWidths.keys.reduce(max) + 1);

    List<double> columnWidths = <double>[];

    double defaultColumnWidth =
        sheetObject.defaultColumnWidth ?? _excelDefaultColumnWidth;

    for (var index = 0; index < columnCount; index++) {
      double width = defaultColumnWidth;

      if (autoFits.containsKey(index) && (!customWidths.containsKey(index))) {
        width = _calcAutoFitColumnWidth(sheetObject, index);
      } else {
        if (customWidths.containsKey(index)) {
          width = customWidths[index]!;
        }
      }

      columnWidths.add(width);

      _addNewColumn(columns, index, index, width);
    }
  }

  bool _setDefaultSheet(String? sheetName) {
    if (sheetName == null || _excel._xmlFiles['xl/workbook.xml'] == null) {
      return false;
    }
    List<XmlElement> sheetList =
        _excel._xmlFiles['xl/workbook.xml']!.findAllElements('sheet').toList();
    XmlElement elementFound = XmlElement(XmlName(''));

    int position = -1;
    for (int i = 0; i < sheetList.length; i++) {
      var _sheetName = sheetList[i].getAttribute('name');
      if (_sheetName != null && _sheetName.toString() == sheetName) {
        elementFound = sheetList[i];
        position = i;
        break;
      }
    }

    if (position == -1) {
      return false;
    }
    if (position == 0) {
      return true;
    }

    _excel._xmlFiles['xl/workbook.xml']!
        .findAllElements('sheets')
        .first
        .children
      ..removeAt(position)
      ..insert(0, elementFound);

    String? expectedSheet = _excel._getDefaultSheet();

    return expectedSheet == sheetName;
  }

  void _setHeaderFooter(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    if (sheet == null) return;

    final xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
    if (xmlFile == null) return;

    final sheetXmlElement = xmlFile.findAllElements("worksheet").first;

    final results = sheetXmlElement.findAllElements("headerFooter");
    if (results.isNotEmpty) {
      sheetXmlElement.children.remove(results.first);
    }

    if (sheet.headerFooter == null) return;

    sheetXmlElement.children.add(sheet.headerFooter!.toXmlElement());
  }

  /// Writing the merged cells information into the excel properties files.
  void _setMerge() {
    _selfCorrectSpanMap(_excel);
    _excel._mergeChangeLook.forEach((s) {
      if (_excel._sheetMap[s] != null &&
          _excel._sheetMap[s]!._spanList.isNotEmpty &&
          _excel._xmlSheetId.containsKey(s) &&
          _excel._xmlFiles.containsKey(_excel._xmlSheetId[s])) {
        Iterable<XmlElement>? iterMergeElement = _excel
            ._xmlFiles[_excel._xmlSheetId[s]]
            ?.findAllElements('mergeCells');
        late XmlElement mergeElement;
        if (iterMergeElement?.isNotEmpty ?? false) {
          mergeElement = iterMergeElement!.first;
        } else {
          if ((_excel._xmlFiles[_excel._xmlSheetId[s]]
                      ?.findAllElements('worksheet')
                      .length ??
                  0) >
              0) {
            int index = _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('worksheet')
                .first
                .children
                .indexOf(_excel._xmlFiles[_excel._xmlSheetId[s]]!
                    .findAllElements("sheetData")
                    .first);
            if (index == -1) {
              _damagedExcel();
            }
            _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('worksheet')
                .first
                .children
                .insert(
                    index + 1,
                    XmlElement(XmlName('mergeCells'),
                        [XmlAttribute(XmlName('count'), '0')]));

            mergeElement = _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('mergeCells')
                .first;
          } else {
            _damagedExcel();
          }
        }

        List<String> _spannedItems =
            List<String>.from(_excel._sheetMap[s]!.spannedItems);

        [
          ['count', _spannedItems.length.toString()],
        ].forEach((value) {
          if (mergeElement.getAttributeNode(value[0]) == null) {
            mergeElement.attributes
                .add(XmlAttribute(XmlName(value[0]), value[1]));
          } else {
            mergeElement.getAttributeNode(value[0])!.value = value[1];
          }
        });

        mergeElement.children.clear();

        _spannedItems.forEach((value) {
          mergeElement.children.add(XmlElement(XmlName('mergeCell'),
              [XmlAttribute(XmlName('ref'), '$value')], []));
        });
      }
    });
  }

  void _setRTL() {
    _excel._rtlChangeLook.forEach((s) {
      var sheetObject = _excel._sheetMap[s];
      if (sheetObject != null &&
          _excel._xmlSheetId.containsKey(s) &&
          _excel._xmlFiles.containsKey(_excel._xmlSheetId[s])) {
        var itrSheetViewsRTLElement = _excel._xmlFiles[_excel._xmlSheetId[s]]
            ?.findAllElements('sheetViews');

        if (itrSheetViewsRTLElement?.isNotEmpty ?? false) {
          var itrSheetViewRTLElement = _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('sheetView');

          if (itrSheetViewRTLElement?.isNotEmpty ?? false) {
            /// clear all the children of the sheetViews here

            _excel._xmlFiles[_excel._xmlSheetId[s]]
                ?.findAllElements('sheetViews')
                .first
                .children
                .clear();
          }

          _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('sheetViews')
              .first
              .children
              .add(XmlElement(
                XmlName('sheetView'),
                [
                  if (sheetObject.isRTL)
                    XmlAttribute(XmlName('rightToLeft'), '1'),
                  XmlAttribute(XmlName('workbookViewId'), '0'),
                ],
              ));
        } else {
          _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('worksheet')
              .first
              .children
              .add(XmlElement(XmlName('sheetViews'), [], [
                XmlElement(
                  XmlName('sheetView'),
                  [
                    if (sheetObject.isRTL)
                      XmlAttribute(XmlName('rightToLeft'), '1'),
                    XmlAttribute(XmlName('workbookViewId'), '0'),
                  ],
                )
              ]));
        }
      }
    });
  }

  /// Writing the value of excel cells into the separate
  /// sharedStrings file so as to minimize the size of excel files.
  void _setSharedStrings() {
    var uniqueCount = 0;
    var count = 0;

    XmlElement shareString = _excel
        ._xmlFiles['xl/${_excel._sharedStringsTarget}']!
        .findAllElements('sst')
        .first;

    shareString.children.clear();

    _excel._sharedStrings._map.forEach((string, ss) {
      uniqueCount += 1;
      count += ss.count;

      shareString.children.add(string.node);
    });

    [
      ['count', '$count'],
      ['uniqueCount', '$uniqueCount']
    ].forEach((value) {
      if (shareString.getAttributeNode(value[0]) == null) {
        shareString.attributes.add(XmlAttribute(XmlName(value[0]), value[1]));
      } else {
        shareString.getAttributeNode(value[0])!.value = value[1];
      }
    });
  }

  /// Writing cell contained text into the excel sheet files.
  void _setSheetElements() {
    _excel._sharedStrings.clear();
    _xmlFilesForSheets.clear();

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      ///
      /// Create the sheet's xml file if it does not exist.
      if (_excel._sheets[sheetName] == null) {
        parser._createSheet(sheetName);
      }

      /// Clear the previous contents of the sheetData DOM element.
      /// The actual data will be written via streaming in _serializeSheetXml.
      if (_excel._sheets[sheetName]?.children.isNotEmpty ?? false) {
        _excel._sheets[sheetName]!.children.clear();
      }

      XmlDocument? xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
      if (xmlFile == null) return;

      // Set default column width and height for the sheet.
      double? defaultRowHeight = sheetObject.defaultRowHeight;
      double? defaultColumnWidth = sheetObject.defaultColumnWidth;

      XmlElement worksheetElement = xmlFile.findAllElements('worksheet').first;

      XmlElement? sheetFormatPrElement =
          worksheetElement.findElements('sheetFormatPr').isNotEmpty
              ? worksheetElement.findElements('sheetFormatPr').first
              : null;

      if (sheetFormatPrElement != null) {
        sheetFormatPrElement.attributes.clear();

        if (defaultRowHeight == null && defaultColumnWidth == null) {
          worksheetElement.children.remove(sheetFormatPrElement);
        }
      } else if (defaultRowHeight != null || defaultColumnWidth != null) {
        sheetFormatPrElement = XmlElement(XmlName('sheetFormatPr'), [], []);
        worksheetElement.children.insert(0, sheetFormatPrElement);
      }

      if (defaultRowHeight != null) {
        sheetFormatPrElement!.attributes.add(XmlAttribute(
            XmlName('defaultRowHeight'), defaultRowHeight.toStringAsFixed(2)));
      }
      if (defaultColumnWidth != null) {
        sheetFormatPrElement!.attributes.add(XmlAttribute(
            XmlName('defaultColWidth'), defaultColumnWidth.toStringAsFixed(2)));
      }

      _setColumns(sheetObject, xmlFile);

      // Register shared strings for this sheet (streaming approach —
      // no DOM row/cell creation needed).
      _registerSharedStringsForSheet(sheetObject);

      // Track this sheet file for streaming serialization
      final xmlFileKey = _excel._xmlSheetId[sheetName];
      if (xmlFileKey != null) _xmlFilesForSheets.add(xmlFileKey);

      _setHeaderFooter(sheetName);
    });
  }

  _BorderSet _createBorderSetFromCellStyle(CellStyle cellStyle) => _BorderSet(
        leftBorder: cellStyle.leftBorder,
        rightBorder: cellStyle.rightBorder,
        topBorder: cellStyle.topBorder,
        bottomBorder: cellStyle.bottomBorder,
        diagonalBorder: cellStyle.diagonalBorder,
        diagonalBorderUp: cellStyle.diagonalBorderUp,
        diagonalBorderDown: cellStyle.diagonalBorderDown,
      );
}
