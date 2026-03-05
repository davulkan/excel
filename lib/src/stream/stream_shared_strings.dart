import 'dart:convert';

/// Collects unique strings for the streaming writer's shared string table.
class SharedStringCollector {
  final Map<String, int> _map = {};
  final List<String> _list = [];
  int _nextIndex = 0;

  /// Returns the shared string index for [value], adding it if new.
  int getOrAdd(String value) {
    final existing = _map[value];
    if (existing != null) return existing;
    final idx = _nextIndex++;
    _map[value] = idx;
    _list.add(value);
    return idx;
  }

  int get length => _list.length;

  /// Generates the complete xl/sharedStrings.xml as UTF-8 bytes.
  List<int> buildXmlBytes() {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write(
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
    buf.write(' count="$_nextIndex" uniqueCount="$_nextIndex">');
    for (final s in _list) {
      buf.write('<si><t');
      if (s.isNotEmpty &&
          (s[0] == ' ' ||
              s[s.length - 1] == ' ' ||
              s.contains('\n') ||
              s.contains('\t'))) {
        buf.write(' xml:space="preserve"');
      }
      buf.write('>');
      buf.write(escapeXmlForStream(s));
      buf.write('</t></si>');
    }
    buf.write('</sst>');
    return utf8.encode(buf.toString());
  }
}

/// Escapes text content for XML: &, <, >
String escapeXmlForStream(String text) {
  if (!text.contains('&') && !text.contains('<') && !text.contains('>')) {
    return text;
  }
  return text
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;');
}
