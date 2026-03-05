part of excel;

class _SharedStringsMaintainer {
  final Map<SharedString, _IndexingHolder> _map =
      <SharedString, _IndexingHolder>{};
  final Map<String, SharedString> _mapString = <String, SharedString>{};
  final List<SharedString> _list = <SharedString>[];
  int _index = 0;

  _SharedStringsMaintainer._();

  SharedString? tryFind(String val) {
    return _mapString[val];
  }

  SharedString addFromString(String val) {
    final newSharedString = SharedString(
        node: XmlElement(XmlName('si'), [], [
      XmlElement(XmlName('t'),
          [XmlAttribute(XmlName("space", "xml"), "preserve")], [XmlText(val)]),
    ]));

    add(newSharedString, val);
    return newSharedString;
  }

  void add(SharedString val, String key) {
    _map[val]?.increaseCount();
    _map.putIfAbsent(val, () {
      _mapString[key] = val;
      _list.add(val);
      return _IndexingHolder(_index++);
    });
  }

  int indexOf(SharedString val) {
    return _map[val] != null ? _map[val]!.index : -1;
  }

  SharedString? value(int i) {
    if (i < _list.length) {
      return _list[i];
    } else {
      return null;
    }
  }

  void clear() {
    _index = 0;
    _list.clear();
    _map.clear();
    _mapString.clear();
  }
}

class _IndexingHolder {
  final int index;
  late int count;
  _IndexingHolder(this.index, [int _count = 1]) {
    this.count = _count;
  }

  void increaseCount() {
    this.count += 1;
  }
}

class SharedString {
  final XmlElement node;
  final String _cachedStringValue;
  final int _hashCode;

  SharedString._({required this.node, required String stringVal, required int hash})
      : _cachedStringValue = stringVal,
        _hashCode = hash;

  factory SharedString({required XmlElement node}) {
    final String stringVal = _computeStringValue(node);
    final int hash = _computeStructuralHash(node);
    return SharedString._(node: node, stringVal: stringVal, hash: hash);
  }

  /// Extracts the text content from the XML node, excluding <rPh> children.
  static String _computeStringValue(XmlElement node) {
    var buffer = StringBuffer();
    node.findAllElements('t').forEach((child) {
      if (child.parentElement == null ||
          child.parentElement!.name.local != 'rPh') {
        buffer.write(Parser._parseValue(child));
      }
    });
    return buffer.toString();
  }

  /// Computes a structural hash by walking the XML tree without serializing
  /// to a string. Incorporates element names, attribute key-value pairs, and
  /// text content so that structurally different nodes (e.g. plain text vs
  /// rich text with formatting runs) produce different hashes.
  static int _computeStructuralHash(XmlElement element) {
    int hash = element.name.local.hashCode;
    for (final XmlAttribute attr in element.attributes) {
      hash = hash ^ attr.name.local.hashCode ^ attr.value.hashCode;
    }
    for (final XmlNode child in element.children) {
      if (child is XmlElement) {
        hash = hash * 31 + _computeStructuralHash(child);
      } else if (child is XmlText) {
        hash = hash * 31 + child.value.hashCode;
      }
    }
    return hash;
  }

  @override
  String toString() {
    assert(false,
        'prefer stringValue over SharedString.toString() in development');
    return stringValue;
  }

  String get stringValue => _cachedStringValue;

  @override
  int get hashCode => _hashCode;

  @override
  operator ==(Object other) {
    return other is SharedString &&
        other._hashCode == _hashCode &&
        other._cachedStringValue == _cachedStringValue;
  }

  bool matches(String value) {
    return value.isNotEmpty && value == _cachedStringValue;
  }
}
