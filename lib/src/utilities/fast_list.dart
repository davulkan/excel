part of excel;

// A helper class to optimized the usage of Maps
class FastList<K> {
  Set<K> _index = <K>{};

  FastList();

  FastList.from(FastList<K> other) : _index = Set<K>.from(other._index);

  void add(K key) {
    _index.add(key);
  }

  bool contains(K key) {
    return _index.contains(key);
  }

  void remove(K key) {
    _index.remove(key);
  }

  void clear() {
    _index = <K>{};
  }

  List<K> get keys => _index.toList();

  bool get isNotEmpty => _index.isNotEmpty;
}
