part of excel;

Archive _cloneArchive(
  Archive archive,
  Map<String, ArchiveFile> _filesToOverwrite, {
  String? excludedFile,
}) {
  var clone = Archive();
  archive.files.forEach(
    (file) {
      if (!file.isFile) return;
      if (excludedFile != null &&
          file.name.toLowerCase() == excludedFile.toLowerCase()) {
        return;
      }

      ArchiveFile copy;
      if (_filesToOverwrite.containsKey(file.name)) {
        copy = _filesToOverwrite[file.name]!;
      } else {
        var content = file.content as Uint8List;
        var compress = !_noCompression.contains(file.name);
        copy = ArchiveFile(file.name, content.length, content)
          ..compress = compress;
      }
      clone.addFile(copy);
    },
  );
  return clone;
}
