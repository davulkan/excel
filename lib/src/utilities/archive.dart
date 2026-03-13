part of excel;

Archive _cloneArchive(
  Archive archive,
  Map<String, ArchiveFile> _filesToOverwrite, {
  String? excludedFile,
}) {
  var clone = Archive();
  archive.files.where((f) => f.isFile).forEach(
    (file) {
      if (excludedFile != null &&
          file.name.toLowerCase() == excludedFile.toLowerCase()) {
        return;
      }

      ArchiveFile copy;
      if (_filesToOverwrite.containsKey(file.name)) {
        copy = _filesToOverwrite[file.name]!;
      } else {
        var content = file.content;
        var compression = _noCompression.contains(file.name)
            ? CompressionType.none
            : CompressionType.deflate;
        copy = ArchiveFile(file.name, content.length, content)
          ..compression = compression;
      }
      clone.addFile(copy);
    },
  );
  return clone;
}
