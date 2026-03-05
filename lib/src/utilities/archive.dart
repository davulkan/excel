part of excel;

/// Legacy clone that decompresses and re-copies every file.
/// Still used by Excel.delete() which needs the excludedFile parameter.
Archive _cloneArchive(
  Archive archive,
  Map<String, ArchiveFile> _archiveFiles, {
  String? excludedFile,
}) {
  var clone = Archive();
  archive.files.forEach((file) {
    if (file.isFile) {
      if (excludedFile != null &&
          file.name.toLowerCase() == excludedFile.toLowerCase()) {
        return;
      }
      ArchiveFile copy;
      if (_archiveFiles.containsKey(file.name)) {
        copy = _archiveFiles[file.name]!;
      } else {
        var content = file.content as Uint8List;
        var compression = _noCompression.contains(file.name)
            ? CompressionType.none
            : CompressionType.deflate;
        copy = ArchiveFile(file.name, content.length, content)
          ..compression = compression;
      }
      clone.addFile(copy);
    }
  });
  return clone;
}

/// Optimized archive builder that reuses original ArchiveFile objects
/// for unchanged files instead of decompressing and copying them.
Archive _buildOutputArchive(
    Archive source, Map<String, ArchiveFile> updatedFiles) {
  final output = Archive();
  for (final file in source.files) {
    if (!file.isFile) continue;
    if (updatedFiles.containsKey(file.name)) {
      output.addFile(updatedFiles[file.name]!);
    } else {
      // Reuse original ArchiveFile — no decompression needed
      output.addFile(file);
    }
  }
  return output;
}
