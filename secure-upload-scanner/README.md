# secure-upload-scanner

A single-file, dependency-light Python scanner that decides whether an uploaded
file is safe to accept — using **general rules** (name, extension, content, size),
not signatures, so it catches new threats the same way it catches old ones.
Built for file-ingestion pipelines, including LLM/AI apps that read user files.

## Catches

- Path traversal + extension allow-listing (documents, spreadsheets, images only).
- Disguised executables (`MZ` / `ELF` magic), double extensions.
- **Office (xlsx/docx):** VBA macros, OLE embeds, external data links, font-smuggling,
  XXE / entity bombs, formula/DDE injection, zip bombs, files hidden inside the archive.
- **PDF:** active content (`/JavaScript`, `/Launch`, `/EmbeddedFile`, `/XFA`), incl. hex-obfuscated names.
- **CSV:** formula injection (`=cmd|…`, DDE, `HYPERLINK`/`WEBSERVICE` exfiltration).
- **Images:** decompression bombs, absurd dimensions, and injection hidden in
  EXIF/PNG-text metadata (checked in the *actual* metadata, not raw pixels — so real
  photos aren't falsely rejected), plus appended polyglot payloads.
- EICAR test-virus signature anywhere in the file.

## Usage

```python
from upload_scanner import validate_upload, scan_disk_file

ok, reason = validate_upload(filename, data_bytes)   # for uploads
ok, reason = scan_disk_file("/path/to/file")          # for files already on disk

if not ok:
    reject(reason)
```

Only `Pillow` is needed for image scanning (optional — it degrades gracefully if
absent). Everything else is the Python standard library.

## Why

Most upload validators check the extension and stop. This one opens Office archives
and PDFs, inspects image metadata, and blocks the injection and data-exfiltration
tricks a naive check misses — while staying a single file you can drop into any project.

## License

MIT — see `LICENSE`.
