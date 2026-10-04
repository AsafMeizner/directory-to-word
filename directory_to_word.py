#!/usr/bin/env python3
"""Pack a directory into a self-contained Word .docx, storing data in the document body.

Format: DirectoryWordArchive v2.

v1 hid the payload in a custom XML part (customXml/item1.xml). Some apps and transfer
services strip unknown custom parts when they open and re-save a .docx, which silently
destroyed the archive. v2 writes every file's bytes as Base64 **paragraph text in
word/document.xml** - ordinary document content that cannot be dropped without throwing
away the visible document. The result still opens as a normal Word document.

Stdlib only. Usage:
    python directory_to_word.py <directory> [output.docx]
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import stat
import sys
import zlib
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

FORMAT_NAME = "DirectoryWordArchive"
FORMAT_VERSION = 2

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Markers are their own paragraphs. They use only "#", ".", letters and digits, and contain
# no spaces, so they survive run-splitting and whitespace changes by a re-saving program.
# Base64 data never contains "#", so a line starting with "###DWA2" is unambiguously a marker.
MARK = "###DWA2"
LINE_WIDTH = 76            # Base64 chars per line
LINES_PER_PARA = 240       # lines bundled into one paragraph (~18 KB of text)
CHUNK = LINE_WIDTH * LINES_PER_PARA


def xml_safe(text: str) -> str:
    return "".join(
        ch for ch in text
        if ch in "\t\n\r" or 0x20 <= ord(ch) <= 0xD7FF or 0xE000 <= ord(ch) <= 0xFFFD
    )


def xml_text(text: str) -> str:
    return escape(xml_safe(text), {'"': "&quot;"})


def posix_rel(path: Path, root: Path) -> str:
    return PurePosixPath(path.relative_to(root)).as_posix()


def file_mode(path: Path, follow_symlinks: bool = False) -> int:
    try:
        return stat.S_IMODE(path.stat(follow_symlinks=follow_symlinks).st_mode)
    except OSError:
        return 0


def file_mtime_ns(path: Path, follow_symlinks: bool = False) -> int:
    try:
        return path.stat(follow_symlinks=follow_symlinks).st_mtime_ns
    except OSError:
        return 0


def sha256_of(path: Path) -> tuple[str, int]:
    h = hashlib.sha256()
    size = 0
    with open(path, "rb") as fp:
        for block in iter(lambda: fp.read(1024 * 1024), b""):
            h.update(block)
            size += len(block)
    return h.hexdigest(), size


def collect_entries(root: Path, output_docx: Path) -> list[dict]:
    entries: list[dict] = []
    output_abs = output_docx.resolve(strict=False)

    def walk(directory: Path) -> None:
        try:
            children = sorted(os.scandir(directory), key=lambda item: item.name.lower())
        except PermissionError as exc:
            raise RuntimeError(f"Permission denied while reading: {directory}") from exc

        for item in children:
            path = Path(item.path)
            rel = posix_rel(path, root)
            try:
                if path.resolve(strict=False) == output_abs:
                    continue
            except OSError:
                pass

            if item.is_symlink():
                try:
                    target = os.readlink(path)
                except OSError as exc:
                    raise RuntimeError(f"Could not read symlink: {path}") from exc
                entries.append({
                    "path": rel, "type": "symlink", "target": target,
                    "mode": file_mode(path), "mtime_ns": file_mtime_ns(path),
                })
                continue

            if item.is_dir(follow_symlinks=False):
                entries.append({
                    "path": rel, "type": "directory",
                    "mode": file_mode(path), "mtime_ns": file_mtime_ns(path),
                })
                walk(path)
                continue

            if item.is_file(follow_symlinks=False):
                digest, size = sha256_of(path)
                entries.append({
                    "path": rel, "type": "file", "size": size, "sha256": digest,
                    "mode": file_mode(path), "mtime_ns": file_mtime_ns(path),
                    "_src": str(path),
                })

    walk(root)
    return entries


def para(text: str = "", style: str | None = None, preserve: bool = False) -> str:
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    if text == "":
        return f"<w:p>{ppr}</w:p>"
    sp = ' xml:space="preserve"' if preserve else ""
    return f'<w:p>{ppr}<w:r><w:t{sp}>{xml_text(text)}</w:t></w:r></w:p>'


def marker(name: str) -> str:
    return para(f"{MARK}.{name}", style="Code")


def b64_paragraphs(raw: bytes):
    """Yield document paragraphs carrying `raw` as wrapped Base64 text."""
    b64 = base64.b64encode(raw).decode("ascii")
    for i in range(0, len(b64), CHUNK):
        block = b64[i:i + CHUNK]
        lines = [block[j:j + LINE_WIDTH] for j in range(0, len(block), LINE_WIDTH)]
        yield para("\n".join(lines), style="Code", preserve=True)


def build_styles_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Code">
    <w:name w:val="Code"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="14"/></w:rPr>
  </w:style>
</w:styles>'''


def create_docx(source_dir: Path, output_docx: Path) -> None:
    source_dir = source_dir.resolve()
    output_docx = output_docx.resolve(strict=False)
    if not source_dir.is_dir():
        raise ValueError(f"Not a directory: {source_dir}")
    if output_docx.suffix.lower() != ".docx":
        raise ValueError("Output filename must end in .docx")
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    entries = collect_entries(source_dir, output_docx)
    files = [e for e in entries if e["type"] == "file"]
    dirs = [e for e in entries if e["type"] == "directory"]
    links = [e for e in entries if e["type"] == "symlink"]
    total = sum(e["size"] for e in files)

    manifest = {
        "format": FORMAT_NAME,
        "version": FORMAT_VERSION,
        "root_name": source_dir.name,
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "entries": [{k: v for k, v in e.items() if k != "_src"} for e in entries],
    }
    manifest_bytes = zlib.compress(
        json.dumps(manifest, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), 9)

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>'''
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>'''
    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Directory Archive - {xml_text(source_dir.name)}</dc:title>
  <dc:creator>directory_to_word.py</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''

    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
            'w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>')

    with ZipFile(output_docx, "w", compression=ZIP_DEFLATED, compresslevel=6) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/styles.xml", build_styles_xml())
        docx.writestr("word/_rels/document.xml.rels", doc_rels)
        docx.writestr("docProps/core.xml", core)

        with docx.open("word/document.xml", "w") as fp:
            def w(text: str) -> None:
                fp.write(text.encode("utf-8"))

            w('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
            w(f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>')
            # Human-readable header (ignored by the extractor).
            w(para(f"Directory Archive: {source_dir.name}", "Title"))
            w(para(f"Created: {datetime.now().astimezone().isoformat(timespec='seconds')}"))
            w(para(f"Files: {len(files)}   Directories: {len(dirs)}   Symlinks: {len(links)}"))
            w(para(f"Total file bytes: {total:,}"))
            w(para(f"Format: {FORMAT_NAME} v{FORMAT_VERSION} (data stored in the document body)"))
            w(para("Restore with word_to_directory.py. Do not edit the encoded block below."))
            # Machine block.
            w(marker("BEGIN"))
            w(marker("MAN"))
            for p in b64_paragraphs(manifest_bytes):
                w(p)
            w(marker("MANEND"))
            index_of_file = 0
            for e in entries:
                if e["type"] != "file":
                    continue
                w(marker(f"FILE.{index_of_file}"))
                data = Path(e["_src"]).read_bytes()
                check = hashlib.sha256(data).hexdigest()
                if check != e["sha256"]:
                    raise RuntimeError(f"File changed while packing: {e['path']}")
                for p in b64_paragraphs(data):
                    w(p)
                w(marker("FILEEND"))
                index_of_file += 1
            w(marker("END"))
            w(sect)
            w('</w:body></w:document>')

    print(f"Created: {output_docx}")
    print(f"Entries: {len(entries)}  (files: {len(files)}, dirs: {len(dirs)}, links: {len(links)})")
    print(f"Total file bytes: {total:,}")
    print(f"Format: {FORMAT_NAME} v{FORMAT_VERSION}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Pack a directory into a self-contained Word .docx (data stored in the body).")
    parser.add_argument("directory", help="Directory to package")
    parser.add_argument("output", nargs="?", help="Output .docx path. Default: <directory>.docx")
    args = parser.parse_args()

    source = Path(args.directory).expanduser()
    output = Path(args.output).expanduser() if args.output else source.resolve().parent / f"{source.resolve().name}.docx"
    try:
        create_docx(source, output)
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
