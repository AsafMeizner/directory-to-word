#!/usr/bin/env python3
"""Create a self-contained Word .docx archive from a directory using only Python stdlib."""

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
FORMAT_VERSION = 1
PAYLOAD_PART = "customXml/item1.xml"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def xml_safe(text: str) -> str:
    """Remove characters XML 1.0 cannot represent."""
    return "".join(
        ch
        for ch in text
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
                entries.append(
                    {
                        "path": rel,
                        "type": "symlink",
                        "target": target,
                        "mode": file_mode(path, follow_symlinks=False),
                        "mtime_ns": file_mtime_ns(path, follow_symlinks=False),
                    }
                )
                continue

            if item.is_dir(follow_symlinks=False):
                entries.append(
                    {
                        "path": rel,
                        "type": "directory",
                        "mode": file_mode(path),
                        "mtime_ns": file_mtime_ns(path),
                    }
                )
                walk(path)
                continue

            if item.is_file(follow_symlinks=False):
                data = path.read_bytes()
                entries.append(
                    {
                        "path": rel,
                        "type": "file",
                        "size": len(data),
                        "sha256": hashlib.sha256(data).hexdigest(),
                        "mode": file_mode(path),
                        "mtime_ns": file_mtime_ns(path),
                        "data_b64": base64.b64encode(data).decode("ascii"),
                    }
                )

    walk(root)
    return entries


def try_text(data: bytes) -> tuple[bool, str]:
    if b"\x00" in data:
        return False, ""
    try:
        return True, data.decode("utf-8")
    except UnicodeDecodeError:
        return False, ""


def paragraph(text: str = "", style: str | None = None, code: bool = False) -> str:
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    if text == "":
        return f"<w:p>{ppr}</w:p>"
    preserve = ' xml:space="preserve"' if code or text.startswith(" ") or text.endswith(" ") else ""
    run_props = '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/></w:rPr>' if code else ""
    return f'<w:p>{ppr}<w:r>{run_props}<w:t{preserve}>{xml_text(text)}</w:t></w:r></w:p>'


def page_break() -> str:
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def build_document_xml(root: Path, entries: list[dict], payload_sha256: str) -> str:
    files = [e for e in entries if e["type"] == "file"]
    directories = [e for e in entries if e["type"] == "directory"]
    links = [e for e in entries if e["type"] == "symlink"]
    total_bytes = sum(e.get("size", 0) for e in files)

    body: list[str] = []
    body.append(paragraph(f"Directory Archive: {root.name}", "Title"))
    body.append(paragraph(f"Created: {datetime.now().astimezone().isoformat(timespec='seconds')}"))
    body.append(paragraph(f"Files: {len(files)}   Directories: {len(directories)}   Symlinks: {len(links)}"))
    body.append(paragraph(f"Total file bytes: {total_bytes:,}"))
    body.append(paragraph(f"Archive format: {FORMAT_NAME} v{FORMAT_VERSION}"))
    body.append(paragraph(f"Payload SHA-256: {payload_sha256}"))
    body.append(paragraph("Text files are shown below. Binary files are preserved inside the document package."))
    body.append(page_break())

    for index, entry in enumerate(entries):
        rel = entry["path"]
        kind = entry["type"]
        body.append(paragraph(rel, "Heading1"))

        if kind == "directory":
            body.append(paragraph("Directory"))
        elif kind == "symlink":
            body.append(paragraph(f"Symbolic link -> {entry['target']}"))
        else:
            body.append(
                paragraph(
                    f"File | {entry['size']:,} bytes | SHA-256 {entry['sha256']}"
                )
            )
            data = base64.b64decode(entry["data_b64"])
            is_text, text = try_text(data)
            if is_text:
                if text == "":
                    body.append(paragraph("[empty text file]", "Code", code=True))
                else:
                    # splitlines keeps the visible source manageable and preserves indentation.
                    for line in text.splitlines():
                        body.append(paragraph(line, "Code", code=True))
                    if text.endswith(("\n", "\r")):
                        body.append(paragraph("", "Code", code=True))
            else:
                body.append(paragraph("[binary content stored in the embedded archive payload]"))

        if index != len(entries) - 1:
            body.append(paragraph(""))

    sect = (
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
        'w:header="360" w:footer="360" w:gutter="0"/>'
        '</w:sectPr>'
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}">'
        '<w:body>' + "".join(body) + sect + '</w:body></w:document>'
    )


def build_styles_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="240"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="100"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Code">
    <w:name w:val="Code"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="0" w:after="0" w:line="220" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/></w:rPr>
  </w:style>
</w:styles>'''


def build_payload_xml(payload: bytes, payload_sha256: str) -> str:
    encoded = base64.b64encode(payload).decode("ascii")
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<projectPayload format="{FORMAT_NAME}" version="{FORMAT_VERSION}" '
        f'compression="zlib" sha256="{payload_sha256}">'
        f'{encoded}</projectPayload>'
    )


def create_docx(source_dir: Path, output_docx: Path) -> None:
    source_dir = source_dir.resolve()
    output_docx = output_docx.resolve(strict=False)

    if not source_dir.is_dir():
        raise ValueError(f"Not a directory: {source_dir}")
    if output_docx.suffix.lower() != ".docx":
        raise ValueError("Output filename must end in .docx")

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    entries = collect_entries(source_dir, output_docx)

    manifest = {
        "format": FORMAT_NAME,
        "version": FORMAT_VERSION,
        "root_name": source_dir.name,
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "entries": entries,
    }
    raw_json = json.dumps(manifest, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    payload = zlib.compress(raw_json, level=9)
    payload_sha256 = hashlib.sha256(payload).hexdigest()

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/customXml/item1.xml" ContentType="application/xml"/>
</Types>'''

    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
</Relationships>'''

    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Directory Archive - {xml_text(source_dir.name)}</dc:title>
  <dc:creator>directory_to_word.py</dc:creator>
  <cp:lastModifiedBy>directory_to_word.py</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''

    app = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>directory_to_word.py</Application>
</Properties>'''

    document_xml = build_document_xml(source_dir, entries, payload_sha256)
    payload_xml = build_payload_xml(payload, payload_sha256)

    with ZipFile(output_docx, "w", compression=ZIP_DEFLATED, compresslevel=9) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/styles.xml", build_styles_xml())
        docx.writestr("word/_rels/document.xml.rels", doc_rels)
        docx.writestr("docProps/core.xml", core)
        docx.writestr("docProps/app.xml", app)
        docx.writestr(PAYLOAD_PART, payload_xml)

    print(f"Created: {output_docx}")
    print(f"Entries: {len(entries)}")
    print(f"Embedded payload SHA-256: {payload_sha256}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Turn a directory into a structured, self-contained Word .docx file."
    )
    parser.add_argument("directory", help="Directory to package")
    parser.add_argument(
        "output",
        nargs="?",
        help="Output .docx path. Default: <directory-name>.docx beside the directory",
    )
    args = parser.parse_args()

    source = Path(args.directory).expanduser()
    if args.output:
        output = Path(args.output).expanduser()
    else:
        output = source.resolve().parent / f"{source.resolve().name}.docx"

    try:
        create_docx(source, output)
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
