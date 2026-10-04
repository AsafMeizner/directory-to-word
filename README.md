# Directory to Word

`directory_to_word.py` converts an entire directory into one self-contained `.docx` file using only the Python standard library.

## What it includes

- All regular files, including hidden files
- Empty directories
- Binary files
- File timestamps and permission bits where available
- Safe symbolic-link metadata
- A readable Word view of text files using their relative paths as headings
- An embedded machine-readable payload containing exact file bytes
- SHA-256 hashes for integrity

Binary files are stored inside the Word package even though their raw bytes are not printed into the visible document body.

## Requirements

Python 3.10 or newer.

No third-party packages are required.

## Basic usage

```bash
python directory_to_word.py /path/to/project
```

This creates:

```text
/path/to/project.docx
```

You can choose the output filename explicitly:

```bash
python directory_to_word.py /path/to/project /path/to/archive.docx
```

## Windows example

```powershell
py directory_to_word.py "C:\Users\me\project" "C:\Users\me\project_archive.docx"
```

## Ubuntu / Kubuntu / Linux example

```bash
python3 directory_to_word.py ~/project ~/project_archive.docx
```

## Output structure

The Word document starts with archive metadata and then shows the directory contents in relative-path order.

Text files are displayed with their indentation preserved. Binary files are represented by a short note in the visible document and are retained in the embedded payload.

## Notes

- If the output `.docx` is created inside the source directory, that output file is automatically excluded from the archive.
- Symbolic links are recorded as links and are not followed while scanning the directory.
- File content is stored exactly as bytes in the embedded payload.
- The document can be opened normally in Microsoft Word or LibreOffice.
