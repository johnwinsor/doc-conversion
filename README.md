# doc-conversion

Batch-converts legacy Microsoft Word 6.0/95 files to modern `.docx` format using LibreOffice headless, with optional plain-text extraction for full-text search. Written to recover archival documents from the 1990s that modern Word refuses to open.

## Background

Word for Mac 6.0/95 files use the OLE2 binary compound document format. Microsoft Word 2016 and later block these files entirely with the error:

> *"[filename] uses a file type that is blocked from opening in this version."*

Simply adding a `.doc` extension does not resolve this. No pure Python library can convert this format — the spec runs to hundreds of pages and the format is effectively obsolete. LibreOffice is the standard solution, even in commercial document processing pipelines.

## Requirements

- Python 3.10+
- [uv](https://docs.astral.sh/uv/) for environment and dependency management
- [LibreOffice](https://www.libreoffice.org/download/libreoffice/) (free)
  - macOS: install to `/Applications/` as normal
  - Linux: `sudo apt install libreoffice` or equivalent
- `python-docx` — installed automatically via `uv sync` (required for `--extract-text`)

## Setup

```bash
git clone <repo-url>
cd doc-conversion
uv sync
```

## Usage

```bash
uv run converter.py <input_dir> [--output <output_dir>] [--no-extension-only] [--extract-text]
```

### Arguments

| Argument | Description |
|---|---|
| `input_dir` | Folder containing the legacy files (required) |
| `--output`, `-o` | Output folder for `.docx` files. Defaults to a `converted/` subfolder inside `input_dir` |
| `--no-extension-only` | Only process files with no extension; skip any files already named `.doc` |
| `--extract-text` | After converting each file, extract plain text to a `.txt` sidecar in the output folder. Also backfills sidecars for any `.docx` files in the output folder from previous runs |

### File handling

By default the script processes both extension-less legacy files and files with a `.doc` extension. It verifies each candidate by checking its OLE2 magic bytes (`D0 CF 11 E0 A1 B1 1A E1`) before attempting conversion, so stray files of other types in the folder are safely skipped. Original files are never modified or deleted.

### Examples

Convert a folder, writing output to a `converted/` subfolder:
```bash
uv run converter.py ~/Documents/old-files
```

Convert and produce plain-text sidecars for full-text search:
```bash
uv run converter.py ~/Documents/old-files --extract-text
```

Send output to a specific location:
```bash
uv run converter.py ~/Documents/old-files --output ~/Desktop/converted --extract-text
```

Process only extension-less files (skip `.doc` files):
```bash
uv run converter.py ~/Documents/old-files --no-extension-only
```

## Shell helper for repeated conversions (`convert_donors`)

When running conversions across many sibling subdirectories, add the following to `~/.zshrc` to avoid retyping the full base path each time.

```zsh
BASE_DONORS="/Users/johnwinsor/Library/CloudStorage/OneDrive-NortheasternUniversity/Library-Oakland - Documents/Special Collections/Restricted_Spec_Coll/Gifts & Donors (Janice)"

convert_donors() {
  uv run /Users/johnwinsor/projects/doc-conversion/converter.py "$BASE_DONORS/$1" --output "$BASE_DONORS/$1" --extract-text
}

_convert_donors() {
  local subs=("$BASE_DONORS"/*(/:t))
  compadd "$@" -- "${subs[@]}"
}

compdef _convert_donors convert_donors
```

Reload your shell after editing:
```bash
source ~/.zshrc
```

### How it works

`BASE_DONORS` holds the long base path so it never needs to be typed directly. `convert_donors` accepts a single subfolder name and constructs the full input and output paths automatically. `_convert_donors` is a zsh completion function that expands the subdirectories of `BASE_DONORS` at tab-completion time — the `*(/:t)` glob matches all subdirectories (`/`) and returns just their short names rather than full paths (`:t`). `compdef` binds the completion function to the command.

### Usage

```zsh
convert_donors '00-'01
convert_donors <Tab>    # lists available subdirectories
```

## Searching the text archive

Once `.txt` sidecars have been produced across all subdirectories, use [ripgrep](https://github.com/BurntSushi/ripgrep) to search the full archive. Install with `brew install ripgrep`.

```bash
# Simple keyword search
rg "Griffiths" "$BASE_DONORS" --glob "*.txt"

# Case-insensitive regex
rg -i "gift of .+ book" "$BASE_DONORS" --glob "*.txt"

# Show surrounding context (2 lines before and after each match)
rg -i "Stevenson" "$BASE_DONORS" --glob "*.txt" -C 2

# List only filenames that contain a match
rg -l "reunion" "$BASE_DONORS" --glob "*.txt"
```

## Output

```
Using LibreOffice: /Applications/LibreOffice.app/Contents/MacOS/soffice
Output folder:     /path/to/output

Found 3 file(s) to convert:

  Converting: Griffiths ... OK -> Griffiths.docx
  Extracting: Griffiths.docx ... OK -> Griffiths.txt
  Converting: Smith ... OK -> Smith.docx
  Extracting: Smith.docx ... OK -> Smith.txt
  Converting: Jones.doc ... OK -> Jones.docx
  Extracting: Jones.docx ... OK -> Jones.txt

Done. 3 converted, 0 failed.
      3 text extracted, 0 failed.
```

## Notes

- Conversion is non-recursive (top-level directory only per run)
- If a file fails conversion it may be corrupt or a non-Word OLE2 format (e.g. an old Excel or PowerPoint file, which share the same container format)
- Conversion fidelity depends on LibreOffice's Word 6 support; formatting in very old files may not render perfectly, but text content is reliably preserved
