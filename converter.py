#!/usr/bin/env python3
"""
convert_legacy_word.py

Batch-converts legacy Word 6.0/95 files (with or without .doc extension)
to modern .docx using LibreOffice. Works on macOS and Linux.

Usage:
    python3 convert_legacy_word.py /path/to/folder
    python3 convert_legacy_word.py /path/to/folder --output /path/to/output
    python3 convert_legacy_word.py /path/to/folder --no-extension-only
"""

import argparse
import subprocess
import sys
import shutil
from pathlib import Path


# LibreOffice binary locations to try, in order
SOFFICE_CANDIDATES = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
    "soffice",                                                 # Linux (on PATH)
    "/usr/bin/soffice",                                        # Linux explicit
]


def find_soffice() -> str | None:
    """Return the path to soffice, or None if not found."""
    for candidate in SOFFICE_CANDIDATES:
        if shutil.which(candidate) or Path(candidate).exists():
            return candidate
    return None


def is_legacy_word_file(path: Path) -> bool:
    """
    Check the file magic bytes to confirm it's an OLE2 compound document
    (the container format used by Word 6/95/97-2003 .doc files).
    Magic bytes: D0 CF 11 E0 A1 B1 1A E1
    """
    try:
        with open(path, "rb") as f:
            header = f.read(8)
        return header == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    except OSError:
        return False


def collect_files(input_dir: Path, no_extension_only: bool) -> list[Path]:
    """
    Gather candidate files from input_dir (non-recursive).
    If no_extension_only is True, only files with no extension are included.
    Otherwise, files with no extension OR a .doc extension are included.
    Files are verified to be OLE2 compound documents before being included.
    """
    candidates = []
    for f in sorted(input_dir.iterdir()):
        if not f.is_file():
            continue
        ext = f.suffix.lower()
        if no_extension_only:
            if ext != "":
                continue
        else:
            if ext not in ("", ".doc"):
                continue
        if is_legacy_word_file(f):
            candidates.append(f)
        else:
            print(f"  [skip] {f.name} — not an OLE2 Word file")
    return candidates


def convert_file(soffice: str, src: Path, output_dir: Path) -> bool:
    """
    Convert a single file to .docx using LibreOffice headless.
    LibreOffice writes the output file named <stem>.docx into output_dir.
    Returns True on success.
    """
    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to", "docx",
            "--outdir", str(output_dir),
            str(src),
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"  [error] {src.name}")
        print(f"          {result.stderr.strip()}")
        return False
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Batch-convert legacy Word 6/95 files to .docx via LibreOffice."
    )
    parser.add_argument(
        "input_dir",
        type=Path,
        help="Folder containing the legacy files.",
    )
    parser.add_argument(
        "--output", "-o",
        type=Path,
        default=None,
        help="Output folder for .docx files (default: 'converted' subfolder inside input_dir).",
    )
    parser.add_argument(
        "--no-extension-only",
        action="store_true",
        help="Only process files with no extension (skip .doc files).",
    )
    args = parser.parse_args()

    # Validate input directory
    input_dir: Path = args.input_dir.resolve()
    if not input_dir.is_dir():
        print(f"Error: '{input_dir}' is not a directory.")
        sys.exit(1)

    # Find LibreOffice
    soffice = find_soffice()
    if not soffice:
        print("Error: LibreOffice not found.")
        print("Install it from https://www.libreoffice.org/download/libreoffice/")
        print("On macOS, the expected path is:")
        print("  /Applications/LibreOffice.app/Contents/MacOS/soffice")
        sys.exit(1)
    print(f"Using LibreOffice: {soffice}")

    # Set up output directory
    output_dir: Path = args.output.resolve() if args.output else input_dir / "converted"
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output folder:     {output_dir}\n")

    # Collect files
    files = collect_files(input_dir, args.no_extension_only)
    if not files:
        print("No legacy Word files found in the input directory.")
        sys.exit(0)
    print(f"Found {len(files)} file(s) to convert:\n")

    # Convert
    success_count = 0
    fail_count = 0
    for f in files:
        print(f"  Converting: {f.name} ...", end=" ", flush=True)
        ok = convert_file(soffice, f, output_dir)
        if ok:
            out_name = output_dir / (f.stem + ".docx")
            print(f"OK → {out_name.name}")
            success_count += 1
        else:
            fail_count += 1

    # Summary
    print(f"\nDone. {success_count} converted, {fail_count} failed.")
    if fail_count:
        print("Failed files may be corrupt or a non-Word OLE2 format.")


if __name__ == "__main__":
    main()
