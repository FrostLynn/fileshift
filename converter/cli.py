import sys
import argparse
import os
from converter.engine import convert_file, SUPPORTED_CONVERSIONS

def main():
    parser = argparse.ArgumentParser(
        prog="fileshift",
        description="FileShift: Local-First File Converter (XLSX <-> CSV, PDF <-> DOCX)"
    )
    parser.add_argument("input", help="Path to input file")
    parser.add_argument("output", nargs="?", help="Path to output file (optional)")
    parser.add_argument("--from", dest="from_fmt", help="Explicit source format (e.g. xlsx, csv, pdf, docx)")
    parser.add_argument("--to", dest="to_fmt", help="Explicit target format (e.g. csv, xlsx, docx, pdf)")
    parser.add_argument("--delimiter", help="CSV delimiter (e.g. ',' or ';')", default=None)
    parser.add_argument("--sheet", help="Sheet index for XLSX (default: 0)", type=int, default=0)

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: File input '{args.input}' tidak ditemukan.", file=sys.stderr)
        sys.exit(1)

    with open(args.input, "rb") as f:
        file_bytes = f.read()

    from_fmt = args.from_fmt or (args.input.rsplit(".", 1)[-1].lower() if "." in args.input else "")
    to_fmt = args.to_fmt

    if not to_fmt and args.output:
        to_fmt = args.output.rsplit(".", 1)[-1].lower() if "." in args.output else ""

    if not to_fmt:
        # Default target mapping
        defaults = {
            "xlsx": "csv",
            "csv": "xlsx",
            "pdf": "docx",
            "docx": "pdf"
        }
        to_fmt = defaults.get(from_fmt)

    if not to_fmt:
        print(f"Error: Format tujuan tidak dapat ditentukan untuk '{from_fmt}'. Gunakan opsi --to.", file=sys.stderr)
        sys.exit(1)

    options = {
        "delimiter": args.delimiter,
        "sheet_index": args.sheet
    }

    try:
        out_bytes, out_name, _ = convert_file(
            file_bytes=file_bytes,
            from_fmt=from_fmt,
            to_fmt=to_fmt,
            original_filename=args.input,
            options=options
        )

        out_path = args.output or out_name
        with open(out_path, "wb") as f:
            f.write(out_bytes)

        print(f"Sukses mengonversi '{args.input}' -> '{out_path}' ({len(out_bytes):,} bytes)")
    except Exception as err:
        print(f"Gagal melakukan konversi: {err}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
