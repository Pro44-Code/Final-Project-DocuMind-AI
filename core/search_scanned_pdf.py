import argparse
from pathlib import Path

import pytesseract
from pdf2image import convert_from_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Search keyword in scanned/image PDF using OCR."
    )
    parser.add_argument("--pdf", required=True, help="Path to the PDF file")
    parser.add_argument("--word", required=True, help="Keyword to search")
    parser.add_argument(
        "--lang",
        default="eng",
        help='Tesseract languages (example: "eng" or "eng+kat")',
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="PDF render DPI for OCR quality (default: 300)",
    )
    parser.add_argument(
        "--tesseract-path",
        default="",
        help="Optional full path to tesseract.exe",
    )
    parser.add_argument(
        "--poppler-path",
        default="",
        help="Optional path to Poppler bin folder (where pdftoppm.exe exists)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    pdf_path = Path(args.pdf).expanduser().resolve()

    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    if args.tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = args.tesseract_path

    poppler_path = args.poppler_path or None
    pages = convert_from_path(str(pdf_path), dpi=args.dpi, poppler_path=poppler_path)

    keyword = args.word.strip().lower()
    found_pages = []

    for idx, image in enumerate(pages, start=1):
        text = pytesseract.image_to_string(image, lang=args.lang)
        if keyword in text.lower():
            found_pages.append(idx)

    if found_pages:
        print(f'Found "{args.word}" on pages: {found_pages}')
    else:
        print(f'"{args.word}" was not found.')


if __name__ == "__main__":
    main()
