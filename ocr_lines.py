"""Read a photo of a table as text lines with positions, for tables whose cell borders the cell reader misses
(a screenshot cut off at the right edge, an Excel selection box drawn over a cell).

Usage: python ocr_lines.py image2xlsx.py photo.jpg
Prints JSON: {"width": w, "lines": [[{"text", "x0", "x1"}, ...], ...]} top to bottom, left to right.
Uses the PDF bot's image2xlsx.py (named by the first argument) for levelling and offline OCR, so the photo
is read exactly as it is for the table workbook.

Small single figures (a blue "4" in a Qty column) are often skipped by the page pass. On every line with a DN
size, each Qty / PN heading column that has nothing under it is read again from a close crop, the way
image2xlsx re-reads blank cells.
"""
import json
import os
import re
import sys

HEADING = re.compile(r"^(?:qty|quantity|pn|qty\.?\s*\(?(?:nos|pcs)\)?)$", re.I)


def reread(ocr, img, box):
    """Digits inside box (x0, y0, x1, y1) of the page, or None."""
    x0, y0, x1, y1 = (int(round(v)) for v in box)
    h, w = img.shape[:2]
    x0, y0, x1, y1 = max(0, x0), max(0, y0), min(w, x1), min(h, y1)
    if x1 - x0 < 10 or y1 - y0 < 10:
        return None
    piece = ocr.ink_piece(img, (x0, y0, x1, y1))
    if piece is not None:
        (text, score), = ocr.engine().text_recognizer([piece])[0]
        text = "".join(str(text).translate(ocr.FULLWIDTH).split())
        if re.fullmatch(r"\d[\d,]*", text) and float(score) >= ocr.BLANK_CELL_SCORE:
            return text
    found = [i["text"] for i in ocr.read_cell(img, (x0, y0, x1, y1))
             if re.fullmatch(r"\d[\d,]*", i["text"].replace(" ", "")) and float(i["score"]) >= ocr.BLANK_CELL_SCORE]
    return found[0].replace(" ", "") if len(found) == 1 else None


def main(argv):
    if len(argv) != 3:
        print(__doc__)
        return 2
    sys.path.insert(0, os.path.dirname(os.path.abspath(argv[1])))
    import image2xlsx as ocr
    img = ocr.flatten(ocr.deskew(ocr.load(argv[2])))
    lines = ocr.join_lines(ocr.read_text(img))
    width = img.shape[1]
    header = next((n for n, line in enumerate(lines) if any(HEADING.match(i["text"].strip()) for i in line)), None)
    if header is not None:
        columns = [((i["x0"] + i["x1"]) / 2, max(i["x1"] - i["x0"], 0.06 * width) / 2)
                   for i in lines[header] if HEADING.match(i["text"].strip())]
        for line in lines[header + 1:]:
            if not any(re.search(r"DN\s?\d", i["text"], re.I) for i in line):
                continue
            y0, y1 = min(i["y0"] for i in line), max(i["y1"] for i in line)
            pad = 0.4 * (y1 - y0)
            for centre, half in columns:
                if any(abs((i["x0"] + i["x1"]) / 2 - centre) <= half for i in line):
                    continue
                text = reread(ocr, img, (centre - half, y0 - pad, centre + half, y1 + pad))
                if text:
                    line.append({"x0": centre - 1, "x1": centre + 1, "y0": y0, "y1": y1, "text": text, "reread": True})
            line.sort(key=lambda i: i["x0"])
    out = [[{"text": i["text"], "x0": round(i["x0"]), "x1": round(i["x1"]), **({"reread": True} if i.get("reread") else {})}
            for i in line] for line in lines]
    print(json.dumps({"width": int(width), "lines": out}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
