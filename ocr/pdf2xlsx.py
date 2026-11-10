"""Convert the tables in a PDF into an Excel workbook.

Usage: python pdf2xlsx.py input.pdf output.xlsx

Table borders are found by PyMuPDF. Cell text is then re-read from the page
text by position, so that:
- a word that slightly overflows a column border stays in its own cell
  (e.g. "1320LGH*3PCS/LGH" printed wider than its column);
- two words printed without a space across a border are split at the border
  (e.g. the header "项次尺寸Size" becomes "项次" and "尺寸Size");
- page footers such as "第 1 页，共 3 页" or "Page 1 of 3" are excluded;
- an unbordered gap inside a table (a total, a section title) is written as
  its own rows with each value under the column it sits in.
Text outside the tables (titles such as "PO.20288") is written as its own row
in reading order, so nothing on the page is lost. Plain numbers become
numeric cells; "10 LGH" or "1,260PCS" stay text.
"""
import re
import sys
from pathlib import Path

import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

FOOTER = re.compile(
    r"^(?:第\s*\d+\s*页\s*[，,]?\s*共\s*\d+\s*页|page\s*\d+\s*(?:of|/)\s*\d+|\d+\s*/\s*\d+|-\s*\d+\s*-)$",
    re.I,
)
NUMBER = re.compile(r"^-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$")
SPLIT_SHARE = 0.4  # a word reaching this far into a neighbouring cell is split at the border
LINE_TOLERANCE = 3  # points; words whose tops differ less than this share a line


class NoTables(Exception):
    pass


def cell_value(text):
    """Return an int/float for a plain number, otherwise the cleaned text."""
    text = " ".join(text.split())
    if not text:
        return None
    if NUMBER.match(text):
        plain = text.replace(",", "")
        return float(plain) if "." in plain else int(plain)
    return text


class Word:
    __slots__ = ("x0", "y0", "x1", "y1", "text", "chars")

    def __init__(self, chars):
        self.chars = chars  # [(x0, x1, character)]
        self.x0 = min(c[0] for c in chars)
        self.x1 = max(c[1] for c in chars)
        self.y0, self.y1 = chars[0][3], chars[0][4]
        self.text = "".join(c[2] for c in chars)

    @property
    def cx(self):
        return (self.x0 + self.x1) / 2

    @property
    def cy(self):
        return (self.y0 + self.y1) / 2


def page_text(page):
    """Return (words, lines): words with character boxes, and text lines, minus page-number footers."""
    words, lines = [], []
    for block in page.get_text("rawdict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            chars = [c for span in line["spans"] for c in span["chars"]]
            text = " ".join("".join(c["c"] for c in chars).split())
            if not text or FOOTER.match(text):
                continue
            lines.append((fitz.Rect(line["bbox"]), text))
            current = []
            for c in chars:
                if c["c"].isspace():
                    if current:
                        words.append(Word(current))
                    current = []
                else:
                    x0, y0, x1, y1 = c["bbox"]
                    current.append((x0, x1, c["c"], y0, y1))
            if current:
                words.append(Word(current))
    return words, lines


def contains(rect, x, y):
    return rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]


def column_ranges(table):
    """Horizontal ranges of the table's columns, taken from the row with the most cells."""
    best = max(table.rows, key=lambda r: sum(c is not None for c in r.cells))
    return [(c[0], c[2]) for c in best.cells if c is not None]


def column_of(columns, x):
    for i, (x0, x1) in enumerate(columns):
        if x0 <= x <= x1:
            return i
    return min(range(len(columns)), key=lambda i: min(abs(columns[i][0] - x), abs(columns[i][1] - x)))


def group_lines(words):
    """Group words into text lines by their top edge, in reading order."""
    lines = []
    for w in sorted(words, key=lambda w: (w.y0, w.x0)):
        if lines and abs(lines[-1][0] - w.y0) < LINE_TOLERANCE:
            lines[-1][1].append(w)
        else:
            lines.append((w.y0, [w]))
    return [sorted(ws, key=lambda w: w.x0) for _, ws in lines]


def group_phrases(line):
    """Split one line of words into phrases: words separated by no more than a normal space."""
    phrases = []
    for w in line:
        if phrases and (w.x0 - phrases[-1][-1].x1) < 0.7 * (w.y1 - w.y0):
            phrases[-1].append(w)
        else:
            phrases.append([w])
    return phrases


def place_word(word, cells, contents):
    """Put a word into the cell under its centre, splitting it at a border it clearly crosses."""
    index = next((i for i, c in enumerate(cells) if c and contains(c, word.cx, word.cy)), None)
    if index is None:
        return
    cell = cells[index]
    split = False
    if word.x0 < cell[0]:
        left = next((c for c in reversed(cells[:index]) if c), None)
        if left and (cell[0] - word.x0) > SPLIT_SHARE * (left[2] - left[0]):
            split = True
    if word.x1 > cell[2]:
        right = next((c for c in cells[index + 1:] if c), None)
        if right and (word.x1 - cell[2]) > SPLIT_SHARE * (right[2] - right[0]):
            split = True
    if not split:
        contents.setdefault(index, []).append(word)
        return
    piece, piece_index = [], None
    for ch in word.chars:
        cx = (ch[0] + ch[1]) / 2
        i = next((k for k, c in enumerate(cells) if c and c[0] <= cx <= c[2]), index)
        if piece and i != piece_index:
            contents.setdefault(piece_index, []).append(Word(piece))
            piece = []
        piece.append(ch)
        piece_index = i
    if piece:
        contents.setdefault(piece_index, []).append(Word(piece))


def cell_text(words):
    return " ".join(w.text for line in group_lines(words) for w in line)


def table_rows(table, words):
    columns = column_ranges(table)
    rows = []
    for row in table.rows:
        cells = list(row.cells)
        present = [c for c in cells if c]
        if not present:
            continue
        inside = [w for w in words if contains(row.bbox, w.cx, w.cy)]
        spanning = len(present) == 1 and len(columns) > 1 and (present[0][2] - present[0][0]) > 1.5 * max(
            x1 - x0 for x0, x1 in columns)
        if spanning:
            # An unbordered gap: one output row per text line; each phrase goes under the column it starts in.
            for line in group_lines(inside):
                values = {}
                for phrase in group_phrases(line):
                    values.setdefault(column_of(columns, phrase[0].x0), []).extend(phrase)
                rows.append([cell_value(cell_text(values[i])) if i in values else None for i in range(len(columns))])
            continue
        contents = {}
        for w in inside:
            place_word(w, cells, contents)
        rows.append([cell_value(cell_text(contents[i])) if i in contents else None for i in range(len(cells))])
    return [r for r in rows if any(v is not None for v in r)]


def page_items(page):
    """Tables and free text lines of one page, in reading order."""
    words, lines = page_text(page)
    items = []
    covered = []
    for t in page.find_tables().tables:
        rows = table_rows(t, words)
        if rows:
            items.append((t.bbox[1], t.bbox[0], "table", rows))
            covered.append(fitz.Rect(t.bbox))
    for rect, text in lines:
        if any(rect.intersects(c) for c in covered):
            continue
        items.append((rect.y0, rect.x0, "text", text))
    items.sort(key=lambda i: (round(i[0]), i[1]))
    return [(kind, payload) for _, _, kind, payload in items]


def extract(pdf_path):
    """Return one list per page; each holds ('table', rows) and ('text', str) items."""
    doc = fitz.open(pdf_path)
    try:
        if doc.needs_pass:
            raise NoTables("The PDF is password protected.")
        pages = [page_items(p) for p in doc]
    finally:
        doc.close()
    if not any(kind == "table" for page in pages for kind, _ in page):
        if not any(kind == "text" for page in pages for kind, _ in page):
            raise NoTables("No text found in the PDF. Scanned PDFs (images) are not supported.")
        raise NoTables("No table with borders was found in the PDF.")
    return pages


def is_header(row):
    values = [v for v in row if v is not None]
    return len(values) >= 2 and all(isinstance(v, str) and not v[0].isdigit() for v in values)


def trimmed(row):
    row = list(row)
    while row and row[-1] is None:
        row.pop()
    return row


def write_workbook(pages, xlsx_path, title="Tables"):
    wb = Workbook()
    ws = wb.active
    ws.title = title
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold = Font(bold=True)
    widths = {}
    last_header = None
    frozen = False
    row_no = 1
    last_kind = None
    for page in pages:
        first_table_on_page = True
        for kind, payload in page:
            if kind == "text":
                ws.cell(row=row_no, column=1, value=payload).font = bold
                row_no += 1
                last_kind = "text"
                continue
            if last_kind == "table":
                row_no += 1  # gap between two separate tables on the same page
            for i, row in enumerate(payload):
                header = is_header(row)
                if header and i == 0 and first_table_on_page and trimmed(row) == last_header:
                    continue  # header repeated at the top of a continued page
                for col, value in enumerate(row, start=1):
                    cell = ws.cell(row=row_no, column=col, value=value)
                    cell.border = border
                    centred = header or isinstance(value, (int, float))
                    cell.alignment = Alignment(vertical="center", horizontal="center" if centred else "left")
                    if header:
                        cell.font = bold
                    if value is not None:
                        widths[col] = max(widths.get(col, 0), min(len(str(value)), 60))
                if header:
                    last_header = trimmed(row)
                    if not frozen:
                        ws.freeze_panes = ws.cell(row=row_no + 1, column=1)
                        frozen = True
                row_no += 1
            first_table_on_page = False
            last_kind = "table"
        last_kind = None  # no gap across a page break: the table may continue
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(8, width + 2)
    wb.save(xlsx_path)


def sheet_title(pdf_path):
    return re.sub(r"[\\/*?:\[\]]", " ", Path(pdf_path).stem).strip()[:31] or "Tables"


def convert(pdf_path, xlsx_path, title=None):
    pages = extract(pdf_path)
    write_workbook(pages, xlsx_path, title=title or sheet_title(pdf_path))
    tables = sum(1 for page in pages for kind, _ in page if kind == "table")
    rows = sum(len(p) for page in pages for kind, p in page if kind == "table")
    return {"pages": len(pages), "tables": tables, "rows": rows}


def main(argv):
    if len(argv) != 3:
        print(__doc__)
        return 2
    try:
        info = convert(argv[1], argv[2])
    except NoTables as e:
        print(f"NO_TABLES: {e}")
        return 3
    print(f"OK pages={info['pages']} tables={info['tables']} rows={info['rows']}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
