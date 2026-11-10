"""Convert a photo or screenshot of a table into an Excel workbook.

Usage: python image2xlsx.py input.jpg output.xlsx

Runs fully offline on this PC (OpenCV + RapidOCR); the image is not sent anywhere.
1. The photo is levelled (rotation from the table lines) and flattened
   (perspective from the table's outer border).
2. Table cells are found from the ruled lines and arranged into rows/columns.
3. Printed text (English/Chinese) is read with RapidOCR and placed in the cell
   under it. A text box that crosses a cell border is re-read cell by cell.
4. Text outside the table becomes title rows; page footers are dropped.
Tables without ruled lines fall back to lining up text by position.
The workbook layout is the same as pdf2xlsx.py.
"""
import re
import sys

import cv2
import numpy as np

import pdf2xlsx
from pdf2xlsx import FOOTER, NoTables, cell_value

MAX_SIDE = 3000  # larger photos are scaled down for speed
MIN_SIDE = 1000  # small screenshots are scaled up so thin lines and text survive
BLANK_CELL_SCORE = 0.2  # single characters score low even when read correctly
_engine = None
# OCR sometimes returns full-width digits/letters (３, Ａ); Excel needs the normal ones.
FULLWIDTH = str.maketrans({chr(c): chr(c - 0xFEE0) for c in
                           [*range(0xFF10, 0xFF1A), *range(0xFF21, 0xFF3B), *range(0xFF41, 0xFF5B)]})
TEXTLIKE = re.compile(r"[0-9A-Za-z一-鿿]")


def engine():
    global _engine
    if _engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _engine = RapidOCR()
        # The page is levelled first, so text is upright; the angle classifier only flips short numbers (90 read as 06).
        _engine.use_angle_cls = False
    return _engine


def load(path):
    data = np.fromfile(str(path), np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR) if data.size else None
    if img is None:
        raise NoTables("The image could not be read. Please send a JPG or PNG photo.")
    h, w = img.shape[:2]
    longest = max(h, w)
    if longest > MAX_SIDE:
        img = cv2.resize(img, None, fx=MAX_SIDE / longest, fy=MAX_SIDE / longest, interpolation=cv2.INTER_AREA)
    elif longest < MIN_SIDE:
        img = cv2.resize(img, None, fx=MIN_SIDE / longest, fy=MIN_SIDE / longest, interpolation=cv2.INTER_CUBIC)
    return img


def line_mask(img):
    """White where the table's ruled lines are."""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    block = max(15, (min(gray.shape) // 60) | 1)
    bw = cv2.adaptiveThreshold(255 - gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, block, -10)
    h, w = bw.shape
    horiz = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, w // 30), 1)))
    vert = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(20, h // 40))))
    return cv2.dilate(horiz | vert, np.ones((3, 3), np.uint8))


def deskew(img):
    """Rotate so long near-horizontal lines become level."""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape
    edges = cv2.Canny(gray, 50, 150)
    lines = cv2.HoughLinesP(edges, 1, np.pi / 720, threshold=100, minLineLength=w // 4, maxLineGap=10)
    if lines is None:
        return img
    # OpenCV 4 returns shape (N, 1, 4), OpenCV 5 returns (N, 4).
    angles = [np.degrees(np.arctan2(float(y2 - y1), float(x2 - x1))) for x1, y1, x2, y2 in lines.reshape(-1, 4)]
    angles = [a for a in angles if abs(a) < 20]
    if len(angles) < 3:
        return img
    angle = float(np.median(angles))
    if abs(angle) < 0.2:
        return img
    m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    cos, sin = abs(m[0, 0]), abs(m[0, 1])
    nw, nh = int(h * sin + w * cos), int(h * cos + w * sin)
    m[0, 2] += nw / 2 - w / 2
    m[1, 2] += nh / 2 - h / 2
    return cv2.warpAffine(img, m, (nw, nh), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)


def order_corners(pts):
    s = pts.sum(1)
    d = np.diff(pts, axis=1).ravel()  # y - x
    return np.array([pts[np.argmin(s)], pts[np.argmin(d)], pts[np.argmax(s)], pts[np.argmax(d)]], np.float32)


def flatten(img):
    """Undo perspective using the largest table's outer border; the rest of the photo moves with it."""
    mask = line_mask(img)
    h, w = mask.shape
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return img
    big = max(contours, key=cv2.contourArea)
    if cv2.contourArea(big) < 0.05 * w * h:
        return img
    approx = cv2.approxPolyDP(big, 0.02 * cv2.arcLength(big, True), True)
    if len(approx) != 4:
        return img
    src = order_corners(approx.reshape(4, 2).astype(np.float32))
    tl, tr, br, bl = src
    width = max(np.linalg.norm(tr - tl), np.linalg.norm(br - bl))
    height = max(np.linalg.norm(bl - tl), np.linalg.norm(br - tr))
    x0, y0 = float(min(tl[0], bl[0])), float(min(tl[1], tr[1]))
    dst = np.array([[x0, y0], [x0 + width, y0], [x0 + width, y0 + height], [x0, y0 + height]], np.float32)
    if np.abs(src - dst).max() < 3:
        return img
    matrix = cv2.getPerspectiveTransform(src, dst)
    size = (int(max(w, x0 + width + 10)), int(max(h, y0 + height + 10)))
    return cv2.warpPerspective(img, matrix, size, flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)


def find_cells(mask):
    """Rectangles enclosed by ruled lines, as (x0, y0, x1, y1)."""
    h, w = mask.shape
    contours, hierarchy = cv2.findContours(cv2.bitwise_not(mask), cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    if hierarchy is None:
        return []
    min_side = max(8, min(h, w) // 150)
    boxes = []
    for contour, link in zip(contours, hierarchy[0]):
        if link[3] != -1:
            continue
        x, y, bw, bh = cv2.boundingRect(contour)
        if bw < min_side or bh < min_side or bw * bh > 0.5 * w * h:
            continue
        if x <= 1 or y <= 1 or x + bw >= w - 1 or y + bh >= h - 1:
            continue  # background touching the photo edge
        if cv2.contourArea(contour) < 0.6 * bw * bh:
            continue
        boxes.append((x, y, x + bw, y + bh))
    # A box that contains another box is a region with a broken border line; keep the inner boxes.
    return [b for b in boxes if not any(o != b and b[0] <= o[0] and b[1] <= o[1] and o[2] <= b[2] and o[3] <= b[3]
                                        for o in boxes)]


def group_tables(boxes, gap=8):
    """Split cells into tables: cells that touch belong to the same table."""
    parent = list(range(len(boxes)))

    def root(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i, a in enumerate(boxes):
        for j in range(i + 1, len(boxes)):
            b = boxes[j]
            if a[0] - gap <= b[2] and b[0] - gap <= a[2] and a[1] - gap <= b[3] and b[1] - gap <= a[3]:
                parent[root(i)] = root(j)
    groups = {}
    for i, b in enumerate(boxes):
        groups.setdefault(root(i), []).append(b)
    return [g for g in groups.values() if len(g) >= 4]


def cluster(values, tol):
    values = sorted(values)
    groups = [[values[0]]]
    for v in values[1:]:
        if v - groups[-1][-1] <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def nearest(centres, v):
    return min(range(len(centres)), key=lambda i: abs(centres[i] - v))


def read_text(img):
    items = []
    result, _ = engine()(img)
    for box, text, score in result or []:
        text = " ".join(str(text).translate(FULLWIDTH).split())
        if not text:
            continue
        pts = np.array(box, np.float32)
        (x0, y0), (x1, y1) = pts.min(0), pts.max(0)
        items.append({"x0": float(x0), "y0": float(y0), "x1": float(x1), "y1": float(y1), "text": text,
                      "score": float(score)})
    return items


def join_lines(items):
    """Join text boxes into reading order: top-to-bottom lines, left-to-right within a line."""
    lines = []
    for it in sorted(items, key=lambda i: (i["y0"], i["x0"])):
        height = it["y1"] - it["y0"]
        if lines and abs((lines[-1][0]) - (it["y0"] + it["y1"]) / 2) < 0.5 * height:
            lines[-1][1].append(it)
        else:
            lines.append([(it["y0"] + it["y1"]) / 2, [it]])
    return [sorted(ws, key=lambda i: i["x0"]) for _, ws in lines]


def read_cell(img, cell):
    """OCR one cell on its own, for text that the page-level pass merged across a border."""
    x0, y0, x1, y1 = cell
    crop = img[y0 + 2:y1 - 2, x0 + 2:x1 - 2]
    if crop.size == 0:
        return []
    scale = 2.0 if crop.shape[0] < 60 else 1.0
    crop = cv2.resize(crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    crop = cv2.copyMakeBorder(crop, 12, 12, 12, 12, cv2.BORDER_REPLICATE)
    items = read_text(crop)
    for it in items:  # back to page coordinates so reading order still works
        for k, base in (("x0", x0), ("x1", x0), ("y0", y0), ("y1", y0)):
            it[k] = base + 2 + (it[k] - 12) / scale
    return items


def ink_piece(img, cell):
    """The inked part of a cell, padded with its background colour; None if the cell is blank.

    Specks and thin fragments of the cell's own border lines are ignored.
    """
    x0, y0, x1, y1 = cell
    inner = img[y0 + 4:y1 - 4, x0 + 4:x1 - 4]
    if inner.size == 0:
        return None
    gray = cv2.cvtColor(inner, cv2.COLOR_BGR2GRAY)
    _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    if (bw > 0).mean() > 0.5:
        return None  # a solid or dark cell, not text
    count, _, stats, _ = cv2.connectedComponentsWithStats(bw)
    h, w = bw.shape
    # Border fragments are thin lines; text touching the cell edge (tall cells, digits sitting low) is kept.
    keep = [s for s in stats[1:count] if s[4] >= 6 and min(s[2], s[3]) >= 3]
    if not keep:
        return None
    a, b = min(s[0] for s in keep), min(s[1] for s in keep)
    c, d = max(s[0] + s[2] for s in keep), max(s[1] + s[3] for s in keep)
    if d - b < 6:
        return None
    pad = max(4, (d - b) // 4)
    piece = inner[b:d, a:c]
    background = [int(v) for v in np.median(inner[bw == 0], axis=0)] if (bw == 0).any() else [255, 255, 255]
    return cv2.copyMakeBorder(piece, pad, pad, pad * 2, pad * 2, cv2.BORDER_CONSTANT, value=background)


def read_blank_cells(img, cells, contents, line_height):
    """Re-read cells the page pass left empty but that contain ink (single digits are often skipped)."""
    todo = []
    for cell in cells:
        if contents[cell]:
            continue
        piece = ink_piece(img, cell)
        if piece is None:
            continue
        if piece.shape[0] > 2.2 * line_height:
            contents[cell] = read_cell(img, cell)  # several lines: needs detection
        else:
            todo.append((cell, piece))
    if not todo:
        return
    results, _ = engine().text_recognizer([p for _, p in todo])
    for (cell, _), (text, score) in zip(todo, results):
        text = " ".join(str(text).translate(FULLWIDTH).split())
        if not text.replace(",", "").isdigit():
            # A small figure in a tall cell is misread from the bare crop (7 read as 7n, or nothing);
            # a re-read with a margin gives the digits.
            digits = [i for i in read_cell(img, cell) if i["text"].replace(",", "").isdigit()
                      and float(i["score"]) >= BLANK_CELL_SCORE]
            if digits:
                contents[cell] = digits
                continue
        if float(score) >= BLANK_CELL_SCORE and TEXTLIKE.search(text):
            contents[cell] = [{"x0": cell[0], "y0": cell[1], "x1": cell[2], "y1": cell[3], "text": text,
                               "score": float(score)}]


def fix_numeric_columns(img, grid, cell_of):
    """In a column of plain numbers (e.g. item No.), re-read short entries that came out with letters.

    The page pass can misread a small number ("64" as "t9"). The entry is replaced only when the
    cell-by-cell re-read gives digits; headers such as "No" or "项次" stay as they are.
    """
    columns = {}
    for (r, c), text in grid.items():
        columns.setdefault(c, []).append((r, text))
    todo = []
    for c, entries in columns.items():
        numeric = sum(t.replace(",", "").isdigit() for _, t in entries)
        if numeric < 3 or numeric < 0.7 * len(entries):
            continue
        for r, text in entries:
            if len(text) <= 4 and not text.replace(",", "").isdigit():
                piece = ink_piece(img, cell_of[(r, c)])
                if piece is not None:
                    todo.append(((r, c), piece))
    if not todo:
        return
    results, _ = engine().text_recognizer([p for _, p in todo])
    for (key, _), (text, score) in zip(todo, results):
        text = "".join(str(text).translate(FULLWIDTH).split())
        if text.replace(",", "").isdigit() and float(score) >= BLANK_CELL_SCORE:
            grid[key] = text


def table_rows(img, cells, items):
    """Place text into cells; return (rows, used_item_ids)."""
    heights = sorted(c[3] - c[1] for c in cells)
    tol = max(6.0, 0.3 * heights[len(heights) // 2])
    row_tops = cluster([c[1] for c in cells], tol)
    col_lefts = cluster([c[0] for c in cells], tol)
    bx0, by0 = min(c[0] for c in cells), min(c[1] for c in cells)
    bx1, by1 = max(c[2] for c in cells), max(c[3] for c in cells)
    contents = {c: [] for c in cells}
    reread = set()
    used = set()
    for n, it in enumerate(items):
        cx, cy = (it["x0"] + it["x1"]) / 2, (it["y0"] + it["y1"]) / 2
        if not (bx0 <= cx <= bx1 and by0 <= cy <= by1):
            continue
        used.add(n)
        width = max(1.0, it["x1"] - it["x0"])
        overlaps = [c for c in cells if c[1] <= cy <= c[3] and min(it["x1"], c[2]) - max(it["x0"], c[0]) > 0.2 * width]
        if len(overlaps) > 1:
            reread.update(overlaps)
            continue
        if not overlaps:
            overlaps = [min(cells, key=lambda c: abs((c[0] + c[2]) / 2 - cx) + abs((c[1] + c[3]) / 2 - cy))]
        contents[overlaps[0]].append(it)
    for cell in reread:
        contents[cell] = read_cell(img, cell)
    line_heights = sorted(i["y1"] - i["y0"] for i in items) or [20.0]
    read_blank_cells(img, cells, contents, line_heights[len(line_heights) // 2])
    grid = {}
    cell_of = {}
    for cell, found in contents.items():
        key = (nearest(row_tops, cell[1]), nearest(col_lefts, cell[0]))
        cell_of.setdefault(key, cell)
        text = " ".join(" ".join(i["text"] for i in line) for line in join_lines(found))
        if text:
            grid[key] = f"{grid[key]} {text}" if key in grid else text
    fix_numeric_columns(img, grid, cell_of)
    rows =[[cell_value(grid.get((r, c), "")) for c in range(len(col_lefts))] for r in range(len(row_tops))]
    rows = [r for r in rows if any(v is not None for v in r)]
    keep = [c for c in range(len(col_lefts)) if any(r[c] is not None for r in rows)]
    return [[r[c] for c in keep] for r in rows], used, by0


def borderless_rows(items, width):
    """Fallback for tables without ruled lines: line up text boxes by position."""
    lines = join_lines(items)
    if not any(len(line) >= 2 for line in lines):
        return None
    cols = cluster([i["x0"] for i in items], max(8.0, 0.02 * width))
    rows = []
    for line in lines:
        row = {}
        for it in line:
            c = nearest(cols, it["x0"])
            row[c] = f"{row[c]} {it['text']}" if c in row else it["text"]
        rows.append([cell_value(row.get(c, "")) for c in range(len(cols))])
    keep = [c for c in range(len(cols)) if any(r[c] is not None for r in rows)]
    return [[r[c] for c in keep] for r in rows]


def extract(path):
    """Return one page of ('table', rows) and ('text', str) items, like pdf2xlsx.extract."""
    img = flatten(deskew(load(path)))
    items = [i for i in read_text(img) if not FOOTER.match(i["text"].replace(" ", ""))
             and not FOOTER.match(i["text"])]
    if not items:
        raise NoTables("No text was found in the photo. Please send a clearer, well-lit photo of the table.")
    placed = []
    used = set()
    for cells in group_tables(find_cells(line_mask(img))):
        rows, taken, top = table_rows(img, cells, items)
        used |= taken
        if rows:
            placed.append((top, 0, "table", rows))
    if not placed:
        rows = borderless_rows(items, img.shape[1])
        if not rows:
            raise NoTables("No table was found in the photo.")
        return [("table", rows)]
    for line in join_lines([it for n, it in enumerate(items) if n not in used]):
        text = " ".join(i["text"] for i in line)
        placed.append((min(i["y0"] for i in line), min(i["x0"] for i in line), "text", text))
    placed.sort(key=lambda p: (p[0], p[1]))
    return [(kind, payload) for _, _, kind, payload in placed]


def convert(image_path, xlsx_path, title=None):
    page = extract(image_path)
    pdf2xlsx.write_workbook([page], xlsx_path, title=title or pdf2xlsx.sheet_title(image_path))
    tables = sum(1 for kind, _ in page if kind == "table")
    rows = sum(len(p) for kind, p in page if kind == "table")
    return {"pages": 1, "tables": tables, "rows": rows}


def main(argv):
    if len(argv) != 3:
        print(__doc__)
        return 2
    try:
        info = convert(argv[1], argv[2])
    except NoTables as e:
        print(f"NO_TABLES: {e}")
        return 3
    print(f"OK tables={info['tables']} rows={info['rows']}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
