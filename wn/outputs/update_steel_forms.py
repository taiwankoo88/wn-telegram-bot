import argparse
import json
import re
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"

FORM_SETS = {
    "C3D1": {
        "template_dir": Path(__file__).resolve().parent / "C3D1 Forms",
        "templates": {
            "written_notification": {
                "filename": "SS_Steel Work Written Notification C3D1_Batch_2.docx",
                "output_stem": "SS_Steel Work Written Notification C3D1",
            },
            "request_form": {
                "filename": "SS_Request Form for Issuing Wirtten Notification_Batch_2 Amended.docx",
                "output_stem": "SS_Request Form for Issuing Wirtten Notification",
            },
        },
    },
    "C3D2": {
        "template_dir": Path(__file__).resolve().parent / "C3D2 Forms",
        "templates": {
            "written_notification": {
                "filename": "SS_Steel Work Written Notification C3D2.docx",
                "output_stem": "SS_Steel Work Written Notification C3D2",
            },
            "request_form": {
                "filename": "SS_Request Form.docx",
                "output_stem": "SS_Request Form C3D2",
            },
        },
    },
}


def parse_date(value):
    value = str(value).strip()
    formats = [
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d %B %Y",
        "%d %b %Y",
        "%d-%B-%Y",
        "%d-%b-%Y",
    ]
    cleaned = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", value, flags=re.I)
    for fmt in formats:
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            pass
    return None


def ordinal_day(day):
    if 10 <= day % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix}"


def format_dates(raw_date):
    parsed = parse_date(raw_date)
    if not parsed:
        return {
            "letter": str(raw_date).strip(),
            "request": str(raw_date).strip(),
            "long": str(raw_date).strip(),
            "short": str(raw_date).strip(),
        }
    month = parsed.strftime("%B")
    return {
        "letter": f"{parsed.day}-{month}-{parsed.year}",
        "request": f"{parsed.day} {month} {parsed.year}",
        "long": f"{ordinal_day(parsed.day)} {month} {parsed.year}",
        "short": f"{parsed.day}/{parsed.month}/{str(parsed.year)[-2:]}",
    }


def description_lines(value):
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value).replace(";", "\n")
    return [line.strip() for line in text.splitlines() if line.strip()]


def para_text(paragraph):
    return "".join(t.text or "" for t in paragraph.xpath("./w:r//w:t", namespaces=NS))


def cell_text(cell):
    return " ".join((t.text or "").strip() for t in cell.xpath(".//w:t", namespaces=NS) if (t.text or "").strip())


def compact(value):
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def set_paragraph_text(paragraph, value):
    texts = paragraph.xpath("./w:r//w:t", namespaces=NS)
    if not texts:
        return False
    texts[0].text = value
    for text in texts[1:]:
        text.text = ""
    return True


def replace_span(text_nodes, start, end, replacement):
    position = 0
    inserted = False
    changed = False
    for node in text_nodes:
        original = node.text or ""
        node_start = position
        node_end = position + len(original)
        position = node_end

        if node_end <= start or node_start >= end:
            continue

        prefix = original[: max(0, start - node_start)] if start > node_start else ""
        suffix = original[max(0, end - node_start) :] if end < node_end else ""
        if not inserted:
            node.text = prefix + replacement + suffix
            inserted = True
        else:
            node.text = suffix
        changed = True
    return changed


def replace_regex_group(text_nodes, pattern, group_number, replacement):
    full_text = "".join(node.text or "" for node in text_nodes)
    match = re.search(pattern, full_text, flags=re.S)
    if not match:
        return False
    return replace_span(text_nodes, match.start(group_number), match.end(group_number), replacement)


def ensure_child(parent, tag):
    child = parent.find(tag)
    if child is None:
        child = etree.Element(tag)
        if tag in {W + "pPr", W + "rPr", W + "trPr"}:
            parent.insert(0, child)
        else:
            parent.append(child)
    return child


def set_paragraph_spacing(paragraph, before="0", after="0", line="180"):
    ppr = ensure_child(paragraph, W + "pPr")
    spacing = ensure_child(ppr, W + "spacing")
    spacing.set(W + "before", before)
    spacing.set(W + "after", after)
    spacing.set(W + "line", line)
    spacing.set(W + "lineRule", "auto")


def set_paragraph_font_size(paragraph, half_points):
    for run in paragraph.xpath("./w:r", namespaces=NS):
        rpr = ensure_child(run, W + "rPr")
        sz = ensure_child(rpr, W + "sz")
        sz_cs = ensure_child(rpr, W + "szCs")
        sz.set(W + "val", str(half_points))
        sz_cs.set(W + "val", str(half_points))


def remove_paragraph(paragraph):
    parent = paragraph.getparent()
    if parent is not None:
        parent.remove(paragraph)


def is_blank_paragraph(element):
    return element is not None and element.tag == W + "p" and not para_text(element).strip()


def remove_adjacent_blank_paragraphs(paragraph, before=False, after=False):
    changes = 0
    if before:
        sibling = paragraph.getprevious()
        while is_blank_paragraph(sibling):
            previous = sibling.getprevious()
            remove_paragraph(sibling)
            changes += 1
            sibling = previous
    if after:
        sibling = paragraph.getnext()
        while is_blank_paragraph(sibling):
            following = sibling.getnext()
            remove_paragraph(sibling)
            changes += 1
            sibling = following
    return changes


def compact_request_form_layout(root):
    changes = 0

    for row in root.xpath(".//w:tr", namespaces=NS):
        trpr = ensure_child(row, W + "trPr")
        height = ensure_child(trpr, W + "trHeight")
        current = int(height.get(W + "val", "340"))
        if current > 255:
            height.set(W + "val", "255")
            changes += 1

    signature_active = False
    after_signature_note = False
    for paragraph in list(root.xpath(".//w:p", namespaces=NS)):
        text = para_text(paragraph).strip()

        if text.startswith("REQUEST FOR ISSUING"):
            signature_active = False
            after_signature_note = False

        if text.startswith("Note: Only final approved construction drawings"):
            changes += remove_adjacent_blank_paragraphs(paragraph, before=True)
            after_signature_note = True
            set_paragraph_spacing(paragraph, before="0", after="0", line="170")
            set_paragraph_font_size(paragraph, 16)
            changes += 1
            continue

        if after_signature_note and not text:
            remove_paragraph(paragraph)
            changes += 1
            continue

        if text.startswith("REQUESTED BY:"):
            signature_active = True
            after_signature_note = False

        if signature_active:
            if not text:
                remove_paragraph(paragraph)
                changes += 1
                continue
            set_paragraph_spacing(paragraph, before="0", after="0", line="170")
            set_paragraph_font_size(paragraph, 16)
            changes += 1

        if text == "FRANCIS LIM, CHIEF OPERATING OFFICER":
            changes += remove_adjacent_blank_paragraphs(paragraph, after=True)
            signature_active = False

    paragraphs = list(root.xpath(".//w:p", namespaces=NS))
    texts = [para_text(paragraph).strip() for paragraph in paragraphs]
    for index, paragraph in enumerate(paragraphs):
        if texts[index]:
            continue
        previous_text = next((texts[i] for i in range(index - 1, -1, -1) if texts[i]), "")
        next_text = next((texts[i] for i in range(index + 1, len(texts)) if texts[i]), "")
        if previous_text == "No" and next_text.startswith("Note: Only final approved construction drawings"):
            remove_paragraph(paragraph)
            changes += 1

    return changes


def set_cell_lines(cell, lines):
    """Replace the whole content of a value cell with the given lines.

    The template cell may hold several runs, line breaks and extra (blank) paragraphs from the previous
    submission; all of them are removed, otherwise stale blank lines push the form onto a second page."""
    paragraphs = cell.xpath("./w:p", namespaces=NS)
    runs = cell.xpath("./w:p/w:r[w:t]", namespaces=NS) or cell.xpath("./w:p/w:r", namespaces=NS)
    if not paragraphs or not runs:
        return False
    first_run = runs[0]
    first_paragraph = first_run.getparent()
    for paragraph in paragraphs:
        if paragraph is not first_paragraph:
            cell.remove(paragraph)
    for run in first_paragraph.xpath("./w:r", namespaces=NS):
        if run is not first_run:
            first_paragraph.remove(run)
    for child in list(first_run):
        if child.tag in {W + "t", W + "br"}:
            first_run.remove(child)

    for index, line in enumerate(lines):
        if index:
            etree.SubElement(first_run, W + "br")
        t = etree.SubElement(first_run, W + "t")
        t.text = line
        if line.startswith(" ") or line.endswith(" "):
            t.set(f"{{http://www.w3.org/XML/1998/namespace}}space", "preserve")
    return True


def update_table_rows(root, data):
    long_delivery = format_dates(data["required_steel_delivery_date"])["long"]
    desc = description_lines(data["description_of_steel_work"])
    batch = str(data["batch_number"]).strip()
    changes = 0

    row_rules = [
        ("descriptionofthesteelwork", desc),
        ("batchno", [batch]),
        ("requireddateofsteeldeliveredtosite", [long_delivery]),
        ("requireddateofdelivertosite", [long_delivery]),
    ]
    if data.get("estimated_tonnage") not in ("", None, []):
        row_rules.append(("estimatedtonnage", [str(data["estimated_tonnage"]).strip()]))

    for row in root.xpath(".//w:tr", namespaces=NS):
        cells = row.xpath("./w:tc", namespaces=NS)
        labels = [cell_text(cell) for cell in cells]
        for label, replacement in row_rules:
            for index, text in enumerate(labels):
                if label in compact(text) and index + 1 < len(cells):
                    if set_cell_lines(cells[index + 1], replacement):
                        changes += 1
    return changes


def update_paragraphs(root, data, doc_kind):
    current = format_dates(data["current_date"])
    required = format_dates(data["required_steel_delivery_date"])
    ref = str(data["our_reference_number"]).strip()
    changes = 0

    for paragraph in root.xpath(".//w:p", namespaces=NS):
        text = para_text(paragraph)
        collapsed = re.sub(r"\s+", " ", text).strip()

        if doc_kind == "written_notification":
            if collapsed.startswith("Our Ref:"):
                changes += bool(set_paragraph_text(paragraph, f"Our Ref: {ref}"))
            elif collapsed.startswith("Date:"):
                changes += bool(set_paragraph_text(paragraph, f"Date: {current['letter']}"))

        if doc_kind == "request_form" and "DATE:" in text and "DATE REQUIRED" in text:
            text_nodes = paragraph.xpath("./w:r//w:t", namespaces=NS)
            changes += bool(
                replace_regex_group(
                    text_nodes,
                    r"(DATE:\s*)(.*?)(DATE REQUIRED)",
                    2,
                    current["request"] + " ",
                )
            )
            changes += bool(
                replace_regex_group(
                    text_nodes,
                    r"(DATE REQUIRED \(Please specify exact date if possible\):\s*)(.*?)(S/N)",
                    2,
                    required["short"],
                )
            )
        elif doc_kind == "request_form" and "PROJECT:" in text and "DATE:" in text:
            text_nodes = paragraph.xpath("./w:r//w:t", namespaces=NS)
            changes += bool(
                replace_regex_group(
                    text_nodes,
                    r"(DATE:\s*)(.*)$",
                    2,
                    current["request"],
                )
            )
        elif doc_kind == "request_form" and collapsed.startswith("DATE REQUIRED (Please specify exact date if possible):"):
            changes += bool(
                set_paragraph_text(
                    paragraph,
                    f"DATE REQUIRED (Please specify exact date if possible): {required['short']}",
                )
            )

    return changes


def update_docx(template_path, output_path, data, doc_kind, compact_request_layout=False):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    changes = 0
    with zipfile.ZipFile(template_path, "r") as zin, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                try:
                    root = etree.fromstring(content)
                except etree.XMLSyntaxError:
                    pass
                else:
                    changes += update_paragraphs(root, data, doc_kind)
                    changes += update_table_rows(root, data)
                    if compact_request_layout and doc_kind == "request_form" and item.filename == "word/document.xml":
                        changes += compact_request_form_layout(root)
                    content = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, content)

    if changes == 0:
        raise RuntimeError(f"No editable fields were found in {template_path.name}")
    return changes


def load_data(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    required = [
        "our_reference_number",
        "current_date",
        "description_of_steel_work",
        "batch_number",
        "required_steel_delivery_date",
    ]
    missing = [key for key in required if key not in data or data[key] in ("", None, [])]
    if missing:
        raise ValueError("Missing required field(s): " + ", ".join(missing))
    return data


def safe_name(value):
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value).strip()).strip("_")


def main():
    parser = argparse.ArgumentParser(description="Update one steel work Word form set from one JSON file.")
    parser.add_argument("details_json", help="Path to submission_details.json")
    parser.add_argument("--form-set", choices=sorted(FORM_SETS), default="C3D1", help="Which form set to use")
    parser.add_argument("--template-dir", default=None, help="Optional folder override containing the two master .docx forms")
    parser.add_argument("--output-dir", default="updated_forms", help="Folder where updated copies will be written")
    parser.add_argument(
        "--compact-request-layout",
        action="store_true",
        help="Apply the older automatic request-form spacing compaction. Off by default to preserve the amended template.",
    )
    args = parser.parse_args()

    data = load_data(args.details_json)
    form_set = FORM_SETS[args.form_set]
    template_dir = Path(args.template_dir) if args.template_dir else form_set["template_dir"]
    templates = form_set["templates"]
    output_dir = Path(args.output_dir)
    suffix = safe_name(data["batch_number"])

    outputs = {}
    for kind, template_info in templates.items():
        template = template_dir / template_info["filename"]
        if not template.exists():
            raise FileNotFoundError(f"Template not found: {template}")
        output = output_dir / f"{template_info['output_stem']}_Batch_{suffix}.docx"
        outputs[kind] = output
        changes = update_docx(template, output, data, kind, compact_request_layout=args.compact_request_layout)
        print(f"Updated {output} ({changes} field groups changed)")

    print("Done.")


if __name__ == "__main__":
    main()
