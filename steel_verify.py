"""Checks a generated pair of steel-work Word forms against the approved templates.

Usage: python steel_verify.py job.json
job.json: {"templates": {"written": path, "request": path},
           "outputs":   {"written": path, "request": path},
           "expected":  {"ref", "letterDate", "requestDate", "longDelivery", "shortDelivery",
                         "descLines": [...], "batch", "tonnage"}}
Prints one JSON object: {"ok": bool, "problems": [...], "signatures": {"written": [...], "request": [...]}}.

Every field is checked where the template really holds it. The request form keeps two copies of its
content (a text box and its fallback), so its fields must appear exactly twice. All other template
text, including every signature name and title, must survive unchanged.
"""
import json
import re
import sys
import zipfile
from collections import Counter
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist() if re.match(r"word/(document|header\d*|footer\d*)\.xml$", n)}


def para_text(p):
    """Text of one paragraph. A text box's own paragraphs are nested inside its anchor paragraph and are
    listed separately, so the walk does not descend into them."""
    out = []

    def walk(node):
        for child in node:
            if child.tag == W + "p":
                continue
            if child.tag == W + "t":
                out.append(child.text or "")
            elif child.tag == W + "br":
                out.append("\n")
            else:
                walk(child)

    walk(p)
    return "".join(out)


def paragraphs(xml):
    return [para_text(p) for p in ET.fromstring(xml).iter(W + "p")]


def collapsed(text):
    return re.sub(r"\s+", " ", text).strip()


def compact(text):
    return re.sub(r"[^a-z0-9]+", "", text.lower())


def rows(xml):
    """(label cell text, value cell lines) for every table row that has at least two cells."""
    out = []
    for tr in ET.fromstring(xml).iter(W + "tr"):
        cells = [tc for tc in tr if tc.tag == W + "tc"]
        texts = ["\n".join(para_text(p) for p in tc.iter(W + "p")) for tc in cells]
        for i in range(len(texts) - 1):
            lines = [l.strip() for l in texts[i + 1].split("\n") if l.strip()]
            out.append((compact(texts[i]), lines))
    return out


def values_for(rows_, label):
    return [lines for key, lines in rows_ if label in key]


def signatures(paras, kind):
    texts = [collapsed(t) for t in paras if collapsed(t)]
    names = []
    if kind == "written":
        if "Yours faithfully" in texts:
            names = texts[texts.index("Yours faithfully") + 1:]
    else:
        for i, t in enumerate(texts):
            if re.match(r"(REQUESTED|REVIEWED|APPROVED) BY:", t) and i + 1 < len(texts) and not set(texts[i + 1]) <= set("_ "):
                names.append(texts[i + 1])
    seen = []
    for n in names:
        if n and n not in seen:
            seen.append(n)
    return seen


def check(kind, template, output, exp, problems):
    t_parts, o_parts = parts(template), parts(output)
    if set(t_parts) != set(o_parts):
        problems.append(f"{kind}: XML parts differ from the template")
    t_doc, o_doc = paragraphs(t_parts["word/document.xml"]), paragraphs(o_parts["word/document.xml"])
    o_rows = rows(o_parts["word/document.xml"])
    copies = 2 if kind == "request" else 1

    def expect_rows(label, lines):
        found = values_for(o_rows, label)
        if len(found) != copies:
            problems.append(f"{kind}: row '{label}' found {len(found)} time(s), expected {copies}")
        for got in found:
            if got != lines:
                problems.append(f"{kind}: row '{label}' holds {got!r}, expected {lines!r}")

    expect_rows("descriptionofthesteelwork", exp["descLines"])
    expect_rows("batchno", [exp["batch"]])
    if kind == "written":
        expect_rows("requireddateofsteeldeliveredtosite", [exp["longDelivery"]])
        refs = [collapsed(t) for t in o_doc if collapsed(t).startswith("Our Ref:")]
        if refs != [f"Our Ref: {exp['ref']}"]:
            problems.append(f"written: reference paragraphs are {refs!r}")
        dates = [collapsed(t) for t in o_doc if collapsed(t).startswith("Date:")]
        if dates != [f"Date: {exp['letterDate']}"]:
            problems.append(f"written: date paragraphs are {dates!r}")
    else:
        expect_rows("estimatedtonnage", [exp["tonnage"]])
        expect_rows("requireddateofdelivertosite", [exp["longDelivery"]])
        heads = [collapsed(t) for t in o_doc if "PROJECT:" in t]
        if len(heads) != copies or any(not re.search(r"DATE:\s*" + re.escape(exp["requestDate"]) + r"$", h) for h in heads):
            problems.append(f"request: PROJECT/DATE lines are {heads!r}")
        required = [collapsed(t) for t in o_doc if collapsed(t).startswith("DATE REQUIRED")]
        want = f"DATE REQUIRED (Please specify exact date if possible): {exp['shortDelivery']}"
        if required != [want] * copies:
            problems.append(f"request: DATE REQUIRED lines are {required!r}")

    # Everything that is not a field must survive: compare the two paragraph multisets and allow only field text to differ.
    t_values = {tuple(lines) for key, lines in rows(t_parts["word/document.xml"])
                if any(k in key for k in ("descriptionofthesteelwork", "batchno", "estimatedtonnage", "requireddateof"))}
    field_texts = {l for lines in t_values for l in lines} | {exp["batch"], exp["tonnage"], exp["longDelivery"]} | set(exp["descLines"])

    def is_field(text):
        c = collapsed(text)
        return (not c or c.startswith(("Our Ref:", "Date:", "DATE REQUIRED")) or "PROJECT:" in c
                or any(l.strip() and l.strip() in field_texts for l in text.split("\n")))

    lost = Counter(collapsed(t) for t in t_doc if not is_field(t)) - Counter(collapsed(t) for t in o_doc if not is_field(t))
    added = Counter(collapsed(t) for t in o_doc if not is_field(t)) - Counter(collapsed(t) for t in t_doc if not is_field(t))
    for text, n in lost.items():
        problems.append(f"{kind}: template text missing from output ({n}x): {text[:80]!r}")
    for text, n in added.items():
        problems.append(f"{kind}: unexpected new text in output ({n}x): {text[:80]!r}")
    for name in t_parts:
        if name != "word/document.xml" and paragraphs(t_parts[name]) != paragraphs(o_parts.get(name, b"<r/>")):
            problems.append(f"{kind}: {name} text differs from the template")
    return signatures(o_doc, kind)


def main():
    job = json.load(open(sys.argv[1], encoding="utf-8"))
    problems, sigs = [], {}
    for kind in ("written", "request"):
        try:
            sigs[kind] = check(kind, job["templates"][kind], job["outputs"][kind], job["expected"], problems)
        except Exception as e:  # a corrupt package is a failed check, not a crash
            problems.append(f"{kind}: could not be checked ({e})")
            sigs[kind] = []
    print(json.dumps({"ok": not problems, "problems": problems, "signatures": sigs}, ensure_ascii=False))


if __name__ == "__main__":
    main()
