from copy import deepcopy
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET


SRC = Path(r"C:\Users\shrey\OneDrive\Documents\New project\report\Final_Project_Report_ASD_Updated.docx")
OUT = Path(r"C:\Users\shrey\OneDrive\Documents\New project\report\Final_Project_Report_ASD_Updated_Numbered.docx")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

W = "{%s}" % NS["w"]
R = "{%s}" % NS["r"]

for prefix, uri in NS.items():
    ET.register_namespace(prefix if prefix in ("w", "r") else "", uri)


TOC_LINES = [
    ("Chapter 1: Introduction", 6),
    ("1A. Background of the Problem", 7),
    ("1B. Importance & Relevance", 8),
    ("1C. Problem Statement", 9),
    ("1D. Objectives of the Project", 10),
    ("1E. Scope of the Project", 11),
    ("Chapter 2: Literature Review", 12),
    ("2A. Research Papers / Articles Reviewed", 12),
    ("2B. Summary of Existing Work", 13),
    ("2C. Comparative Analysis", 14),
    ("2D. Identification of Research Gap", 15),
    ("Chapter 3: System Analysis", 16),
    ("3A. Existing System", 16),
    ("3B. Limitations of Existing System", 17),
    ("3C. Proposed System", 18),
    ("Chapter 4: System Design", 19),
    ("4A. Data Flow Diagram (DFD)", 19),
    ("4B. Sequence Diagram", 22),
    ("4C. Use Case Diagram", 22),
    ("4D. Schema Diagram", 22),
    ("Chapter 5: Methodology", 23),
    ("5A. Step-by-Step Approach", 23),
    ("5B. Algorithms / Techniques", 23),
    ("5C. Mathematical Model", 23),
    ("5D. Workflow Explanation", 23),
    ("Chapter 6: Implementation", 24),
    ("6A. Tools & Technologies", 24),
    ("6B. Hardware & Software Requirements", 25),
    ("6C. Dataset Description", 26),
    ("6D. Coding Approach / Modules", 27),
    ("6E. Screenshots of Implementation", 28),
    ("Chapter 7: Results and Discussion", 36),
    ("7A. Output Results", 36),
    ("7B. Performance Metrics and Comparison (expected vs actual)", 36),
    ("7C. Tables and Graphs", 37),
    ("7D. Interpretation of Results", 38),
    ("Chapter 8: Advantages and Limitations", 39),
    ("8A. Benefits of Proposed System", 39),
    ("8B. Drawbacks or Constraints", 40),
    ("Chapter 9: Conclusion", 41),
    ("9A. Summary of Work", 41),
    ("9B. Achievement of Objectives", 42),
    ("9C. Final Outcomes", 43),
    ("Chapter 10: Future Scope", 44),
    ("10A. Possible Improvements", 44),
    ("10B. Extensions of Project", 45),
    ("10C. Research Directions", 46),
    ("References", 47),
    ("Appendix", 47),
]


def make_run(text, font_size=22, bold=False):
    run = ET.Element(W + "r")
    rpr = ET.SubElement(run, W + "rPr")
    ET.SubElement(rpr, W + "rFonts", {W + "ascii": "Times New Roman", W + "hAnsi": "Times New Roman"})
    ET.SubElement(rpr, W + "sz", {W + "val": str(font_size)})
    ET.SubElement(rpr, W + "szCs", {W + "val": str(font_size)})
    if bold:
        ET.SubElement(rpr, W + "b")
    t = ET.SubElement(run, W + "t")
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return run


def make_paragraph(text, style=None, center=False, with_tabs=False):
    p = ET.Element(W + "p")
    ppr = ET.SubElement(p, W + "pPr")
    if style:
        ET.SubElement(ppr, W + "pStyle", {W + "val": style})
    if center:
        ET.SubElement(ppr, W + "jc", {W + "val": "center"})
    if with_tabs:
        tabs = ET.SubElement(ppr, W + "tabs")
        ET.SubElement(tabs, W + "tab", {W + "val": "right", W + "leader": "dot", W + "pos": "9350"})
    p.append(make_run(text))
    return p


def make_toc_paragraph(title, page):
    p = ET.Element(W + "p")
    ppr = ET.SubElement(p, W + "pPr")
    ET.SubElement(ppr, W + "pStyle", {W + "val": "BodyText"})
    tabs = ET.SubElement(ppr, W + "tabs")
    ET.SubElement(tabs, W + "tab", {W + "val": "right", W + "leader": "dot", W + "pos": "9350"})
    p.append(make_run(title))
    tab_run = ET.Element(W + "r")
    rpr = ET.SubElement(tab_run, W + "rPr")
    ET.SubElement(rpr, W + "rFonts", {W + "ascii": "Times New Roman", W + "hAnsi": "Times New Roman"})
    ET.SubElement(rpr, W + "sz", {W + "val": "22"})
    ET.SubElement(rpr, W + "szCs", {W + "val": "22"})
    ET.SubElement(tab_run, W + "tab")
    p.append(tab_run)
    p.append(make_run(str(page)))
    return p


def make_footer_xml():
    root = ET.Element(W + "ftr")
    p = ET.SubElement(root, W + "p")
    ppr = ET.SubElement(p, W + "pPr")
    ET.SubElement(ppr, W + "jc", {W + "val": "center"})
    r = ET.SubElement(p, W + "r")
    rpr = ET.SubElement(r, W + "rPr")
    ET.SubElement(rpr, W + "rFonts", {W + "ascii": "Times New Roman", W + "hAnsi": "Times New Roman"})
    ET.SubElement(rpr, W + "sz", {W + "val": "24"})
    ET.SubElement(rpr, W + "szCs", {W + "val": "24"})
    fld = ET.SubElement(p, W + "fldSimple", {"{%s}instr" % NS["w"]: " PAGE "})
    fr = ET.SubElement(fld, W + "r")
    frpr = ET.SubElement(fr, W + "rPr")
    ET.SubElement(frpr, W + "rFonts", {W + "ascii": "Times New Roman", W + "hAnsi": "Times New Roman"})
    ET.SubElement(frpr, W + "sz", {W + "val": "24"})
    ET.SubElement(frpr, W + "szCs", {W + "val": "24"})
    ET.SubElement(fr, W + "t").text = "1"
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


with zipfile.ZipFile(SRC, "r") as zin:
    files = {name: zin.read(name) for name in zin.namelist()}

doc_root = ET.fromstring(files["word/document.xml"])
body = doc_root.find(W + "body")
paragraphs = body.findall(W + "p")

# Find TOC start and content span.
toc_title_idx = None
intro_idx = None
for idx, para in enumerate(paragraphs):
    text = "".join(t.text or "" for t in para.findall(".//" + W + "t")).strip()
    if text == "TABLE OF CONTENTS" and toc_title_idx is None:
        toc_title_idx = idx
    elif text == "1. INTRODUCTION" and toc_title_idx is not None:
        intro_idx = idx
        break

if toc_title_idx is None or intro_idx is None:
    raise RuntimeError("Could not locate the table of contents block.")

# Remove existing TOC paragraphs after title and before introduction.
for idx in range(intro_idx - 1, toc_title_idx, -1):
    body.remove(paragraphs[idx])

insert_pos = toc_title_idx + 1
for offset, (title, page) in enumerate(TOC_LINES):
    body.insert(insert_pos + offset, make_toc_paragraph(title, page))

# Footer wiring.
footer_xml = make_footer_xml()
files["word/footer1.xml"] = footer_xml

rels_root = ET.fromstring(files["word/_rels/document.xml.rels"])
existing_ids = {
    rel.attrib["Id"]
    for rel in rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
}
next_num = 1
while f"rId{next_num}" in existing_ids:
    next_num += 1
footer_rid = f"rId{next_num}"
ET.SubElement(
    rels_root,
    "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
    {
        "Id": footer_rid,
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        "Target": "footer1.xml",
    },
)
files["word/_rels/document.xml.rels"] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

content_root = ET.fromstring(files["[Content_Types].xml"])
has_footer_override = any(
    el.attrib.get("PartName") == "/word/footer1.xml"
    for el in content_root.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Override")
)
if not has_footer_override:
    ET.SubElement(
        content_root,
        "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
        {
            "PartName": "/word/footer1.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        },
    )
files["[Content_Types].xml"] = ET.tostring(content_root, encoding="utf-8", xml_declaration=True)

sect_prs = body.findall(".//" + W + "sectPr")
for sect_pr in sect_prs:
    for child in list(sect_pr):
        if child.tag == W + "footerReference":
            sect_pr.remove(child)
    sect_pr.insert(0, ET.Element(W + "footerReference", {R + "id": footer_rid, W + "type": "default"}))

files["word/document.xml"] = ET.tostring(doc_root, encoding="utf-8", xml_declaration=True)

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
    for name, data in files.items():
        zout.writestr(name, data)

print(f"Created: {OUT}")
