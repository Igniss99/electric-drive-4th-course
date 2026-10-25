"""calc_db.json (LaTeX) -> формулы OMML внутри копии docx. Без Word: LaTeX -> MathML -> OMML (MML2OMML.XSL) -> XML."""
import json
import re
import shutil
import sys
import zipfile

import latex2mathml.converter as l2m
from lxml import etree

src, dst, js, prefix = sys.argv[1:5]
limit = int(sys.argv[5]) if len(sys.argv) > 5 else 10 ** 6
prefix = "" if prefix == "ALL" else prefix
XSL = r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
xslt = etree.XSLT(etree.parse(XSL))


def wrap_cyr(s):
    """Кириллица вне \\text{} -> \\text{...}, чтобы в Word она была прямым шрифтом одним фрагментом."""
    parts = re.split(r"(\\text\{[^{}]*\})", s)
    return "".join(p if p.startswith("\\text{") else re.sub(r"[А-Яа-яЁё]+", lambda m: "\\text{" + m.group(0) + "}", p)
                   for p in parts)


def latex_to_omml(tex):
    tex = tex.replace("max\\,", "\\mathrm{max}\\,")                    # max — прямым шрифтом
    tex = re.sub(r"(?<=\d),(?=\d)", lambda m: "\\text{,}", tex)        # десятичная запятая без пробела
    tex = wrap_cyr(tex)
    tex = tex.replace("\\left", "").replace("\\right", "")        # скобки обычные, подгон по высоте не нужен
    mml = l2m.convert(tex)
    omml = xslt(etree.fromstring(mml.encode("utf8")))
    return omml.getroot()                                          # <m:oMath>


def para_text(text, bold=False):
    p = etree.Element(f"{{{W}}}p")
    r = etree.SubElement(p, f"{{{W}}}r")
    if bold:
        rpr = etree.SubElement(r, f"{{{W}}}rPr")
        etree.SubElement(rpr, f"{{{W}}}b")
    t = etree.SubElement(r, f"{{{W}}}t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return p


def para_math(tex):
    p = etree.Element(f"{{{W}}}p", nsmap={"m": M})
    pm = etree.SubElement(p, f"{{{M}}}oMathPara")
    pm.append(latex_to_omml(tex))
    return p


items = [i for i in json.load(open(js, encoding="utf8"))
         if i["kind"] in ("ДАННЫЕ", "ФОРМУЛА", "РЕЗУЛЬТАТ") and i["section"].startswith(prefix)][:limit]

shutil.copyfile(src, dst)
zin = zipfile.ZipFile(src)
doc = etree.fromstring(zin.read("word/document.xml"))
body = doc.find(f"{{{W}}}body")
sect = body.find(f"{{{W}}}sectPr")
new = [para_text(f"ТЕСТ: формулы из calc_db (разделы {prefix}…)", bold=True)]
for it in items:
    new.append(para_text(f"{it['id']} {it['title']}"))
    new.append(para_math(it["latex"]))
for p in new:
    sect.addprevious(p)

with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
    for info in zin.infolist():
        data = zin.read(info.filename)
        if info.filename == "word/document.xml":
            data = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
        zout.writestr(info, data)
print("inserted formulas:", len(items))

