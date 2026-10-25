"""Сборка курсовой работы в docx: шаблон оформления — RabotaIgnata24.10.26.docx, числа/формулы/графики — из out/variant_<N>/.

Запуск:  python build_coursework.py            (после python Coursework_for_4th_course.py)
Результат: out/variant_<N>/Курсовая_вариант_<N>.docx ; оглавление и поля обновляются через Word (COM), если он установлен.
Формулы вставляются как настоящие формулы Word (LaTeX -> MathML -> OMML через MML2OMML.XSL).
"""
import json
import math
import re
import shutil
import struct
import sys
import zipfile
from pathlib import Path

import latex2mathml.converter as l2m
from lxml import etree

sys.path.insert(0, str(Path(__file__).parent))
from Coursework_for_4th_course import VARIANT, OUT_DIR, fmt, BETA_DC, K_MP, K_MPER, group_of  # noqa: E402
from variants_data import AC_GROUPS, DC_GROUPS  # noqa: E402

HERE = Path(__file__).parent
TEMPLATE = HERE / "RabotaIgnata24.10.26.docx"
XSL = r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL"
RESULT_DOCX = OUT_DIR / f"Курсовая_вариант_{VARIANT}.docx"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
xslt = etree.XSLT(etree.parse(XSL))

DB = json.load(open(OUT_DIR / "calc_db.json", encoding="utf8"))
RES = json.load(open(OUT_DIR / "results.json", encoding="utf8"))


def w(tag):
    return f"{{{W}}}{tag}"


# ----------------------------------------------------------------------------- доступ к базе расчётов
def _find(kind, sec, key):
    for it in DB:
        if it["kind"] == kind and it.get("key") == key and it["section"].startswith(sec):
            return it
    raise KeyError((kind, sec, key))


def res(sec, key):
    """LaTeX результата шага (формула = подстановка = значение с единицами)."""
    return _find("РЕЗУЛЬТАТ", sec, key)["latex"]


def formula(sec, key):
    return _find("ФОРМУЛА", sec, key)["latex"]


def val(sec, key):
    return _find("РЕЗУЛЬТАТ", sec, key)["value"]


# ----------------------------------------------------------------------------- LaTeX -> OMML
def wrap_cyr(s):
    parts = re.split(r"(\\text\{[^{}]*\})", s)
    return "".join(p if p.startswith("\\text{") else re.sub(r"[А-Яа-яЁё]+", lambda m: "\\text{" + m.group(0) + "}", p)
                   for p in parts)


def latex_to_omml(tex):
    tex = tex.replace("max\\,", "\\mathrm{max}\\,")
    tex = re.sub(r"(?<=\d),(?=\d)", lambda m: "\\text{,}", tex)
    tex = wrap_cyr(tex)
    tex = tex.replace("\\left", "").replace("\\right", "")
    mml = l2m.convert(tex)
    return xslt(etree.fromstring(mml.encode("utf8"))).getroot()


# ----------------------------------------------------------------------------- элементы документа
def _rpr(b=False, i=False, sz=28, sub=False):
    rpr = etree.Element(w("rPr"))
    f = etree.SubElement(rpr, w("rFonts"))
    for a in ("ascii", "hAnsi", "cs"):
        f.set(w(a), "Times New Roman")
    if b:
        etree.SubElement(rpr, w("b"))
    if i:
        etree.SubElement(rpr, w("i"))
    etree.SubElement(rpr, w("sz")).set(w("val"), str(sz))
    etree.SubElement(rpr, w("szCs")).set(w("val"), str(sz))
    if sub:
        etree.SubElement(rpr, w("vertAlign")).set(w("val"), "subscript")
    return rpr


def _run(text, **kw):
    r = etree.Element(w("r"))
    r.append(_rpr(**kw))
    t = etree.SubElement(r, w("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")
    return r


def _ppr(style="a0", jc="both", first=0, left=0, before=0, after=0, keep_next=False, page_break=False, line=288):
    ppr = etree.Element(w("pPr"))
    etree.SubElement(ppr, w("pStyle")).set(w("val"), style)
    if keep_next:
        etree.SubElement(ppr, w("keepNext"))
    if page_break:
        etree.SubElement(ppr, w("pageBreakBefore"))
    sp = etree.SubElement(ppr, w("spacing"))
    sp.set(w("before"), str(before))
    sp.set(w("after"), str(after))
    sp.set(w("line"), str(line))
    sp.set(w("lineRule"), "auto")
    ind = etree.SubElement(ppr, w("ind"))
    ind.set(w("left"), str(left))
    ind.set(w("firstLine"), str(first))
    etree.SubElement(ppr, w("jc")).set(w("val"), jc)
    return ppr


def para(text="", b=False, i=False, jc="both", first=0, sz=28, **kw):
    """Абзац; в text формулы заключаются в $...$ (LaTeX), **...** — полужирный."""
    p = etree.Element(w("p"))
    p.append(_ppr(jc=jc, first=first, **kw))
    for k, seg in enumerate(re.split(r"\$", text)):
        if k % 2 == 1:                                            # формула в строке
            p.append(latex_to_omml(seg))
        else:
            for j, sub in enumerate(re.split(r"\*\*", seg)):
                if sub:
                    p.append(_run(sub, b=b or j % 2 == 1, i=i, sz=sz))
    return p


def display(tex, before=60, after=60):
    """Отдельная формула по центру; части, разделённые \\qquad / \\quad, выводятся отдельными абзацами."""
    out = []
    for part in re.split(r",?\s*\\q?quad\s*", tex):
        part = part.strip().rstrip(",;")
        if not part:
            continue
        p = etree.Element(w("p"), nsmap={"m": M_NS})
        p.append(_ppr(jc="center", before=before, after=after, line=259))
        pm = etree.SubElement(p, f"{{{M_NS}}}oMathPara")
        pm.append(latex_to_omml(part))
        out.append(p)
    return out if len(out) > 1 else out[0]


def heading1(text):
    p = etree.Element(w("p"))
    ppr = etree.SubElement(p, w("pPr"))
    etree.SubElement(ppr, w("pStyle")).set(w("val"), "1")
    etree.SubElement(ppr, w("pageBreakBefore"))
    sp = etree.SubElement(ppr, w("spacing"))
    sp.set(w("after"), "120")
    p.append(_run(text, b=True, sz=28))
    return p


def heading2(text):
    p = etree.Element(w("p"))
    ppr = etree.SubElement(p, w("pPr"))
    etree.SubElement(ppr, w("pStyle")).set(w("val"), "2")
    etree.SubElement(ppr, w("keepNext"))
    sp = etree.SubElement(ppr, w("spacing"))
    sp.set(w("before"), "200")
    sp.set(w("after"), "120")
    etree.SubElement(ppr, w("jc")).set(w("val"), "left")
    p.append(_run(text, b=True, sz=28))
    return p


def table(headers, rows, widths_cm, sz=24):
    tbl = etree.Element(w("tbl"))
    pr = etree.SubElement(tbl, w("tblPr"))
    etree.SubElement(pr, w("tblW")).attrib.update({w("w"): "0", w("type"): "auto"})
    etree.SubElement(pr, w("jc")).set(w("val"), "center")
    bd = etree.SubElement(pr, w("tblBorders"))
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = etree.SubElement(bd, w(e))
        b.attrib.update({w("val"): "single", w("sz"): "4", w("space"): "0", w("color"): "000000"})
    grid = etree.SubElement(tbl, w("tblGrid"))
    for wc in widths_cm:
        etree.SubElement(grid, w("gridCol")).set(w("w"), str(int(wc * 567)))

    def row(cells, bold):
        tr = etree.SubElement(tbl, w("tr"))
        if bold:
            etree.SubElement(etree.SubElement(tr, w("trPr")), w("tblHeader"))
        for c, wc in zip(cells, widths_cm):
            tc = etree.SubElement(tr, w("tc"))
            tcp = etree.SubElement(tc, w("tcPr"))
            etree.SubElement(tcp, w("tcW")).attrib.update({w("w"): str(int(wc * 567)), w("type"): "dxa"})
            etree.SubElement(tcp, w("vAlign")).set(w("val"), "center")
            p = para(str(c), b=bold, jc="center", sz=sz, line=240)
            tc.append(p)

    row(headers, True)
    for r in rows:
        row(r, False)
    return tbl


class Images:
    """Рисунки: копируются в пакет docx, возвращается w:p с inline-изображением."""

    def __init__(self):
        self.files = {}      # имя в пакете -> путь
        self.rels = []       # (rId, target)
        self.n = 0

    def figure(self, path, width_cm, caption):
        self.n += 1
        rid, name = f"rIdFig{self.n}", f"media/fig_{self.n}.png"
        self.files["word/" + name] = Path(path)
        self.rels.append((rid, name))
        with open(path, "rb") as f:
            head = f.read(24)
        wpx, hpx = struct.unpack(">II", head[16:24])
        cx = int(width_cm * 360000)
        cy = int(cx * hpx / wpx)
        xml = (f'<w:p xmlns:w="{W}" xmlns:r="{R_NS}" xmlns:wp="{WP}" xmlns:a="{A_NS}" xmlns:pic="{PIC}">'
               f'<w:pPr><w:keepNext/><w:spacing w:before="120" w:after="60" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'
               f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
               f'<wp:docPr id="{100 + self.n}" name="Рисунок {self.n}"/><wp:cNvGraphicFramePr/>'
               f'<a:graphic><a:graphicData uri="{PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="{100 + self.n}" name="fig{self.n}.png"/>'
               f'<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
               f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
               f'</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
        pic = etree.fromstring(xml)
        cap = para(caption, i=True, jc="center", after=160, line=240)
        return [pic, cap]


IMG = Images()


# ----------------------------------------------------------------------------- вспомогательные тексты
def chg(old, new, unit="", nd=4, abs_=False):
    """«увеличился с a до b» / «уменьшился …» / «не изменился (a)»."""
    if abs_:
        old, new = abs(old), abs(new)
    a, b = fmt(old, nd), fmt(new, nd)
    if abs(new - old) <= 1e-3 * max(abs(old), 1e-12):
        return f"не изменился (остался {a}{(' ' + unit) if unit else ''})"
    word = "увеличился" if new > old else "уменьшился"
    return f"{word} с {a} до {b}{(' ' + unit) if unit else ''}"


def rel_pct(a, b):
    return abs(b - a) / abs(a) * 100


# ----------------------------------------------------------------------------- содержимое
def steps_word(n):
    """Согласование слова «ступень» с числом: 1 ступень, 2–4 ступени, 5 и более ступеней."""
    n = int(n)
    if 11 <= n % 100 <= 14:
        return "ступеней"
    return {1: "ступень", 2: "ступени", 3: "ступени", 4: "ступени"}.get(n % 10, "ступеней")


def ac_conditions_text():
    ku, kr1, kx, kf, kp = AC_GROUPS[group_of(VARIANT)]
    return (f"Условия варианта {VARIANT} (Приложение 4): $U = {fmt(ku, 3)}U_{{ном}}$; $r_1 = {fmt(kr1, 3)} r_1$; "
            f"$x_{{кз}} = {fmt(kx, 3)} x_{{кз}}$; $f_1 = {fmt(kf, 3)} f_1$; $p_{{п}} = {fmt(kp, 3)} p_{{п}}$.")


def build_body():
    dc, ac = RES["dc"], RES["ac"]
    mdc, mac = dc["motor"], ac["motor"]
    nat = dc["natural"]
    body = []
    def A(x):
        body.extend(x) if isinstance(x, list) else body.append(x)


    def ext(items):
        body.extend(items)

    # ===================================================================== 1
    A(heading1("1. ЭЛЕКТРОМЕХАНИЧЕСКИЕ И МЕХАНИЧЕСКИЕ ХАРАКТЕРИСТИКИ ДВИГАТЕЛЯ ПОСТОЯННОГО ТОКА С ПАРАЛЛЕЛЬНЫМ ВОЗБУЖДЕНИЕМ"))
    A(para(f"Исходные паспортные данные (Приложение 1, вариант {VARIANT}):", b=True, jc="left"))
    for line in [f"Тип двигателя: {mdc['type']}",
                 f"Мощность: $P_{{ном}} = {fmt(mdc['P_kW'], 4)}$ кВт",
                 f"Частота вращения двигателя: $n_{{ном}} = {fmt(mdc['n_nom'], 4)}$ об/мин",
                 f"Ток якоря: $I_{{ном}} = {fmt(mdc['I_nom'], 4)}$ А",
                 f"Сопротивление якоря двигателя: $R_{{я}} = {fmt(mdc['R_ya'], 3)}$ Ом",
                 f"Момент инерции: $J = {fmt(mdc['J'], 3)}$ кг·м²",
                 f"Номинальное напряжение сети: $U_{{ном}} = {fmt(mdc['U_nom'], 3)}$ В"]:
        A(para(line, jc="left"))

    A(heading2("1.1 Расчет и построение естественных характеристик электродвигателя"))
    A(para("Естественные электромеханическая и механическая характеристики двигателя постоянного тока с параллельным "
           "возбуждением при пренебрежении реакцией якоря являются прямолинейными и описываются уравнениями:"))
    A(display(r"n = \frac{U_{ном}}{c\Phi_{н}} - \frac{R_{я}}{c\Phi_{н}} I_{я},\qquad "
              r"n = \frac{U_{ном}}{c\Phi_{н}} - \frac{R_{я}}{k c\Phi_{н}^2} M"))
    A(para("где $n$ – частота вращения двигателя, об/мин; $U_{ном}$ – номинальное напряжение сети, В; $R_{я}$ – сопротивление "
           "якоря двигателя, Ом; $c$, $k$ – конструктивные коэффициенты двигателя; $\\Phi_{н}$ – номинальный магнитный поток "
           "двигателя, Вб; $M$ – момент на валу, Н·м."))
    A(para("Уравнения можно записать и для угловой скорости, если учесть, что $\\omega = 2\\pi n / 60$."))
    A(para("В связи с прямолинейностью естественные характеристики строятся по двум точкам, соответствующим номинальному и "
           "пограничному (пусковому) режимам работы. Номинальному режиму соответствуют точки: $n = n_{ном}$ при "
           "$I = I_{ном}$ – для электромеханической характеристики и $n = n_{ном}$ при $M = M_{ном}$ – для механической."))
    A(para("Пограничному режиму соответствует условие $I = 0$ ($M = 0$), при котором частота вращения (идеальный холостой ход) "
           "равна:"))
    A(display(r"n_{0е} = \frac{U_{ном}}{c\Phi_{ном}}"))
    A(para("Пусковому режиму (режиму стоянки под током) отвечает условие $n = 0$:"))
    A(display(r"I_{пе} = \frac{U_{ном}}{R_{я}},\qquad M_{пе} = \frac{U_{ном} k\Phi_{ном}}{R_{я}} = I_{пе} k\Phi_{ном}"))
    A(para("Конструктивные коэффициенты определяем по номинальному режиму. Номинальная угловая скорость и номинальный момент:"))
    A(display(res("1.1", "dc.w_nom")))
    A(display(res("1.1", "dc.M_nom")))
    A(para("Коэффициент $c\\Phi_{ном}$ находим из формулы $n_{ном} = (U_{ном} - I_{ном} R_{я}) / c\\Phi_{ном}$, а коэффициент "
           "$k\\Phi_{ном}$ – из соотношения $M_{ном} = I_{ном} k\\Phi_{ном}$:"))
    A(display(res("1.1", "dc.cPhi_nom")))
    A(display(res("1.1", "dc.kPhi_nom")))
    A(para("Пограничная частота вращения (идеальный холостой ход):"))
    A(display(res("1.1", "dc.n0")))
    A(para("Ток и момент стоянки под током:"))
    A(display(res("1.1", "dc.I_p")))
    A(display(res("1.1", "dc.M_p")))
    A(para("Проверка: подстановка $c\\Phi_{ном}$ в формулу номинальной частоты вращения возвращает исходное значение:"))
    A(display(res("1.1", "dc.n_check")))
    A(para("Характеристики строятся для одного направления вращения в диапазоне от идеального холостого хода до режима "
           "стоянки под током (рис. 1)."))
    ext(IMG.figure(OUT_DIR / "fig_1_1.png", 16.0,
                   "Рис. 1. Электромеханические и механические характеристики двигателя постоянного тока с параллельным "
                   "возбуждением (естественные)"))
    dn = nat["n0"] - mdc["n_nom"]
    A(para(f"**Вывод:** естественные характеристики прямолинейны. Скорость идеального холостого хода "
           f"$n_{{0е}} = {fmt(nat['n0'], 4)}$ об/мин, при номинальной нагрузке скорость падает всего на "
           f"{fmt(dn, 2)} об/мин ({fmt(dn / nat['n0'] * 100, 2)} %), то есть характеристика жёсткая. Пусковой ток "
           f"$I_{{пе}} = {fmt(nat['I_p'], 5)}$ А превышает номинальный в {fmt(nat['I_p'] / mdc['I_nom'], 3)} раза, поэтому "
           f"прямой пуск двигателя недопустим."))

    A(heading2("1.2 Расчет и построение искусственных характеристик электродвигателя"))
    A(para("Искусственные характеристики строятся по двум точкам, вычисленным по формулам электромеханической и механической "
           "характеристик для новых значений изменяемого параметра. Для этого определяют ток (момент) стоянки при $n = 0$ и "
           "пограничную частоту вращения при $I = 0$ ($M = 0$). Если одновременно изменяются напряжение $U$, сопротивление "
           "якоря $R$ и магнитный поток $\\Phi$, получим:"))
    A(display(r"I_{п} = \frac{U}{R},\qquad n_0 = \frac{U}{k_{\Phi} c\Phi_{ном}},\qquad "
              r"M_{п} = \frac{U k_{\Phi} k\Phi_{ном}}{R}"))
    A(para("Здесь $k_{\\Phi}$ – относительное значение магнитного потока, $c\\Phi_{ном}$ и $k\\Phi_{ном}$ рассчитаны выше для "
           "номинального режима. Каждая характеристика отвечает изменению только одного параметра по сравнению с номинальным "
           "режимом; на каждом графике приведена естественная характеристика для сравнения."))
    kf, kr, ku = DC_GROUPS[group_of(VARIANT)]
    A(para(f"Условия варианта {VARIANT} (Приложение 2): 1) $\\Phi = {fmt(kf, 3)}\\Phi_{{ном}}$; 2) $R = {fmt(kr, 3)} R_{{я}}$; "
           f"3) $U = {fmt(ku, 3)} U_{{ном}}$."))
    dc_conc = {
        1: ("Магнитный поток уменьшен, поэтому скорость идеального холостого хода выросла, а момент стоянки уменьшился при "
            "неизменном токе стоянки: характеристика стала мягче (скорость сильнее зависит от нагрузки)."),
        2: ("Сопротивление цепи якоря увеличено, поэтому ток и момент стоянки уменьшились, скорость идеального холостого хода "
            "не изменилась: характеристика стала мягче."),
        3: ("Напряжение повышено, поэтому характеристика сместилась параллельно естественной (жёсткость не изменилась): "
            "скорость холостого хода, ток и момент стоянки выросли."),
    }
    names_dc = {1: "Фиг. 1", 2: "Фиг. 2", 3: "Фиг. 3"}
    letters = {1: "а", 2: "б", 3: "в"}
    for i, art in enumerate(dc["artificial"], 1):
        sec = f"1.2 ДПТ: искусственная характеристика {i}"
        A(para(f"**{i}) {art['name']}**", jc="left", keep_next=True, before=120))
        if i == 1:
            pass
        A(display(res(sec, f"dc.art{i}.U")))
        A(display(res(sec, f"dc.art{i}.R")))
        A(display(res(sec, f"dc.art{i}.n0")))
        A(display(res(sec, f"dc.art{i}.I_p")))
        A(display(res(sec, f"dc.art{i}.M_p")))
        ext(IMG.figure(OUT_DIR / f"fig_1_2_{i}.png", 16.0,
                       f"Рис. 1{letters[i]}. Искусственные характеристики двигателя постоянного тока при {art['name']}"))
        A(para(f"**Вывод:** при {art['name']} скорость идеального холостого хода {chg(nat['n0'], art['n0'], 'об/мин')}, "
               f"ток стоянки {chg(nat['I_p'], art['I_p'], 'А', 5)}, момент стоянки {chg(nat['M_p'], art['M_p'], 'Н·м', 5)}. "
               + dc_conc[i]))
    rows = [["естественная", fmt(nat["n0"], 4), fmt(nat["I_p"], 5), fmt(nat["M_p"], 5)]]
    for art in dc["artificial"]:
        rows.append([art["name"], fmt(art["n0"], 4), fmt(art["I_p"], 5), fmt(art["M_p"], 5)])
    A(para("Сводные результаты расчёта характеристик двигателя постоянного тока:", jc="left", keep_next=True, before=120))
    A(table(["Характеристика", "$n_0$, об/мин", "$I_{п}$, А", "$M_{п}$, Н·м"], rows, [6.0, 3.2, 3.2, 3.6]))

    # ===================================================================== 2
    A(heading1("2. ЭЛЕКТРОМЕХАНИЧЕСКИЕ И МЕХАНИЧЕСКИЕ ХАРАКТЕРИСТИКИ АСИНХРОННОГО ДВИГАТЕЛЯ"))
    A(para(f"Исходные паспортные данные (Приложение 3, вариант {VARIANT}):", b=True, jc="left"))
    for line in [f"Тип двигателя: {mac['type']}",
                 f"Линейное напряжение сети $U_{{л}} = 380$ В, фазное $U_{{ф}} = 220$ В, частота $f_1 = 50$ Гц",
                 f"Мощность: $P_{{ном}} = {fmt(mac['P_kW'], 4)}$ кВт",
                 f"Частота вращения двигателя: $n_{{ном}} = {fmt(mac['n_nom'], 4)}$ об/мин",
                 f"Статор: $r_1 = {fmt(mac['r1'], 3)}$ Ом; $x_1 = {fmt(mac['x1'], 3)}$ Ом",
                 f"Ротор: $r_2 = {fmt(mac['r2'], 3)}$ Ом; $x_2 = {fmt(mac['x2'], 3)}$ Ом",
                 f"Линейная ЭДС между кольцами ротора: $E_{{2к}} = {fmt(mac['E2k'], 3)}$ В",
                 f"Момент инерции: $J = {fmt(mac['J'], 3)}$ кг·м²"]:
        A(para(line, jc="left"))

    A(heading2("2.1 Расчет и построение естественных характеристик двигателя"))
    A(para("Естественная механическая характеристика может быть приближённо построена по номинальным данным с помощью "
           "формулы Клосса:"))
    A(display(r"M = \frac{2 M_{max}(1 + a s_{к})}{\frac{s}{s_{к}} + \frac{s_{к}}{s} + 2 a s_{к}}"))
    A(para("где $M$, $M_{max}$ – соответственно текущий и максимальный моменты двигателя, Н·м; $s$, $s_{к}$ – текущее и "
           "критическое скольжение; $a = r_1 / r_2'$ – постоянная двигателя. Значения $M_{max}$ и $s_{к}$ вычисляются по "
           "формулам:"))
    A(display(r"M_{max} = \frac{m U_{ф}^2}{2\omega_0 \left(r_1 \pm \sqrt{r_1^2 + x_{кз}^2}\right)},\qquad "
              r"s_{к} = \pm\frac{r_2'}{\sqrt{r_1^2 + x_{кз}^2}}"))
    A(para("где $m = 3$ – число фаз; $U_{ф}$ – фазное напряжение, В; $\\omega_0$ – угловая скорость вращения магнитного поля, "
           "рад/с; $r_1$ – активное сопротивление фазы статора, Ом; $r_2'$ – активное приведённое к статору сопротивление фазы "
           "ротора, Ом; $x_{кз} = x_1 + x_2'$ – индуктивное сопротивление короткого замыкания, Ом. Знак «+» – двигательный "
           "режим, знак «−» – генераторный."))
    A(para("Угловая скорость поля $\\omega_0 = 2\\pi n_0 / 60 = 2\\pi f_1 / p_{п}$, где $f_1 = 50$ Гц – частота сети; "
           "$p_{п}$ – число пар полюсов. Если $p_{п}$ не задано, его определяют как целое число, полученное делением $60 f_1$ "
           "на ближайшую большую по сравнению с номинальной частотой вращения синхронную скорость $n_0$ из ряда "
           "3000/$p_{п}$ (3000, 1500, 1000, 750, … об/мин). Для $n_{ном} = " + fmt(mac["n_nom"], 4) + "$ об/мин это $n_0 = " + fmt(ac["natural"]["n0"], 4) + "$ об/мин:"))
    A(display(res("2.1", "ac.p")))
    A(para("Приведение параметров ротора к статору выполняется через коэффициент трансформации $k = 0{,}95 U_{ф} / E_{2ф}$, "
           "где $E_{2ф} = E_{2к} / \\sqrt{3}$ – фазная ЭДС неподвижного ротора (в таблице задана линейная ЭДС между кольцами):"
           .replace("{,}", ",")))
    A(display(res("2.1", "ac.E2f")))
    A(display(res("2.1", "ac.k")))
    A(display(res("2.1", "ac.x2p")))
    A(display(res("2.1", "ac.r2p")))
    A(display(res("2.1", "ac.xkz")))
    A(para("Скорость идеального холостого хода и угловая скорость поля:"))
    A(display(res("2.1", "n0")))
    A(display(res("2.1", "w0")))
    A(para("Вспомогательная величина и постоянная двигателя:"))
    A(display(res("2.1", "Z")))
    A(display(res("2.1", "a")))
    A(para("Максимальные моменты в двигательном и генераторном режимах и критическое скольжение:"))
    A(display(res("2.1", "Mmax_d")))
    A(display(res("2.1", "Mmax_g")))
    A(display(res("2.1", "sk")))
    A(para("Пусковой момент (при $s = 1$):"))
    A(display(res("2.1", "M_start")))
    A(para("Для справки: номинальное скольжение, номинальный момент на валу и перегрузочная способность:"))
    A(display(res("2.1", "ac.s_nom")))
    A(display(res("2.1", "ac.M_nom")))
    nat_ac = ac["natural"]
    A(display(rf"\lambda_{{M}} = \frac{{M_{{max\,д}}}}{{M_{{ном}}}} = \frac{{{fmt(nat_ac['Mmax_d'], 5)}}}{{{fmt(ac['M_nom_shaft'], 5)}}} = {fmt(nat_ac['Mmax_d'] / ac['M_nom_shaft'], 3)}"))
    A(para("Далее задаёмся рядом значений скольжения $s$ в диапазоне от генераторного режима ($s < 0$) до двигательного "
           "($s > 0$), по формуле $n = n_0 (1 - s)$ определяем частоту вращения и по формуле Клосса – соответствующие моменты "
           "(для $s > 0$ – с $M_{max\\,д}$ и $s_{к}$, для $s < 0$ – с $M_{max\\,г}$ и $-s_{к}$). Искомая характеристика "
           "приведена на рис. 2."))
    ext(IMG.figure(OUT_DIR / "fig_2_1.png", 12.5, "Рис. 2. Естественная механическая характеристика асинхронного двигателя"))
    A(para(f"**Вывод:** по естественной характеристике определены критические моменты: в двигательном режиме "
           f"$M_{{max\\,д}} = {fmt(nat_ac['Mmax_d'], 5)}$ Н·м, в генераторном $M_{{max\\,г}} = {fmt(nat_ac['Mmax_g'], 5)}$ Н·м "
           f"(генераторный режим по модулю больше из-за влияния сопротивления статора), критическое скольжение "
           f"$s_{{к}} = {fmt(nat_ac['sk_d'], 4)}$ и пусковой момент $M_{{п}} = {fmt(nat_ac['M_start'], 4)}$ Н·м. "
           f"Двигатель обладает перегрузочной способностью {fmt(nat_ac['Mmax_d'] / ac['M_nom_shaft'], 3)}."))

    A(heading2("2.2 Расчет и построение искусственных характеристик двигателя"))
    A(para("Порядок расчета искусственных механических характеристик аналогичен расчету естественных: на первом этапе "
           "вычисляются критическое скольжение и максимальные моменты для заданных условий, а затем по формуле Клосса – "
           "значения текущих моментов при выбранных скольжениях. По полученным данным в осях $n(s)$ и $M$ строятся "
           "искусственные механические характеристики; на каждом рисунке для сравнения приведена естественная "
           "характеристика. " + ac_conditions_text() + ""))
    ac_expl = {
        "U": ("Момент асинхронного двигателя пропорционален квадрату напряжения, поэтому критические и пусковой моменты "
              "уменьшились пропорционально $U^2$; критическое скольжение и скорость идеального холостого хода от напряжения "
              "не зависят."),
        "r1": ("Увеличение активного сопротивления статора слегка уменьшило критическое скольжение и критический момент в "
               "двигательном режиме, а в генераторном режиме критический момент по модулю вырос."),
        "xkz": ("Индуктивное сопротивление короткого замыкания входит в знаменатели обеих формул, поэтому его увеличение "
                "резко уменьшило критические моменты и критическое скольжение."),
        "f1": ("Рост частоты сети увеличил скорость идеального холостого хода пропорционально $f_1$ и уменьшил критические и "
               "пусковой моменты (момент обратно пропорционален $\\omega_0$ при неизменных сопротивлениях)."),
        "p": ("Увеличение числа пар полюсов уменьшило скорость идеального холостого хода пропорционально росту $p_{п}$ и во "
              "столько же раз увеличило моменты (при той же мощности момент обратно пропорционален скорости поля)."),
    }
    sum_rows = [["естественная", fmt(nat_ac["Mmax_d"], 5), fmt(abs(nat_ac["Mmax_g"]), 5), fmt(nat_ac["sk_d"], 4),
                 fmt(nat_ac["n0"], 4), fmt(nat_ac["M_start"], 4)]]
    for i, art in enumerate(ac["artificial"], 1):
        sec = f"2.2 АД: искусственная характеристика {i}"
        A(para(f"**{i}) {art['name']}**", jc="left", keep_next=True, before=120))
        A(display(res(sec, f"ac.art{i}.param")))
        if art["key"] in ("f1", "p"):
            A(display(res(sec, "n0")))
            A(display(res(sec, "w0")))
        A(display(res(sec, "Z")))
        if art["key"] == "r1":
            A(display(res(sec, "a")))
        A(display(res(sec, "Mmax_d")))
        A(display(res(sec, "Mmax_g")))
        A(display(res(sec, "sk")))
        A(display(res(sec, "M_start")))
        ext(IMG.figure(OUT_DIR / f"fig_2_2_{i}.png", 12.5,
                       f"Рис. 2{'абвгд'[i - 1]}. Искусственные механические характеристики асинхронного двигателя при {art['name']}"))
        A(para(f"**Вывод:** при {art['name']} критический момент в двигательном режиме {chg(nat_ac['Mmax_d'], art['Mmax_d'], 'Н·м', 5)}, "
               f"в генераторном режиме (по модулю) {chg(nat_ac['Mmax_g'], art['Mmax_g'], 'Н·м', 5, True)}; "
               f"критическое скольжение {chg(nat_ac['sk_d'], art['sk_d'], '', 4)}; скорость идеального холостого хода "
               f"{chg(nat_ac['n0'], art['n0'], 'об/мин')}; пусковой момент {chg(nat_ac['M_start'], art['M_start'], 'Н·м', 4)}. "
               + ac_expl[art["key"]]))
        sum_rows.append([art["name"], fmt(art["Mmax_d"], 5), fmt(abs(art["Mmax_g"]), 5), fmt(art["sk_d"], 4),
                         fmt(art["n0"], 4), fmt(art["M_start"], 4)])
    A(para("Сводные результаты расчёта характеристик асинхронного двигателя:", jc="left", keep_next=True, before=120))
    A(table(["Характеристика", r"$M_{max\,д}$, Н·м", r"$|M_{max\,г}|$, Н·м", "$s_{к}$", "$n_0$, об/мин", "$M_{п}$, Н·м"], sum_rows,
            [3.6, 2.6, 2.7, 1.8, 2.5, 2.3], sz=22))

    # ===================================================================== 3
    st, tr = RES["start"], RES["transient"]
    A(heading1("3. РАСЧЕТ И ПОСТРОЕНИЕ ПУСКОВЫХ ДИАГРАММ ДЛЯ ДВИГАТЕЛЕЙ С ПРЯМОЛИНЕЙНЫМИ ХАРАКТЕРИСТИКАМИ"))
    A(para("Основной задачей расчета является определение числа и величин ступеней сопротивления пускового реостата в цепи "
           "ротора асинхронного двигателя. На практике применяются графический и аналитический способы расчета пусковых "
           "диаграмм. Типовая пусковая диаграмма для линейных механических характеристик вследствие пропорциональной "
           "зависимости $M \\sim I$ между током и моментом справедлива как для электромеханических, так и для механических "
           "характеристик."))
    A(para("При построении пусковой диаграммы следует стремиться к тому, чтобы выход на естественную характеристику "
           "осуществлялся при $M = M_{п}$ ($I = I_{п}$). Для этого изменяют $M_{п}$ или $M_{пер}$ в допустимых пределах и "
           "вновь строят диаграмму. Чем меньше разность между $M_{п}$ и $M_{пер}$ при том же $M_{п}$, тем больше ступеней "
           "имеет реостат, пуск плавнее и короче за счёт увеличения среднего избыточного момента; однако большое число "
           "ступеней усложняет и удорожает схему."))
    A(para("Принятые условия: момент сопротивления равен номинальному моменту двигателя $M_{с} = M_{ном}$; пусковой момент "
           f"$M_{{п}} = {fmt(K_MP, 3)} M_{{ном}}$ (меньше $M_{{max}}$); желаемый момент переключения "
           f"$M_{{пер}}' = {fmt(K_MPER, 3)} M_{{ном}}$ (из диапазона $1{{,}}05 \\div 1{{,}}3$). Участок характеристик до $M_{{п}}$ "
           "линеен, так как $s < s_{к}$.".replace("{,}", ",")))
    A(heading2("3.1 Аналитический способ"))
    A(para("Полное приведённое сопротивление цепи фазы ротора при пуске следует из пусковой диаграммы:"))
    A(display(formula("3.", "st.Rp")))
    A(display(res("3.", "st.Mc")))
    A(display(res("3.", "st.Mp")))
    A(display(res("3.", "st.Mper_t")))
    A(display(res("3.", "st.Rp")))
    A(para("Число ступеней сопротивления вычисляется по формуле $n = \\ln(R_{п}/r_2') / \\ln\\lambda$, где "
           "$\\lambda = M_{п} / M_{пер}$:"))
    A(display(res("3.", "st.lam0")))
    A(display(res("3.", "st.n_calc")))
    A(para("Число ступеней округляем до целого в большую сторону, после чего уточняем $\\lambda$ так, чтобы выход на "
           "естественную характеристику происходил при $M = M_{п}$, и находим момент переключения:"))
    A(display(res("3.", "st.n")))
    A(display(res("3.", "st.lam")))
    A(display(res("3.", "st.Mper")))
    A(para(f"Момент переключения $M_{{пер}} = {fmt(st['M_per'], 5)}$ Н·м удовлетворяет условиям $M_{{с}} < M_{{пер}} < M_{{п}}$ "
           f"и $(1{{,}}05 \\div 1{{,}}3) M_{{ном}}$.".replace("{,}", ",")))
    A(para("Значения сопротивлений отдельных ступеней пускового резистора (приведённые к статору) определяются по формуле "
           "$R_{дi} = r_2' \\lambda^{i-1} (\\lambda - 1)$:"))
    for i in range(1, st["n_st"] + 1):
        A(display(res("3.", f"st.Rd{i}")))
    A(display(res("3.", "st.sumRd")))
    A(para("Реальные сопротивления секций реостата в цепи ротора получаются делением приведённых значений на $k^2$:"))
    for i in range(1, st["n_st"] + 1):
        A(display(res("3.", f"st.Rdr{i}")))
    A(heading2("3.2 Графический способ"))
    A(para("При графическом способе расчета масштаб сопротивления $m_2$ определяется при пусковой нагрузке по рис. 3. "
           "Диаграмма построена в масштабе: скорость $n_0$ соответствует отрезку 200 мм; длины отрезков на вертикали $M = M_{п}$ "
           "измерены с точностью 0,5 мм. Масштаб сопротивления и сопротивления ступеней:"))
    A(display(res("3.", "st.m2")))
    A(display(r"R_{дi} = m_2 l_i"))
    for i in range(1, st["n_st"] + 1):
        A(display(res("3.", f"st.Rdg{i}")))
    ext(IMG.figure(OUT_DIR / "fig_3_1.png", 13.5, "Рис. 3. Пусковая диаграмма для линейных механических характеристик двигателя"))
    rows = [[i, fmt(st["Rd"][i - 1], 4), fmt(st["Rd_graph"][i - 1], 4), fmt(st["dev_pct"][i - 1], 3)]
            for i in range(1, st["n_st"] + 1)]
    A(para("Сравнение результатов аналитического и графического способов (допустимое расхождение 5–10 %):", jc="left",
           keep_next=True, before=120))
    A(table(["Ступень", "$R_{д}$ аналитически, Ом", "$R_{д}$ графически, Ом", "Расхождение, %"], rows, [2.4, 4.6, 4.6, 3.8]))
    A(para(f"**Вывод:** для пуска асинхронного двигателя с заданным пусковым моментом $M_{{п}} = {fmt(st['M_p'], 5)}$ Н·м и "
           f"моментом переключения $M_{{пер}} = {fmt(st['M_per'], 5)}$ Н·м требуется {st['n_st']} {steps_word(st['n_st'])} пускового реостата "
           f"(полное сопротивление при пуске $R_{{п}} = {fmt(st['R_p'], 4)}$ Ом). Результаты графического и аналитического "
           f"способов не отличаются более чем на {fmt(max(st['dev_pct']), 2)} %, что меньше допустимых 10 %, следовательно, "
           f"расчёт выполнен верно.", before=120))

    # ===================================================================== 4
    A(heading1("4. РАСЧЕТ И ПОСТРОЕНИЕ ГРАФИКОВ МЕХАНИЧЕСКОГО ПЕРЕХОДНОГО РЕЖИМА АСИНХРОННОГО ДВИГАТЕЛЯ ПРИ СТУПЕНЧАТОМ ПУСКЕ"))
    A(para("Анализ переходных режимов предусматривает определение продолжительности, а также характера изменения основных "
           "параметров (частоты вращения, момента на валу) во времени. Анализу подлежит механический переходный режим при "
           "прямолинейно убывающем избыточном моменте в случае ступенчатого реостатного пуска. Частоту вращения и момент на "
           "валу на $i$-й ступени находят по формулам:"))
    A(display(r"n_i = n_{уi}\left(1 - e^{-\frac{t}{T_{Мi}}}\right) + n_{начi} e^{-\frac{t}{T_{Мi}}},\qquad "
              r"M_i = M_{уi}\left(1 - e^{-\frac{t}{T_{Мi}}}\right) + M_{начi} e^{-\frac{t}{T_{Мi}}}"))
    A(para("где $n_{уi}$ – установившаяся частота вращения на $i$-й ступени при $M = M_{с}$, об/мин; $T_{Мi}$ – механическая "
           "постоянная времени на $i$-й ступени, с; $t$ – время, с. Механическая постоянная времени и время разгона на ступени:"))
    A(display(r"T_{Мi} = \frac{J n_{уi}}{9{,}55 M_{икрi}} = \frac{J (n_{конi} - n_{начi})}{9{,}55 (M_{иначi} - M_{иконi})},\qquad "
              r"t_i = T_{Мi} \ln\frac{M_{иначi}}{M_{иконi}}".replace("{,}", ",")))
    A(para("где $J$ – момент инерции системы, кг·м²; $M_{иначi}$, $M_{иконi}$ – избыточные моменты в начале и в конце разгона на "
           "$i$-й ступени, Н·м; $M_{икрi}$ – критический избыточный момент на $i$-й ступени (момент при $n = 0$ минус $M_{с}$). "
           "Момент инерции принят равным моменту инерции двигателя $J = " + fmt(mac["J"], 3) + "$ кг·м² (сведений о "
           "моменте инерции нагрузки в задании нет)."))
    A(para("Для всех ступеней одинаковы:"))
    A(display(res("4.", "tr.My")))
    A(display(res("4.", "tr.Min")))
    A(display(res("4.", "tr.Mik")))
    A(display(res("4.", "tr.ln")))
    A(para("Число характеристик на 1 больше числа ступеней реостата: $m + 1 = " + str(st["n_st"] + 1) + "$ (последняя – "
           "естественная). Для каждой характеристики определяются диапазон изменения частоты вращения и постоянная времени. "
           "Расчёт для первой ступени:"))
    for key in ("ny1", "Mikr1", "nk1", "T1", "t1"):
        A(display(res("4.", "tr." + key)))
    A(para("Для остальных ступеней аналогично находим $T_{Мi}$ и $t_i$:"))
    for i in range(2, st["n_st"] + 2):
        A(display(res("4.", f"tr.T{i}")))
        A(display(res("4.", f"tr.t{i}")))
    A(para("Результаты вычислений по ступеням сведены в таблицу:", jc="left", keep_next=True, before=120))
    rows = []
    for s in tr["stages"]:
        rows.append([s["i"], fmt(s["n_start"], 4), fmt(s["n_end"], 4), fmt(s["n_y"], 4), fmt(s["T"], 4), fmt(s["t"], 4)])
    A(table(["Ступень", "$n_{нач}$, об/мин", "$n_{кон}$, об/мин", "$n_{уст}$, об/мин", "$T_{М}$, с", "$t$, с"], rows,
            [2.0, 3.0, 3.0, 3.0, 2.6, 2.6]))
    A(para("Полное время разгона, суммированное по ступеням (по графику переходного процесса), и аналитическое значение по "
           "общей формуле $t = \\dfrac{J n_{ном}}{9{,}55 (M_{инач} - M_{икон})} \\ln\\dfrac{M_{инач}}{M_{икон}}$:"
           .replace("\\dfrac", "\\frac").replace("{,}", ","), before=120))
    A(display(res("4.", "tr.tsum")))
    A(display(res("4.", "tr.tan")))
    A(display(res("4.", "tr.dev")))
    A(para("Графики изменения частоты вращения и момента во времени приведены на рис. 4 (пусковая диаграмма – на рис. 3)."))
    ext(IMG.figure(OUT_DIR / "fig_4_1.png", 14.5,
                   "Рис. 4. График механического переходного процесса при ступенчатом пуске: n = f(t) и M = f(t)"))
    A(para(f"**Вывод:** выполнены расчёты и построены графики механического переходного процесса. Разгон двигателя по "
           f"{st['n_st']} ступеням реостата и естественной характеристике занимает {fmt(tr['t_sum'], 4)} с; значение, "
           f"рассчитанное по общей формуле, – {fmt(tr['t_an'], 4)} с (расхождение {fmt(tr['dev_pct'], 2)} %, допустимо "
           f"до 10 %). Момент на каждой ступени снижается от $M_{{п}} = {fmt(st['M_p'], 5)}$ Н·м до "
           f"$M_{{пер}} = {fmt(st['M_per'], 5)}$ Н·м, после чего очередная секция реостата закорачивается."))

    # ===================================================================== 5
    d5 = RES["dc_transient"]
    A(heading1("5. РАСЧЕТ И ПОСТРОЕНИЕ ГРАФИКОВ ЭЛЕКТРОМЕХАНИЧЕСКОГО ПЕРЕХОДНОГО РЕЖИМА ДВИГАТЕЛЯ ПОСТОЯННОГО ТОКА ПРИ РАЗГОНЕ ДВИГАТЕЛЯ ПО ЕСТЕСТВЕННОЙ ХАРАКТЕРИСТИКЕ"))
    A(para("Анализ электромеханических переходных режимов из-за необходимости учёта механической и электромагнитной "
           "инерций значительно труднее рассмотренных выше механических режимов. Для электропривода постоянного тока при "
           "условии постоянства потока и индуктивности обмоток, а также при пренебрежении потерями на вихревые токи и "
           "гистерезис, характер переходного процесса можно установить аналитически. При разгоне из неподвижного состояния "
           "процесс состоит из двух частей. В первой части (двигатель неподвижен) происходит чисто электромагнитный "
           "переходный процесс изменения тока и момента до тех пор, пока электромагнитный момент не превысит момент "
           "сопротивления $M_{с}$. В этот период:"))
    A(display(r"I = \frac{U}{R}\left(1 - e^{-\frac{t}{T_{э}}}\right),\qquad M = I k\Phi_{ном}"))
    A(para("Во второй части (с началом вращения) переходный процесс электромеханический:"))
    A(display(r"n = A_1 e^{p_1 t} + A_2 e^{p_2 t} + n_{ном},\quad I = B_1 e^{p_1 t} + B_2 e^{p_2 t} + I_{ном},\quad "
              r"M = C_1 e^{p_1 t} + C_2 e^{p_2 t} + M_{ном}"))
    A(para("Корни характеристического уравнения:"))
    A(display(r"p_{1,2} = -\frac{1}{2 T_{э}} \pm \sqrt{\frac{1}{4 T_{э}^2} - \frac{1}{T_{э} T_{М}}}"))
    A(para("где $T_{М} = J n_{у} / (9{,}55 M_{ик})$ – механическая постоянная времени, $T_{э} = L / R$ – электромагнитная "
           "постоянная времени якоря. Индуктивность цепи якоря приближённо:".replace("{,}", ",")))
    A(display(r"L = \beta\frac{U}{p_{п}\omega_{ном} I_{ном}} = 9{,}55\,\beta\frac{U}{p_{п} n_{ном} I_{ном}}".replace("{,}", ",")))
    A(para(f"где $\\beta = 0{{,}}6 \\ldots 0{{,}}25$ (нижний предел – для компенсированных машин; принято $\\beta = {fmt(BETA_DC, 3)}$ "
           f"как в образце); $p_{{п}} = 2$ – число пар полюсов (в паспортных данных не указано, принято по типу машины). "
           f"Момент сопротивления принят равным номинальному: $M_{{с}} = M_{{ном}}$, $n_{{с}} = n_{{ном}}$."
           .replace("{,}", ",")))
    A(display(res("5.", "tr5.Mc")))
    A(display(res("5.", "tr5.nc")))
    A(display(res("5.", "tr5.L")))
    A(display(res("5.", "tr5.Te")))
    A(para("Механическая постоянная времени (для естественной характеристики $n_{у} = n_{ном}$, $M_{ик} = M_{п} - M_{с}$):"))
    A(display(res("5.", "tr5.Tm")))
    A(para("Длительность первой части процесса (якорь неподвижен, ток растёт до номинального $I_{ном}$, при котором "
           "$M = M_{с}$):"))
    A(display(res("5.", "tr5.dt")))
    A(para("Дискриминант характеристического уравнения:"))
    A(display(res("5.", "tr5.D")))
    A(para("Так как $D < 0$, корни комплексные, процесс колебательный: $p_{1,2} = -\\delta \\pm j\\nu$. Расчёт выполняется по "
           "формулам, полученным с помощью преобразований Эйлера:"))
    A(display(r"n = A e^{-\frac{t}{2T_{э}}}\sin(\nu t + \varphi) + n_{ном},\quad "
              r"I = B e^{-\frac{t}{2T_{э}}}\sin(\nu t + \psi) + I_{ном},\quad "
              r"M = C e^{-\frac{t}{2T_{э}}}\sin(\nu t + \gamma) + M_{ном}"))
    A(display(res("5.", "tr5.alpha")))
    A(display(res("5.", "tr5.nu")))
    A(display(res("5.", "tr5.A")))
    A(para("Начальная фаза по формуле методики $\\varphi = \\arcsin(-\\nu\\sqrt{T_{э} T_{М}})$ даёт главное значение арксинуса. "
           "Условию трогания якоря с нулевой скоростью ($n(0) = 0$ и $n'(0) = 0$) удовлетворяет второй корень уравнения "
           "$\\sin\\varphi = -\\nu\\sqrt{T_{э} T_{М}}$, поэтому принимаем:"))
    A(display(res("5.", "tr5.phi0")))
    A(display(res("5.", "tr5.phi")))
    A(display(res("5.", "tr5.B")))
    A(display(res("5.", "tr5.C")))
    A(display(r"\psi = 0;\qquad \gamma = 0"))
    al, nu, Aa, ph, Bb, Cc = d5["alpha"], d5["nu"], d5["A"], d5["phi"], d5["B"], d5["C"]
    A(para("Окончательно, при отсчёте времени от начала вращения:"))
    A(display(rf"n = {fmt(Aa, 4)}\, e^{{-{fmt(al, 4)} t}} \sin({fmt(nu, 4)} t - {fmt(abs(ph), 4)}) + {fmt(d5['n_nom'], 4)}"))
    A(display(rf"I = {fmt(Bb, 4)}\, e^{{-{fmt(al, 4)} t}} \sin({fmt(nu, 4)} t) + {fmt(d5['I_nom'], 4)}"))
    A(display(rf"M = {fmt(Cc, 4)}\, e^{{-{fmt(al, 4)} t}} \sin({fmt(nu, 4)} t) + {fmt(d5['M_nom'], 5)}"))
    A(para("Характерные значения переходного процесса. Первые максимумы момента и тока достигаются при "
           "$\\operatorname{tg}(\\nu t) = 2T_{э}\\nu$, частоты вращения – при $\\operatorname{tg}(\\nu t + \\varphi) = 2T_{э}\\nu$:"
           .replace("\\operatorname{tg}", "\\mathrm{tg}")))
    A(display(res("5.", "tr5.tMmax")))
    A(display(res("5.", "tr5.Mmax")))
    A(display(res("5.", "tr5.Imax")))
    A(display(res("5.", "tr5.nmax")))
    A(display(res("5.", "tr5.sigma")))
    A(para("Время переходного процесса принимаем равным времени затухания огибающей колебаний до 5 % амплитуды:"))
    A(display(res("5.", "tr5.tset")))
    A(para("Механическая характеристика и графики колебательного электромеханического процесса приведены на рис. 5."))
    ext(IMG.figure(OUT_DIR / "fig_5_1.png", 16.3,
                   "Рис. 5. Характеристики двигателя постоянного тока с параллельным возбуждением: а) механическая "
                   "характеристика; в) характеристика колебательного электромеханического процесса"))
    A(para(f"**Вывод:** электромеханический переходный процесс двигателя постоянного тока при пуске от сети по естественной "
           f"характеристике имеет колебательный характер (корни комплексные: $T_{{э}} = {fmt(d5['Te'], 4)}$ с значительно "
           f"больше $T_{{М}} = {fmt(d5['Tm'], 3)}$ с). Первая часть процесса (якорь неподвижен) длится "
           f"$\\Delta t = {fmt(d5['dt1'], 3)}$ с. Максимальные значения тока $I_{{max}} = {fmt(d5['I_max'], 5)}$ А и момента "
           f"$M_{{max}} = {fmt(d5['M_max'], 5)}$ Н·м; частота вращения кратковременно превышает скорость идеального холостого "
           f"хода и достигает {fmt(d5['n_max'], 4)} об/мин (перерегулирование {fmt(d5['sigma_pct'], 3)} %). Процесс "
           f"затухает за время около {fmt(d5['t_set'], 3)} с."))

    # ===================================================================== литература
    A(heading1("Список литературы"))
    A(para("1. Бова Е.В., Гельвер Ф.А., Саушев А.В. Электрический привод: учебно-методическое пособие. — СПб.: Изд-во ГУМРФ "
           "им. адм. С. О. Макарова, 2019. 25 с."))
    return body


# ----------------------------------------------------------------------------- сборка пакета
def main():
    z = zipfile.ZipFile(TEMPLATE)
    doc = etree.fromstring(z.read("word/document.xml"))
    bodyel = doc.find(w("body"))
    children = list(bodyel)
    first_sect = children[240].find(f".//{w('sectPr')}")      # портретный раздел с titlePg (титул без номера)
    keep = children[:27]                                      # титул + оглавление (sdt) + служебный абзац
    for t in keep[12].iter(w('t')):
        t.text = f'Вариант {VARIANT} '
    for c in children:
        bodyel.remove(c)
    for c in keep:
        bodyel.append(c)
    bodyel.append(etree.fromstring(f'<w:p xmlns:w="{W}"><w:pPr><w:pStyle w:val="a0"/></w:pPr></w:p>'))
    for el in build_body():
        bodyel.append(el)
    bodyel.append(first_sect)

    # связи и типы содержимого для рисунков
    rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
    for rid, target in IMG.rels:
        e = etree.SubElement(rels, f"{{{REL_PKG}}}Relationship")
        e.set("Id", rid)
        e.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        e.set("Target", target)
    settings = z.read("word/settings.xml").decode("utf8")
    if "w:updateFields" not in settings:
        settings = settings.replace("<w:defaultTabStop", '<w:updateFields w:val="true"/><w:defaultTabStop', 1)

    RESULT_DOCX.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(RESULT_DOCX, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in z.infolist():
            data = z.read(info.filename)
            if info.filename == "word/document.xml":
                data = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif info.filename == "word/_rels/document.xml.rels":
                data = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif info.filename == "word/settings.xml":
                data = settings.encode("utf8")
            zout.writestr(info, data)
        for name, path in IMG.files.items():
            zout.write(path, name)
    print("Собрано:", RESULT_DOCX)
    update_fields_with_word()


def _retry(fn, tries=30, pause=1.0):
    """Word занят пересчётом страниц — COM отвечает «Call was rejected»; повторяем вызов."""
    import time
    import pywintypes
    for k in range(tries):
        try:
            return fn()
        except pywintypes.com_error:
            time.sleep(pause)
    return fn()


def update_fields_with_word():
    """Обновляет оглавление и номера страниц через Word (если установлен) и сохраняет файл."""
    try:
        import win32com.client
    except ImportError:
        print("pywin32 не установлен: оглавление обновится при открытии файла в Word (согласитесь на обновление полей).")
        return
    w_app = win32com.client.Dispatch("Word.Application")
    w_app.Visible = False
    try:
        d = _retry(lambda: w_app.Documents.Open(str(RESULT_DOCX)))
        _retry(lambda: d.Repaginate())
        _retry(lambda: d.Fields.Update())
        for k in range(1, _retry(lambda: d.TablesOfContents.Count) + 1):
            _retry(lambda: d.TablesOfContents(k).Update())
        _retry(lambda: d.Save())
        _retry(lambda: d.Close())
        print("Оглавление обновлено через Word.")
    finally:
        _retry(lambda: w_app.Quit())


if __name__ == "__main__":
    main()
