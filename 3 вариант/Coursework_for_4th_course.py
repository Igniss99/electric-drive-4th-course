"""Курсовая работа «Электрический привод» (ГУМРФ). Разделы 1–2: характеристики ДПТ и АД.

Запуск:  python Coursework_for_4th_course.py
Выход (out/variant_<N>/):
    results.json   — итоговые значения
    calc_db.md     — база расчётов: формулы LaTeX (для Word) с ID, подстановки, результаты
    calc_db.json   — то же в машинном виде (ID -> формула, значение, единицы)
    fig_*.png      — графики
В терминале печатается подробный лог всех расчётов.
Нумерация пунктов и формулы — по методичке (Metodichka.docx).
"""
import json
import math
import re
import sys
from dataclasses import dataclass
from pathlib import Path

import matplotlib
import matplotlib.ticker

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

VARIANT = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else 3
OUT_DIR = Path(__file__).parent / "out" / f"variant_{VARIANT}"

# ---------------------------------------------------------------- константы
F1 = 50.0          # частота сети, Гц
U_LINE = 380.0     # линейное напряжение АД, В (Прил. 3)
U_PHASE = 220.0    # фазное напряжение, В (принято как в образце: 380/√3 ≈ 220)
M_PHASES = 3       # число фаз
K_SAFETY = 0.95    # коэффициент 0,95 в k = 0,95·Uф/E2ф (п. 2.1)
SYNC_NUM = 60 * F1  # 60·f1 = 3000; синхронная скорость n0 = 3000/pп


def rpm_to_rad(n):
    return 2 * math.pi * n / 60


def rad_to_rpm(w):
    return 60 * w / (2 * math.pi)


# ---------------------------------------------------------------- лог и база расчётов
def fmt(x, nd=4):
    """До nd значащих цифр, хвостовые нули отбрасываются, десятичная запятая (для Word и терминала)."""
    x = float(x)
    if x == 0:
        return "0"
    digits = max(nd - 1 - math.floor(math.log10(abs(x))), 0)
    out = f"{x:.{digits}f}"
    if "." in out:
        out = out.rstrip("0").rstrip(".")   # 220,0 -> 220; 0,01400 -> 0,014
    return out.replace(".", ",")


def _group(s, i):
    """Содержимое {...}, открывающаяся скобка в s[i]; возвращает (текст, индекс после закрывающей)."""
    depth = 0
    for j in range(i, len(s)):
        if s[j] == "{":
            depth += 1
        elif s[j] == "}":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
    return s[i + 1:], len(s)


def tex2txt(s):
    """LaTeX -> читаемый текст для терминала (frac, sqrt, text, индексы, греческие буквы)."""
    s = s.replace("\\left", "").replace("\\right", "")
    out, i = [], 0
    while i < len(s):
        if s.startswith("\\frac", i):
            num, i = _group(s, i + 5)
            den, i = _group(s, i)
            out.append(f"({tex2txt(num)})/({tex2txt(den)})")
        elif s.startswith("\\sqrt", i):
            g, i = _group(s, i + 5)
            out.append(f"√({tex2txt(g)})")
        elif s.startswith("\\text", i):
            g, i = _group(s, i + 5)
            out.append(g)
        elif s[i] in "_^" and i + 1 < len(s) and s[i + 1] == "{":
            g, i2 = _group(s, i + 1)
            out.append(s[i] + tex2txt(g))
            i = i2
        elif s[i] == "{":
            g, i = _group(s, i)
            out.append(tex2txt(g))
        else:
            out.append(s[i])
            i += 1
    t = "".join(out)
    for a, b in [("\\cdot", "·"), ("\\pi", "π"), ("\\omega", "ω"), ("\\Phi", "Φ"), ("\\pm", "±"),
                 ("\\approx", "≈"), ("\\ ", " "), ("\\,", " "), ("^2", "²")]:
        t = t.replace(a, b)
    return t




class CalcLog:
    """Каждый шаг: [ID_ДАННЫЕ_n] | [ID_ФОРМУЛА_n] + [ID_РЕЗУЛЬТАТ_n] (n общий для пары)."""

    def __init__(self):
        self.items = []
        self.n_step = 0
        self.n_data = 0
        self.section_name = ""

    def section(self, name):
        self.section_name = name
        self.items.append(dict(kind="РАЗДЕЛ", section=name))
        print(f"\n{'=' * 78}\n{name}\n{'=' * 78}")

    def data(self, key, title, lhs, value, unit, nd=4, note=""):
        self.n_data += 1
        iid = f"[ID_ДАННЫЕ_{self.n_data}]"
        tex = f"{lhs} = {fmt(value, nd)}" + (rf"\ \text{{{unit}}}" if unit else "")
        self.items.append(dict(kind="ДАННЫЕ", id=iid, key=key, section=self.section_name,
                               title=title, latex=tex, value=float(value), unit=unit, note=note))
        print(f"{iid} {title}: {tex2txt(tex)}" + (f"   ({note})" if note else ""))
        return value

    def step(self, key, title, lhs, rhs_sym, rhs_num, value, unit, nd=4, note="", ref="п."):
        self.n_step += 1
        n = self.n_step
        fid, rid = f"[ID_ФОРМУЛА_{n}]", f"[ID_РЕЗУЛЬТАТ_{n}]"
        tex_f = f"{lhs} = {rhs_sym}"
        u = (r"\ \%" if unit == "%" else rf"\ \text{{{unit}}}") if unit else ""
        tex_r = f"{lhs} = {rhs_sym} = {rhs_num} = {fmt(value, nd)}{u}"
        self.items.append(dict(kind="ФОРМУЛА", id=fid, key=key, section=self.section_name,
                               title=title, latex=tex_f, ref=ref, note=note))
        self.items.append(dict(kind="РЕЗУЛЬТАТ", id=rid, key=key, section=self.section_name,
                               title=title, latex=tex_r, value=float(value), unit=unit, ref=ref))
        print(f"{fid} {title}  [{ref}]")
        print(f"      {tex2txt(tex_f)}")
        print(f"      {rid}  {tex2txt(lhs)} = {tex2txt(rhs_num)} = {fmt(value, nd)} {unit}")
        if note:
            print(f"      примечание: {note}")
        return value

    def text(self, title, tex):
        self.n_step += 1
        n = self.n_step
        fid = f"[ID_ФОРМУЛА_{n}]"
        self.items.append(dict(kind="ФОРМУЛА", id=fid, key="text", section=self.section_name,
                               title=title, latex=tex, ref=""))
        print(f"{fid} {title}\n      {tex2txt(tex)}")

    def table(self, title, headers, rows):
        self.items.append(dict(kind="ТАБЛИЦА", section=self.section_name, title=title, headers=headers, rows=rows))
        print(f"{title}")
        widths = [max(len(str(x)) for x in [h] + [r[j] for r in rows]) for j, h in enumerate(headers)]
        print("      " + " | ".join(str(h).ljust(w) for h, w in zip(headers, widths)))
        for r in rows:
            print("      " + " | ".join(str(x).replace(".", ",").ljust(w) for x, w in zip(r, widths)))

    def save(self, outdir):
        (outdir / "calc_db.json").write_text(
            json.dumps(self.items, ensure_ascii=False, indent=2), encoding="utf-8")
        lines = ["% База расчётов (LaTeX для Word: вставка через Alt+= , затем Ctrl+= для преобразования)",
                 f"% Вариант {VARIANT}", ""]
        for it in self.items:
            if it["kind"] == "РАЗДЕЛ":
                lines += ["", "% " + "=" * 70, f"% {it['section']}", "% " + "=" * 70]
            elif it["kind"] == "ТАБЛИЦА":
                lines.append(f"% {it['title']}")
                lines.append("% " + "\t".join(map(str, it["headers"])))
                lines += ["% " + "\t".join(str(x).replace(".", ",") for x in r) for r in it["rows"]]
                lines.append("")
            else:
                lines.append(f"% {it['id']} {it['title']}" + (f" [{it['ref']}]" if it.get("ref") else ""))
                lines.append(it["latex"])
                lines.append("")
        body = "\n".join(lines)
        (outdir / "calc_db.md").write_text(
            "# База расчётов\n\n```latex\n" + body + "\n```\n", encoding="utf-8")


log = CalcLog()


# ---------------------------------------------------------------- исходные данные
@dataclass
class DCMotor:            # Прил. 1
    type: str
    P_kW: float
    n_nom: float          # об/мин
    I_nom: float          # А
    R_ya: float           # Ом
    J: float              # кг·м²
    U_nom: float          # В


@dataclass
class ACMotor:            # Прил. 3
    type: str
    P_kW: float
    n_nom: float          # об/мин
    r1: float
    x1: float
    r2: float
    x2: float
    E2k: float            # линейная ЭДС между кольцами ротора, В
    J: float


from variants_data import AC_GROUPS, AC_TABLE, DC_GROUPS, DC_TABLE  # noqa: E402

DC_MOTORS = {r[0]: DCMotor(r[1], r[2], r[3], r[4], r[5], r[6], r[7]) for r in DC_TABLE}
AC_MOTORS = {r[0]: ACMotor(r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9]) for r in AC_TABLE}


def _k(x):
    return fmt(x, 3)


def dc_art_conditions(variant):
    """Приложение 2: (название, k_Ф, k_R, k_U) для характеристик 1…3 группы, в которую входит вариант."""
    kf, kr, ku = DC_GROUPS[group_of(variant)]
    return [(f"Ф = {_k(kf)}Фном", kf, 1.0, 1.0),
            (f"R = {_k(kr)}Rя", 1.0, kr, 1.0),
            (f"U = {_k(ku)}Uном", 1.0, 1.0, ku)]


def ac_art_conditions(variant):
    """Приложение 4: (название, ключ параметра, множитель) для характеристик 1…5."""
    ku, kr1, kx, kf, kp = AC_GROUPS[group_of(variant)]
    return [(f"U = {_k(ku)}Uном", "U", ku),
            (f"r1 = {_k(kr1)}r1", "r1", kr1),
            (f"xкз = {_k(kx)}xкз", "xkz", kx),
            (f"f1 = {_k(kf)}f1", "f1", kf),
            (f"pп = {_k(kp)}pп" if kp != 1 else "pп", "p", float(kp))]


def group_of(variant):
    for lo, hi in [(1, 6), (7, 12), (13, 18), (19, 24), (25, 29)]:
        if lo <= variant <= hi:
            return f"{lo}-{hi}"


# ---------------------------------------------------------------- раздел 1
def dc_natural(m: DCMotor):
    """п. 1.1: естественные характеристики ДПТ."""
    log.section("1.1 ДПТ: естественные характеристики")
    P = log.data("dc.P", "Номинальная мощность", "P_{ном}", m.P_kW * 1e3, "Вт", 5)
    n_nom = log.data("dc.n_nom", "Номинальная частота вращения", "n_{ном}", m.n_nom, "об/мин", 4)
    I_nom = log.data("dc.I_nom", "Номинальный ток якоря", "I_{ном}", m.I_nom, "А", 4)
    R = log.data("dc.R_ya", "Сопротивление якоря", "R_{я}", m.R_ya, "Ом", 3)
    U = log.data("dc.U_nom", "Номинальное напряжение", "U_{ном}", m.U_nom, "В", 3)

    w_nom = log.step("dc.w_nom", "Номинальная угловая скорость", r"\omega_{ном}",
                     r"\frac{2\pi n_{ном}}{60}", rf"\frac{{2\pi \cdot {fmt(n_nom)}}}{{60}}",
                     rpm_to_rad(n_nom), "рад/с", 5, ref="п. 1.1")
    M_nom = log.step("dc.M_nom", "Номинальный момент", "M_{ном}",
                     r"\frac{P_{ном}}{\omega_{ном}}", rf"\frac{{{fmt(P, 5)}}}{{{fmt(w_nom, 5)}}}",
                     P / w_nom, "Н·м", 5, ref="п. 1.1")
    cPhi = log.step("dc.cPhi_nom", "Коэффициент cΦном", r"c\Phi_{ном}",
                    r"\frac{U_{ном} - I_{ном} R_{я}}{n_{ном}}",
                    rf"\frac{{{fmt(U)} - {fmt(I_nom)} \cdot {fmt(R)}}}{{{fmt(n_nom)}}}",
                    (U - I_nom * R) / n_nom, "В·мин/об", 4, ref="п. 1.1")
    kPhi = log.step("dc.kPhi_nom", "Коэффициент kΦном", r"k\Phi_{ном}",
                    r"\frac{M_{ном}}{I_{ном}}", rf"\frac{{{fmt(M_nom, 5)}}}{{{fmt(I_nom)}}}",
                    M_nom / I_nom, "Н·м/А", 4, ref="п. 1.1")
    n0 = log.step("dc.n0", "Пограничная частота вращения (идеальный х.х.)", "n_{0е}",
                  r"\frac{U_{ном}}{c\Phi_{ном}}", rf"\frac{{{fmt(U)}}}{{{fmt(cPhi)}}}",
                  U / cPhi, "об/мин", 4, ref="п. 1.1")
    I_p = log.step("dc.I_p", "Ток стоянки (пусковой)", "I_{пе}",
                   r"\frac{U_{ном}}{R_{я}}", rf"\frac{{{fmt(U)}}}{{{fmt(R)}}}",
                   U / R, "А", 5, ref="п. 1.1")
    M_p = log.step("dc.M_p", "Момент стоянки (пусковой)", "M_{пе}",
                   r"\frac{U_{ном} k\Phi_{ном}}{R_{я}} = I_{пе} k\Phi_{ном}",
                   rf"{fmt(I_p, 5)} \cdot {fmt(kPhi)}", I_p * kPhi, "Н·м", 5, ref="п. 1.1")
    n_chk = log.step("dc.n_check", "Проверка: nном по cΦном", "n_{ном}",
                     r"\frac{U_{ном} - I_{ном} R_{я}}{c\Phi_{ном}}",
                     rf"\frac{{{fmt(U)} - {fmt(I_nom)} \cdot {fmt(R)}}}{{{fmt(cPhi)}}}",
                     (U - I_nom * R) / cPhi, "об/мин", 4, ref="проверка")
    assert abs(n_chk - n_nom) < 1e-6
    assert abs(M_p / M_nom - I_p / I_nom) < 1e-9
    return dict(P=P, w_nom=w_nom, M_nom=M_nom, cPhi_nom=cPhi, kPhi_nom=kPhi,
                n0=n0, I_p=I_p, M_p=M_p)


def dc_artificial(m: DCMotor, nat, tag, phi, rk, uk):
    """п. 1.2: искусственная характеристика при изменённых Ф, R, U."""
    U = log.step(f"dc.{tag}.U", "Напряжение", "U", r"k_U U_{ном}",
                 rf"{fmt(uk)} \cdot {fmt(m.U_nom)}", uk * m.U_nom, "В", 4, ref="п. 1.2")
    R = log.step(f"dc.{tag}.R", "Сопротивление цепи якоря", "R", r"k_R R_{я}",
                 rf"{fmt(rk)} \cdot {fmt(m.R_ya)}", rk * m.R_ya, "Ом", 4, ref="п. 1.2")
    n0 = log.step(f"dc.{tag}.n0", "Пограничная частота вращения", "n_0",
                  r"\frac{U}{k_\Phi c\Phi_{ном}}",
                  rf"\frac{{{fmt(U)}}}{{{fmt(phi)} \cdot {fmt(nat['cPhi_nom'])}}}",
                  U / (nat["cPhi_nom"] * phi), "об/мин", 4, ref="п. 1.2")
    I_p = log.step(f"dc.{tag}.I_p", "Ток стоянки", "I_{п}", r"\frac{U}{R}",
                   rf"\frac{{{fmt(U)}}}{{{fmt(R)}}}", U / R, "А", 5, ref="п. 1.2")
    M_p = log.step(f"dc.{tag}.M_p", "Момент стоянки", "M_{п}",
                   r"\frac{U k_\Phi k\Phi_{ном}}{R}",
                   rf"\frac{{{fmt(U)} \cdot {fmt(phi)} \cdot {fmt(nat['kPhi_nom'])}}}{{{fmt(R)}}}",
                   U * nat["kPhi_nom"] * phi / R, "Н·м", 5, ref="п. 1.2")
    return dict(U=U, R=R, phi=phi, n0=n0, I_p=I_p, M_p=M_p)


def dc_run(m: DCMotor, variant):
    nat = dc_natural(m)
    arts = []
    for i, (name, phi, rk, uk) in enumerate(dc_art_conditions(variant), 1):
        log.section(f"1.2 ДПТ: искусственная характеристика {i}: {name}")
        log.text("Множители", rf"k_\Phi = {fmt(phi)};\ k_R = {fmt(rk)};\ k_U = {fmt(uk)}")
        arts.append(dict(name=name, **dc_artificial(m, nat, f"art{i}", phi, rk, uk)))
    return nat, arts


def plot_dc(m, nat, arts, outdir):
    def draw(ax_i, ax_m, d, label, **kw):
        ax_i.plot([0, d["I_p"]], [d["n0"], 0], label=label, **kw)
        ax_m.plot([0, d["M_p"]], [d["n0"], 0], label=label, **kw)

    def setup(fig, ax_i, ax_m):
        ax_i.set(xlabel="I, А", ylabel="n, об/мин", title="n = f(I)")
        ax_m.set(xlabel="M, Н·м", ylabel="n, об/мин", title="n = f(M)")
        for ax in (ax_i, ax_m):
            ax.grid(True, alpha=.4)
            ax.set_xlim(left=0)
            ax.set_ylim(bottom=0)
            ax.legend()
        fig.tight_layout()

    fig, (a, b) = plt.subplots(1, 2, figsize=(11, 4.5))
    draw(a, b, nat, "естественная", color="k")
    a.plot(m.I_nom, m.n_nom, "ro")
    b.plot(nat["M_nom"], m.n_nom, "ro")
    setup(fig, a, b)
    fig.savefig(outdir / "fig_1_1.png", dpi=300)
    plt.close(fig)

    for i, art in enumerate(arts, 1):
        fig, (a, b) = plt.subplots(1, 2, figsize=(11, 4.5))
        draw(a, b, nat, "естественная", color="k")
        draw(a, b, art, "искусственная: " + art["name"], color="C3")
        setup(fig, a, b)
        fig.savefig(outdir / f"fig_1_2_{i}.png", dpi=300)
        plt.close(fig)


# ---------------------------------------------------------------- раздел 2
def ac_pole_pairs(n_nom, f1=F1):
    """п. 2.1: pп — делим 60·f1 на ближайшее большее nном число, на которое 3000 делится нацело (ряд 3000/p)."""
    cand = []
    p = 1
    while 60 * f1 / p > n_nom:
        cand.append((p, 60 * f1 / p))
        p += 1
    return cand[-1][0], cand


def ac_params(m: ACMotor):
    """Приведение ротора к статору, xкз (п. 2.1)."""
    log.section("2.1 АД: исходные данные, приведение ротора, число пар полюсов")
    P = log.data("ac.P", "Номинальная мощность", "P_{ном}", m.P_kW * 1e3, "Вт", 5)
    n_nom = log.data("ac.n_nom", "Номинальная частота вращения", "n_{ном}", m.n_nom, "об/мин", 4)
    log.data("ac.U_line", "Линейное напряжение", "U_{л}", U_LINE, "В", 3)
    U = log.data("ac.U_ph", "Фазное напряжение", "U_{ф}", U_PHASE, "В", 3, note="Uл/√3 ≈ 219,4 → принято 220, как в образце")
    r1 = log.data("ac.r1", "Активное сопротивление фазы статора", "r_1", m.r1, "Ом", 3)
    x1 = log.data("ac.x1", "Реактивное сопротивление рассеяния статора", "x_1", m.x1, "Ом", 3)
    r2 = log.data("ac.r2", "Активное сопротивление фазы ротора", "r_2", m.r2, "Ом", 3)
    x2 = log.data("ac.x2", "Реактивное сопротивление рассеяния ротора", "x_2", m.x2, "Ом", 3)
    E2k = log.data("ac.E2k", "Линейная ЭДС между кольцами ротора", "E_{2к}", m.E2k, "В", 3)
    f1 = log.data("ac.f1", "Частота сети", "f_1", F1, "Гц", 3)

    log.text("Ряд синхронных скоростей (п. 2.1: 60·f1 / ближайшее большее nном)",
             rf"n_0 = \frac{{60 f_1}}{{p_{{п}}}} = \frac{{{fmt(SYNC_NUM, 4)}}}{{p_{{п}}}}")
    p, cand = ac_pole_pairs(n_nom)
    for pc, nc in cand:
        print(f"      pп = {pc}: n0 = {SYNC_NUM:.0f}/{pc} = {nc:.0f} об/мин  > nном = {n_nom:.0f}")
    print(f"      → ближайшая синхронная скорость выше nном: n0 = {cand[-1][1]:.0f} об/мин, pп = {p}")
    p = log.step("ac.p", "Число пар полюсов", "p_{п}", r"\frac{60 f_1}{n_0}",
                 rf"\frac{{{fmt(SYNC_NUM, 4)}}}{{{fmt(cand[-1][1], 4)}}}", p, "", 3, ref="п. 2.1",
                 note=f"{p} пары полюсов = {2 * p} полюсов; тип 4МТН225L6 — 6 полюсов (последняя цифра в типе)")
    E2f = log.step("ac.E2f", "Фазная ЭДС неподвижного ротора", "E_{2ф}",
                   r"\frac{E_{2к}}{\sqrt{3}}", rf"\frac{{{fmt(E2k)}}}{{\sqrt{{3}}}}",
                   E2k / math.sqrt(3), "В", 4, ref="п. 2.1")
    k = log.step("ac.k", "Коэффициент трансформации", "k",
                 r"\frac{0,95 U_{ф}}{E_{2ф}}", rf"\frac{{{fmt(K_SAFETY)} \cdot {fmt(U)}}}{{{fmt(E2f)}}}",
                 K_SAFETY * U / E2f, "", 4, ref="п. 2.1")
    x2p = log.step("ac.x2p", "Приведённое реактивное сопротивление ротора", "x_2'",
                   r"k^2 x_2", rf"{fmt(k)}^2 \cdot {fmt(x2)}", k ** 2 * x2, "Ом", 4, ref="п. 2.1")
    r2p = log.step("ac.r2p", "Приведённое активное сопротивление ротора", "r_2'",
                   r"k^2 r_2", rf"{fmt(k)}^2 \cdot {fmt(r2)}", k ** 2 * r2, "Ом", 4, ref="п. 2.1")
    xkz = log.step("ac.xkz", "Индуктивное сопротивление короткого замыкания", "x_{кз}",
                   r"x_1 + x_2'", rf"{fmt(x1)} + {fmt(x2p)}", x1 + x2p, "Ом", 4, ref="п. 2.1")
    return dict(E2f=E2f, k=k, r2p=r2p, x2p=x2p, xkz=xkz), p


def klos(s, Mmax, sk, a):
    """M(s) по формуле Клосса с учётом r1 (п. 2.1)."""
    s = np.asarray(s, dtype=float)
    return 2 * Mmax * (1 + a * sk) / (s / sk + sk / s + 2 * a * sk)


def ac_char(U, r1, xkz, f1, p, r2p):
    """Mmax, sk (двигательный и генераторный режимы), n0, ω0, a, Mп — п. 2.1. Пишет шаги в лог."""
    n0 = log.step("n0", "Скорость идеального х.х.", "n_0", r"\frac{60 f_1}{p_{п}}",
                  rf"\frac{{60 \cdot {fmt(f1)}}}{{{fmt(p, 3)}}}", 60 * f1 / p, "об/мин", 4, ref="п. 2.1")
    w0 = log.step("w0", "Угловая скорость поля", r"\omega_0",
                  r"\frac{2\pi f_1}{p_{п}}", rf"\frac{{2\pi \cdot {fmt(f1)}}}{{{fmt(p, 3)}}}",
                  2 * math.pi * f1 / p, "рад/с", 5, ref="п. 2.1",
                  note=f"проверка 2π·n0/60 = {fmt(2 * math.pi * n0 / 60, 5)}")
    Z = log.step("Z", "Вспомогательная величина √(r1²+xкз²)", r"Z_{к}",
                 r"\sqrt{r_1^2 + x_{кз}^2}", rf"\sqrt{{{fmt(r1)}^2 + {fmt(xkz)}^2}}",
                 math.hypot(r1, xkz), "Ом", 4, ref="п. 2.1")
    a = log.step("a", "Постоянная двигателя", "a", r"\frac{r_1}{r_2'}",
                 rf"\frac{{{fmt(r1)}}}{{{fmt(r2p)}}}", r1 / r2p, "", 4, ref="п. 2.1")
    Mmax_d = log.step("Mmax_d", "Максимальный момент, двигательный режим", r"M_{max\,д}",
                      r"\frac{m U_{ф}^2}{2\omega_0 \left(r_1 + \sqrt{r_1^2 + x_{кз}^2}\right)}",
                      rf"\frac{{{M_PHASES} \cdot {fmt(U)}^2}}{{2 \cdot {fmt(w0, 5)} \cdot ({fmt(r1)} + {fmt(Z)})}}",
                      M_PHASES * U ** 2 / (2 * w0 * (r1 + Z)), "Н·м", 5, ref="п. 2.1")
    Mmax_g = log.step("Mmax_g", "Максимальный момент, генераторный режим", r"M_{max\,г}",
                      r"\frac{m U_{ф}^2}{2\omega_0 \left(r_1 - \sqrt{r_1^2 + x_{кз}^2}\right)}",
                      rf"\frac{{{M_PHASES} \cdot {fmt(U)}^2}}{{2 \cdot {fmt(w0, 5)} \cdot ({fmt(r1)} - {fmt(Z)})}}",
                      M_PHASES * U ** 2 / (2 * w0 * (r1 - Z)), "Н·м", 5, ref="п. 2.1")
    sk = log.step("sk", "Критическое скольжение (двигательный режим; в генераторном −sк)", "s_{к}",
                  r"\frac{r_2'}{\sqrt{r_1^2 + x_{кз}^2}}", rf"\frac{{{fmt(r2p)}}}{{{fmt(Z)}}}",
                  r2p / Z, "", 4, ref="п. 2.1")
    Mst = log.step("M_start", "Пусковой момент (s = 1) по формуле Клосса", "M_{п}",
                   r"\frac{2 M_{max\,д}(1 + a s_{к})}{\frac{1}{s_{к}} + \frac{s_{к}}{1} + 2 a s_{к}}",
                   rf"\frac{{2 \cdot {fmt(Mmax_d, 5)} \cdot (1 + {fmt(a)} \cdot {fmt(sk)})}}"
                   rf"{{\frac{{1}}{{{fmt(sk)}}} + \frac{{{fmt(sk)}}}{{1}} + 2 \cdot {fmt(a)} \cdot {fmt(sk)}}}",
                   float(klos(1.0, Mmax_d, sk, a)), "Н·м", 4, ref="п. 2.1")
    assert abs(float(klos(sk, Mmax_d, sk, a)) - Mmax_d) < 1e-6
    return dict(n0=n0, w0=w0, Z=Z, a=a, Mmax_d=Mmax_d, Mmax_g=Mmax_g, sk_d=sk, sk_g=-sk, M_start=Mst)


def ac_run(m: ACMotor, variant):
    prm, p = ac_params(m)
    base = dict(U=U_PHASE, r1=m.r1, xkz=prm["xkz"], f1=F1, p=p)
    nat = ac_char(base["U"], base["r1"], base["xkz"], base["f1"], base["p"], prm["r2p"])
    s_nom = log.step("ac.s_nom", "Номинальное скольжение", "s_{ном}",
                     r"\frac{n_0 - n_{ном}}{n_0}", rf"\frac{{{fmt(nat['n0'])} - {fmt(m.n_nom)}}}{{{fmt(nat['n0'])}}}",
                     (nat["n0"] - m.n_nom) / nat["n0"], "", 4, ref="п. 2.1")
    w_nom = rpm_to_rad(m.n_nom)
    M_nom_shaft = log.step("ac.M_nom", "Номинальный момент на валу", "M_{ном}",
                           r"\frac{P_{ном}}{\omega_{ном}}",
                           rf"\frac{{{fmt(m.P_kW * 1e3, 5)}}}{{{fmt(w_nom)}}}", m.P_kW * 1e3 / w_nom, "Н·м", 4,
                           ref="справочно", note="ω_ном = 2π·nном/60")
    M_nom_klos = log.step("ac.M_nom_klos", "Момент при sном по формуле Клосса (для сравнения)", "M(s_{ном})",
                          r"\frac{2 M_{max\,д}(1 + a s_{к})}{\frac{s_{ном}}{s_{к}} + \frac{s_{к}}{s_{ном}} + 2 a s_{к}}",
                          rf"\frac{{2 \cdot {fmt(nat['Mmax_d'], 5)} \cdot (1 + {fmt(nat['a'])} \cdot {fmt(nat['sk_d'])})}}"
                          rf"{{\frac{{{fmt(s_nom)}}}{{{fmt(nat['sk_d'])}}} + \frac{{{fmt(nat['sk_d'])}}}{{{fmt(s_nom)}}} + 2 \cdot {fmt(nat['a'])} \cdot {fmt(nat['sk_d'])}}}",
                          float(klos(s_nom, nat["Mmax_d"], nat["sk_d"], nat["a"])), "Н·м", 4, ref="справочно")
    arts = []
    for i, (name, key, mult) in enumerate(ac_art_conditions(variant), 1):
        log.section(f"2.2 АД: искусственная характеристика {i}: {name}")
        q = dict(base)
        q[key] = base[key] * mult
        q["p"] = int(round(q["p"]))
        sym = {"U": "U_{ф}", "r1": "r_1", "xkz": "x_{кз}", "f1": "f_1", "p": "p_{п}"}[key]
        log.step(f"ac.art{i}.param", f"Изменённый параметр {sym}", sym + "'", rf"k \cdot {sym}",
                 rf"{fmt(mult)} \cdot {fmt(base[key])}", q[key], {"U": "В", "r1": "Ом", "xkz": "Ом", "f1": "Гц", "p": ""}[key], 4, ref="п. 2.2",
                 note="остальные параметры — как в естественной характеристике")
        c = ac_char(q["U"], q["r1"], q["xkz"], q["f1"], q["p"], prm["r2p"])
        arts.append(dict(name=name, key=key, mult=mult, **c))
    return dict(params=prm, p=p, s_nom=s_nom, M_nom_shaft=M_nom_shaft,
                M_nom_klos=M_nom_klos, natural=nat), arts


def plot_ac(m, res, arts, outdir):
    nat = res["natural"]
    s = np.concatenate([np.linspace(-2, -1e-3, 800), np.linspace(1e-3, 2, 800)])

    def curve(c):
        M = np.where(s > 0, klos(s, c["Mmax_d"], c["sk_d"], c["a"]),
                     klos(s, c["Mmax_g"], c["sk_g"], c["a"]))
        return M, c["n0"] * (1 - s)

    def fmtax(ax):
        ax.set(xlabel="M, Н·м", ylabel="n, об/мин")
        ax.axhline(0, color="gray", lw=.5)
        ax.axvline(0, color="gray", lw=.5)
        ax.grid(True, alpha=.4)
        ax.legend()

    fig, ax = plt.subplots(figsize=(7, 5))
    ax.plot(*curve(nat), "k", label="естественная")
    fmtax(ax)
    fig.tight_layout()
    fig.savefig(outdir / "fig_2_1.png", dpi=300)
    plt.close(fig)

    for i, art in enumerate(arts, 1):
        fig, ax = plt.subplots(figsize=(7, 5))
        ax.plot(*curve(nat), "k", label="естественная")
        ax.plot(*curve(art), "C3", label="искусственная: " + art["name"])
        fmtax(ax)
        fig.tight_layout()
        fig.savefig(outdir / f"fig_2_2_{i}.png", dpi=300)
        plt.close(fig)


# ---------------------------------------------------------------- раздел 3
K_MP = 1.9            # Mп = 1,9·Mном: пусковой момент на 1-й ступени (как в образце)
K_MPER = 1.05         # целевой Mпер = 1,05·Mном (нижняя граница диапазона 1,05…1,3·Mном; по замечанию преподавателя)
MM_N0 = 200.0         # масштаб графического построения: n0 = 200 мм
MM_M_P = 160.0        # масштаб построения от руки: Mп = 160 мм (по оси моментов)
MM_READ = 0.5         # точность «измерения» отрезков линейкой, мм
TOL_PCT = 10.0        # допустимое расхождение графического и аналитического расчёта, %


def ac_start_diagram(m: ACMotor, res):
    """Раздел 3: число и сопротивления ступеней пускового реостата (аналитически и «графически»)."""
    log.section("3. Пусковая диаграмма: реостатный пуск АД (прямолинейные характеристики)")
    prm, nat = res["params"], res["natural"]
    r2p, k, n0, s_nom, M_nom = prm["r2p"], prm["k"], nat["n0"], res["s_nom"], res["M_nom_shaft"]

    M_c = log.step("st.Mc", "Момент сопротивления (принят равным номинальному)", "M_{с}", "M_{ном}",
                   fmt(M_nom, 5), M_nom, "Н·м", 5, ref="п. 3, 4")
    M_p = log.step("st.Mp", "Пусковой момент на 1-й ступени", "M_{п}", r"k_{п} M_{ном}",
                   rf"{fmt(K_MP)} \cdot {fmt(M_nom, 5)}", K_MP * M_nom, "Н·м", 5, ref="п. 3",
                   note="k_п = 1,9 — как в образце; должно быть Mп < Mmax")
    assert M_p < nat["Mmax_d"]
    Mper_t = log.step("st.Mper_t", "Желаемый момент переключения", "M_{пер}'", r"k_{пер} M_{ном}",
                      rf"{fmt(K_MPER)} \cdot {fmt(M_nom, 5)}", K_MPER * M_nom, "Н·м", 5, ref="п. 3",
                      note="k_пер = 1,05 — нижняя граница диапазона 1,05…1,3")
    R_p = log.step("st.Rp", "Полное приведённое сопротивление цепи ротора при пуске", "R_{п}",
                   r"\frac{r_2' M_{ном}}{s_{ном} M_{п}}",
                   rf"\frac{{{fmt(r2p)} \cdot {fmt(M_nom, 5)}}}{{{fmt(s_nom)} \cdot {fmt(M_p, 5)}}}",
                   r2p * M_nom / (s_nom * M_p), "Ом", 4, ref="п. 3")
    lam0 = log.step("st.lam0", "Предварительное λ", r"\lambda'", r"\frac{M_{п}}{M_{пер}'}",
                    rf"\frac{{{fmt(M_p, 5)}}}{{{fmt(Mper_t, 5)}}}", M_p / Mper_t, "", 4, ref="п. 3")
    n_calc = log.step("st.n_calc", "Число ступеней (расчётное)", "n",
                      r"\frac{\ln\frac{R_{п}}{r_2'}}{\ln \lambda'}",
                      rf"\frac{{\ln\frac{{{fmt(R_p)}}}{{{fmt(r2p)}}}}}{{\ln {fmt(lam0)}}}",
                      math.log(R_p / r2p) / math.log(lam0), "", 4, ref="п. 3")
    n_st = int(math.ceil(n_calc - 1e-9))
    n_st = log.step("st.n", "Число ступеней (округлено вверх)", "m", r"\lceil n \rceil",
                    rf"\lceil {fmt(n_calc)} \rceil", n_st, "", 3, ref="п. 3")
    n_st = int(n_st)
    lam = log.step("st.lam", "Уточнённое λ (чтобы выход на естественную был при Mп)", r"\lambda",
                   r"\left(\frac{R_{п}}{r_2'}\right)^{\frac{1}{m}}",
                   rf"\left(\frac{{{fmt(R_p)}}}{{{fmt(r2p)}}}\right)^{{\frac{{1}}{{{n_st}}}}}",
                   (R_p / r2p) ** (1 / n_st), "", 4, ref="п. 3")
    M_per = log.step("st.Mper", "Момент переключения", "M_{пер}", r"\frac{M_{п}}{\lambda}",
                     rf"\frac{{{fmt(M_p, 5)}}}{{{fmt(lam)}}}", M_p / lam, "Н·м", 5, ref="п. 3")
    assert M_c < M_per < M_p and 1.05 * M_nom <= M_per <= 1.3 * M_nom
    assert abs(lam ** n_st - R_p / r2p) < 1e-9

    Rd = []
    for i in range(1, n_st + 1):
        Rd.append(log.step(f"st.Rd{i}", f"Сопротивление ступени {i}", f"R_{{д{i}}}",
                           rf"r_2' \lambda^{{{i - 1}}} (\lambda - 1)",
                           rf"{fmt(r2p)} \cdot {fmt(lam)}^{{{i - 1}}} \cdot ({fmt(lam)} - 1)",
                           r2p * lam ** (i - 1) * (lam - 1), "Ом", 4, ref="п. 3"))
    sum_rd = log.step("st.sumRd", "Проверка: сумма ступеней", r"\sum R_{д}", r"R_{п} - r_2'",
                      rf"{fmt(R_p)} - {fmt(r2p)}", sum(Rd), "Ом", 4, ref="проверка")
    assert abs(sum_rd - (R_p - r2p)) < 1e-9
    Rd_rotor = []
    for i in range(1, n_st + 1):
        Rd_rotor.append(log.step(f"st.Rdr{i}", f"Ступень {i}, реальное сопротивление в цепи ротора",
                                 f"R_{{д{i}}}^{{рот}}", rf"\frac{{R_{{д{i}}}}}{{k^2}}",
                                 rf"\frac{{{fmt(Rd[i - 1])}}}{{{fmt(k)}^2}}", Rd[i - 1] / k ** 2, "Ом", 4,
                                 ref="справочно", note="приведённое → реальное: делим на k²"))

    # --- «графический» способ: построение лесенки в масштабе, отрезки округляются как при измерении
    log.text("Графическое построение: отрезки на вертикали M = Mп в масштабе n0 = 200 мм",
             rf"l = \frac{{\Delta n \cdot {fmt(MM_N0)}}}{{n_0}},\quad \text{{точность отсчёта}}\ {fmt(MM_READ)}\ \text{{мм}}")
    kmm = MM_N0 / n0
    n_at_Mp = [n0 * (1 - s_nom * lam ** (n_st + 1 - j) * M_p / M_nom) for j in range(1, n_st + 2)]
    seg = [(n_at_Mp[j] - n_at_Mp[j - 1]) * kmm for j in range(1, n_st + 1)]  # между хар-ой j и j+1
    l_exact = {i: seg[n_st - i] for i in range(1, n_st + 1)}                 # i = n_st + 1 - j
    l_zhk_exact = (n0 - n_at_Mp[-1]) * kmm
    l_zhk = round(l_zhk_exact / MM_READ) * MM_READ
    m2 = log.step("st.m2", "Масштаб сопротивления", "m_2", r"\frac{r_2'}{l_{жк}}",
                  rf"\frac{{{fmt(r2p)}}}{{{fmt(l_zhk)}}}", r2p / l_zhk, "Ом/мм", 4, ref="п. 3",
                  note=f"l_жк = {l_zhk:.1f} мм (точное {l_zhk_exact:.2f})")
    Rd_gr, dev = [], []
    for i in range(1, n_st + 1):
        l_i = round(l_exact[i] / MM_READ) * MM_READ
        r_g = log.step(f"st.Rdg{i}", f"Ступень {i}, графически", f"R_{{д{i}}}'", rf"m_2 l_{{{i}}}",
                       rf"{fmt(m2)} \cdot {fmt(l_i)}", m2 * l_i, "Ом", 4, ref="п. 3",
                       note=f"l_{i} = {l_i:.1f} мм")
        Rd_gr.append(r_g)
        d = abs(r_g - Rd[i - 1]) / Rd[i - 1] * 100
        dev.append(d)
        log.step(f"st.dev{i}", f"Расхождение по ступени {i}", rf"\delta_{{{i}}}",
                 rf"\frac{{|R_{{д{i}}}' - R_{{д{i}}}|}}{{R_{{д{i}}}}} \cdot 100",
                 rf"\frac{{|{fmt(r_g)} - {fmt(Rd[i - 1])}|}}{{{fmt(Rd[i - 1])}}} \cdot 100", d, "%", 3,
                 ref="проверка 5–10 %")
    assert max(dev) <= TOL_PCT, f"расхождение {max(dev):.1f} % > {TOL_PCT} %"
    # --- координаты точек диаграммы (для построения от руки): M, n и положение в миллиметрах
    names, it = {}, iter(POINT_LETTERS)
    assert 2 * n_st + 1 <= len(POINT_LETTERS), f"для {n_st} ступеней не хватает букв для точек диаграммы"
    for j in range(1, n_st + 2):
        names[("p", j)] = next(it)
        if j <= n_st:
            names[("per", j)] = next(it)
    points = []
    for j in range(1, n_st + 2):
        for kind, M in (("p", M_p), ("per", M_per)):
            if kind == "per" and j > n_st:
                continue
            n_val = n0 * (1 - s_nom * lam ** (n_st + 1 - j) * M / M_nom)
            n_val = 0.0 if abs(n_val) < 1e-9 else n_val   # n = 0 в точке «а» без ошибки округления
            points.append(dict(name=names[(kind, j)], char=j, M=M, n=n_val,
                               x_mm=M / M_p * MM_M_P, y_mm=n_val * kmm))
    l_mm = [Rd_gr[i] / m2 for i in range(n_st)]
    return dict(M_c=M_c, M_p=M_p, M_per=M_per, R_p=R_p, lam=lam, n_calc=n_calc, n_st=n_st,
                Rd=Rd, Rd_rotor=Rd_rotor, Rd_graph=Rd_gr, dev_pct=dev, m2=m2, r2p=r2p, n0=n0,
                s_nom=s_nom, M_nom=M_nom, points=points, l_mm=l_mm, l_zhk_mm=r2p / m2,
                mm_n0=MM_N0, mm_Mp=MM_M_P)


POINT_LETTERS = list("абвгдежзилорсуфхцчшэюя")   # без к, м, н, п: они заняты под k, m, n и «пуск»


def _nice_step(span, target=10):
    raw = span / target
    mag = 10 ** math.floor(math.log10(raw))
    return min(c * mag for c in (1, 2, 2.5, 5, 10) if c * mag >= raw)


def plot_start_diagram(st, outdir):
    """Рис. 3 методички: пусковая диаграмма для линейных характеристик (по моменту) с размеченными шкалами."""
    n0, s_nom, M_nom, lam, m = st["n0"], st["s_nom"], st["M_nom"], st["lam"], st["n_st"]
    M_p, M_per, M_c = st["M_p"], st["M_per"], st["M_c"]
    n_at = lambda j, M: n0 * (1 - s_nom * lam ** (m + 1 - j) * M / M_nom)   # характеристика j, j = 1…m+1
    pts = {(q["char"], "p" if abs(q["M"] - M_p) < 1e-9 else "per"): q for q in st["points"]}

    fig, ax = plt.subplots(figsize=(11, 10.5))
    ax.set_xlim(-0.06 * M_p, 1.42 * M_p)
    ax.set_ylim(-0.2 * n0, 1.1 * n0)
    ax.axis("off")
    arrow = dict(arrowstyle="-|>", color="k", lw=1.4, shrinkA=0, shrinkB=0)
    ax.annotate("", xy=(0, 1.08 * n0), xytext=(0, 0), arrowprops=arrow)
    ax.annotate("", xy=(1.12 * M_p, 0), xytext=(0, 0), arrowprops=arrow)
    ax.text(0.012 * M_p, 1.085 * n0, "$n$, об/мин", ha="left", va="center", fontsize=12)
    ax.text(1.135 * M_p, 0.03 * n0, "$I$, А", ha="left", va="center", fontsize=12)
    ax.text(1.135 * M_p, -0.02 * n0, "$M$, Н·м", ha="left", va="center", fontsize=12)

    # шкалы: деления и числа по осям M и n, тонкая сетка
    sx, sy = _nice_step(1.1 * M_p), _nice_step(n0)
    for x in np.arange(0, 1.1 * M_p + 1e-9, sx):
        ax.plot([x, x], [-0.008 * n0, 0.008 * n0], "k", lw=0.9)
        ax.plot([x, x], [0, n0], color="0.85", lw=0.5, zorder=0)
        ax.text(x, -0.016 * n0, f"{x:.0f}", ha="center", va="top", fontsize=8, color="0.25")
    for y in np.arange(0, n0 + 1e-9, sy):
        ax.plot([-0.008 * M_p, 0.008 * M_p], [y, y], "k", lw=0.9)
        ax.plot([0, 1.1 * M_p], [y, y], color="0.85", lw=0.5, zorder=0)
        if abs(y - n0) > 1e-6:      # на уровне n0 стоит собственная подпись
            ax.text(-0.012 * M_p, y, f"{y:.0f}", ha="right", va="center", fontsize=8, color="0.25")

    # естественная и реостатные характеристики: от (0, n0) до M = Mп
    for j in range(1, m + 2):
        ax.plot([0, M_p], [n0, n_at(j, M_p)], "k", lw=1.8, solid_capstyle="butt")

    # вертикали: Mс — сплошная, Mпер и Mп — штриховые; горизонталь на уровне n0
    ax.plot([M_c, M_c], [0, n0], "k", lw=2.2)
    for x in (M_per, M_p):
        ax.plot([x, x], [0, n0], "k", lw=0.9, ls=(0, (5, 4)))
    ax.plot([0, M_p], [n0, n0], "k", lw=0.9, ls=(0, (5, 4)))
    ax.text(-0.04 * M_p, n0, f"$n_0$ = {n0:.0f}", ha="right", va="center", fontsize=11)

    for j in range(1, m + 1):                       # переход со ступени j на j+1: горизонталь (Mпер, j) → (Mп, j+1)
        ax.plot([M_per, M_p], [n_at(j, M_per)] * 2, "k", lw=0.9, ls=(0, (5, 4)))
    for j in range(1, m + 2):
        q = pts[(j, "p")]
        ax.plot([M_p], [q["n"]], "ko", ms=4)
        ax.text(M_p + 0.012 * M_p, q["n"], q["name"], style="italic", fontsize=12, ha="left", va="center")
        if j <= m:
            q2 = pts[(j, "per")]
            ax.plot([M_per], [q2["n"]], "ko", ms=4)
            ax.text(M_per + 0.008 * M_p, q2["n"] + 0.018 * n0, q2["name"], style="italic", fontsize=12,
                    ha="left", va="center")
    ax.text(M_p + 0.012 * M_p, n0, "$k$", fontsize=12, ha="left", va="center")
    ax.text(M_c - 0.012 * M_p, n0 + 0.005 * n0, "$m$", fontsize=12, ha="right", va="bottom")
    ax.text(M_per + 0.008 * M_p, n0 + 0.005 * n0, "$n$", fontsize=12, ha="left", va="bottom")

    # отрезки l_i, l_жк на вертикали M = Mп (по ним считается R_дi = m2·l_i)
    xd = 1.2 * M_p
    dim = dict(arrowstyle="<|-|>", color="k", lw=0.9, shrinkA=0, shrinkB=0, mutation_scale=7)
    levels = [n_at(j, M_p) for j in range(1, m + 2)] + [n0]
    for idx in range(m + 1):          # idx = 0…m: между характеристиками j = idx+1 и idx+2 (последний — до k)
        y0, y1 = levels[idx], levels[idx + 1]
        if idx < m:
            i = m - idx
            lab, mm = rf"$l_{{{i}}}$", st["l_mm"][i - 1]
        else:
            lab, mm = r"$l_{жк}$", st["l_zhk_mm"]
        ax.annotate("", xy=(xd, y0), xytext=(xd, y1), arrowprops=dim)
        for y in (y0, y1):
            ax.plot([M_p + 0.06 * M_p, xd + 0.015 * M_p], [y, y], color="0.6", lw=0.5)
        ax.text(xd + 0.025 * M_p, (y0 + y1) / 2, lab + f" = {mm:.1f} мм".replace(".", ","),
                ha="left", va="center", fontsize=10)

    # подписи под осью: символы и числовые значения
    for x, lab, val, ha in [(M_c, r"$M_{\mathrm{с}},\ I_{\mathrm{с}}$", M_c, "right"),
                            (M_per, r"$M_{\mathrm{пер}},\ I_{\mathrm{пер}}$", M_per, "left"),
                            (M_p, r"$M_{\mathrm{п}},\ I_{\mathrm{п}}$", M_p, "center")]:
        dx = {"right": 0.006, "left": -0.006, "center": 0}[ha] * M_p
        ax.text(x + dx, -0.075 * n0, lab, ha=ha, va="top", fontsize=12)
        ax.text(x + dx, -0.125 * n0, f"{val:.1f} Н·м".replace(".", ","), ha=ha, va="top", fontsize=9, color="0.3")
    fig.savefig(outdir / "fig_3_1.png", dpi=300, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------- раздел 4
def ac_transient(m: ACMotor, res, st):
    """Раздел 4: механический переходный процесс при ступенчатом пуске."""
    log.section("4. Механический переходный процесс при ступенчатом пуске")
    n0, s_nom, M_nom, lam, n_st = st["n0"], st["s_nom"], st["M_nom"], st["lam"], st["n_st"]
    M_p, M_per, M_c, J = st["M_p"], st["M_per"], st["M_c"], m.J
    n_nom = m.n_nom
    J = log.data("tr.J", "Момент инерции системы (только двигатель: данных о нагрузке нет)", "J", J, "кг·м²", 3)
    M_y = log.step("tr.My", "Установившийся момент для всех ступеней", "M_{уi}", "M_{ном}",
                   fmt(M_nom, 5), M_nom, "Н·м", 5, ref="п. 4")
    Mi_n = log.step("tr.Min", "Избыточный момент в начале ступени", "M_{инач}", r"M_{п} - M_{с}",
                    rf"{fmt(M_p, 5)} - {fmt(M_c, 5)}", M_p - M_c, "Н·м", 5, ref="п. 4")
    Mi_k = log.step("tr.Mik", "Избыточный момент в конце ступени", "M_{иконi}", r"M_{пер} - M_{с}",
                    rf"{fmt(M_per, 5)} - {fmt(M_c, 5)}", M_per - M_c, "Н·м", 5, ref="п. 4")
    ln_r = log.step("tr.ln", "Логарифм отношения избыточных моментов (одинаков для всех ступеней)", r"\ln\frac{M_{инач}}{M_{икон}}",
                    r"\ln\frac{M_{инач}}{M_{икон}}", rf"\ln\frac{{{fmt(Mi_n, 5)}}}{{{fmt(Mi_k, 5)}}}",
                    math.log(Mi_n / Mi_k), "", 4, ref="п. 4")
    rows, stages = [], []
    n_start = 0.0
    for i in range(1, n_st + 2):
        e = n_st + 1 - i
        ex = rf"\lambda^{{{e}}}" if e > 0 else "1"
        lam_e = lam ** e
        n_y = log.step(f"tr.ny{i}", f"Ступень {i}: установившаяся частота вращения", f"n_{{у{i}}}",
                       rf"n_0 (1 - s_{{ном}} {ex})", rf"{fmt(n0)} \cdot (1 - {fmt(s_nom)} \cdot {fmt(lam_e)})",
                       n0 * (1 - s_nom * lam_e), "об/мин", 4, ref="п. 4")
        M_k = M_nom / (s_nom * lam_e)
        M_ik = log.step(f"tr.Mikr{i}", f"Ступень {i}: критический избыточный момент (n = 0)", f"M_{{икр{i}}}",
                        rf"\frac{{M_{{ном}}}}{{s_{{ном}} {ex}}} - M_{{с}}",
                        rf"\frac{{{fmt(M_nom, 5)}}}{{{fmt(s_nom)} \cdot {fmt(lam_e)}}} - {fmt(M_c, 5)}",
                        M_k - M_c, "Н·м", 5, ref="п. 4")
        n_k = log.step(f"tr.nk{i}", f"Ступень {i}: частота вращения в конце ступени", f"n_{{кон{i}}}",
                       rf"n_0 \left(1 - s_{{ном}} {ex} \frac{{M_{{пер}}}}{{M_{{ном}}}}\right)",
                       rf"{fmt(n0)} \cdot \left(1 - {fmt(s_nom)} \cdot {fmt(lam_e)} \cdot \frac{{{fmt(M_per, 5)}}}{{{fmt(M_nom, 5)}}}\right)",
                       n0 * (1 - s_nom * lam_e * M_per / M_nom), "об/мин", 4, ref="п. 4")
        T = log.step(f"tr.T{i}", f"Ступень {i}: механическая постоянная времени", f"T_{{М{i}}}",
                     rf"\frac{{J n_{{у{i}}}}}{{9{{,}}55 M_{{икр{i}}}}}".replace("{,}", ","),
                     rf"\frac{{{fmt(J)} \cdot {fmt(n_y)}}}{{9,55 \cdot {fmt(M_ik, 5)}}}",
                     J * n_y / (9.55 * M_ik), "с", 4, ref="п. 4")
        T2 = J * (n_k - n_start) / (9.55 * (Mi_n - Mi_k))
        assert abs(T - T2) / T < 1e-6, (i, T, T2)
        t_i = log.step(f"tr.t{i}", f"Ступень {i}: время разгона", f"t_{{{i}}}",
                       rf"T_{{М{i}}} \ln\frac{{M_{{инач}}}}{{M_{{икон}}}}",
                       rf"{fmt(T)} \cdot {fmt(ln_r)}", T * ln_r, "с", 4, ref="п. 4",
                       note=f"проверка по формуле методички J(nкон−nнач)/(9,55(Mинач−Mикон))·ln = "
                            f"{fmt(T2 * ln_r)} с")
        rows.append([i, n_start, n_k, n_y, T, t_i])
        stages.append(dict(i=i, n_start=n_start, n_end=n_k, n_y=n_y, T=T, t=t_i))
        n_start = n_k
    t_sum = log.step("tr.tsum", "Время разгона: сумма по ступеням (по графику)", "t_{\\Sigma}",
                     r"\sum t_{i}", " + ".join(fmt(r[5]) for r in rows), sum(r[5] for r in rows), "с", 4, ref="п. 4")
    t_an = log.step("tr.tan", "Время разгона по общей формуле", "t",
                    r"\frac{J n_{ном}}{9,55 (M_{инач} - M_{икон})} \ln\frac{M_{инач}}{M_{икон}}",
                    rf"\frac{{{fmt(J)} \cdot {fmt(n_nom)}}}{{9,55 \cdot ({fmt(Mi_n, 5)} - {fmt(Mi_k, 5)})}} \cdot {fmt(ln_r)}",
                    J * n_nom / (9.55 * (Mi_n - Mi_k)) * ln_r, "с", 4, ref="п. 4")
    dev = log.step("tr.dev", "Расхождение времён", r"\delta", r"\frac{|t_{\Sigma} - t|}{t} \cdot 100",
                   rf"\frac{{|{fmt(t_sum)} - {fmt(t_an)}|}}{{{fmt(t_an)}}} \cdot 100",
                   abs(t_sum - t_an) / t_an * 100, "%", 3, ref="проверка 5–10 %")
    assert dev <= TOL_PCT
    log.table("Таблица. Результаты расчёта по ступеням",
              ["Ступень", "n_нач, об/мин", "n_кон, об/мин", "n_уст, об/мин", "T_М, с", "t, с"],
              [[r[0]] + [round(x, 4) for x in r[1:]] for r in rows])
    return dict(stages=stages, t_sum=t_sum, t_an=t_an, dev_pct=dev, M_i_start=Mi_n, M_i_end=Mi_k)


SETTLE_T_M = 4          # установившийся режим показываем на интервале 4·Tм (≈ 98 % от конечного значения)
TAIL_MIN_SHARE = 0.2     # но не короче 20 % длительности ступенчатого разгона, чтобы участок был виден


def plot_transient(st, tr, outdir, m_nom_rpm):
    """Рис. 4 методички: n = f(t) и M = f(t) на общей оси времени; подписи уровней, кривых и интервалов t_i."""
    n0, s_nom, M_nom, lam, m = st["n0"], st["s_nom"], st["M_nom"], st["lam"], st["n_st"]
    M_p, M_per, M_c = st["M_p"], st["M_per"], st["M_c"]
    t0, T_all, N_all, M_all, edges = 0.0, [], [], [], [0.0]
    for s in tr["stages"]:
        t = np.linspace(0, s["t"], 200)
        ex = np.exp(-t / s["T"])
        N_all.append(s["n_y"] * (1 - ex) + s["n_start"] * ex)
        M_all.append(M_c * (1 - ex) + M_p * ex)
        T_all.append(t0 + t)
        t0 += s["t"]
        edges.append(t0)
    # установившийся режим: после t_m+1 двигатель на естественной характеристике плавно выходит на nном и Mс
    last = tr["stages"][-1]
    t_sw = edges[-1]                                   # конец ступенчатого разгона (M = Mпер)
    tail = max(SETTLE_T_M * last["T"], TAIL_MIN_SHARE * t_sw)
    tt = np.linspace(0, tail, 200)
    ex = np.exp(-tt / last["T"])
    N_all.append(last["n_y"] + (last["n_end"] - last["n_y"]) * ex)
    M_all.append(M_c + (M_per - M_c) * ex)
    T_all.append(t_sw + tt)
    T_all, N_all, M_all = map(np.concatenate, (T_all, N_all, M_all))
    t_end = t_sw + tail

    fig, (a, b) = plt.subplots(2, 1, figsize=(9, 8.5), sharex=True)
    a.plot(T_all, N_all, "k", lw=1.4)
    b.plot(T_all, M_all, "k", lw=1.4)
    for x in edges[1:]:
        for ax in (a, b):
            ax.axvline(x, color="gray", ls=(0, (5, 4)), lw=0.8)

    # уровни по оси моментов: Mп, Mпер, Mс (штриховые линии, подписи на оси)
    for y in (M_p, M_per):
        b.axhline(y, color="gray", ls=(0, (5, 4)), lw=0.8)
    b.axhline(M_c, color="k", lw=1.0)
    b.set_yticks([M_c, M_per, M_p])
    b.set_yticklabels([r"$M_{\mathrm{с}}$", r"$M_{\mathrm{пер}}$", r"$M_{\mathrm{п}}$"], fontsize=12)
    b.set_ylim(M_c - 0.06 * (M_p - M_c), M_p + 0.08 * (M_p - M_c))
    b.text(edges[1] * 0.45, M_c + 0.88 * (M_p - M_c), "$M = f(t)$", fontsize=12, ha="left", va="center")

    # уровни частоты вращения: n_i на концах ступеней, n_ном
    st_levels = [s["n_end"] for s in tr["stages"][:-1]] + [m_nom_rpm]
    names = [f"{v:.0f}" for v in st_levels]
    span = max(st_levels)
    for k in range(len(st_levels) - 1):
        if (st_levels[k + 1] - st_levels[k]) < 0.03 * span:   # подписи слились бы — опускаем меньшую
            names[k] = ""
    n_ticks = st_levels
    for y in n_ticks:
        a.axhline(y, color="gray", ls=(0, (5, 4)), lw=0.6)
    a.set_yticks(n_ticks)
    a.set_yticklabels(names, fontsize=11)
    a.set_ylim(0, 1.05 * max(n_ticks))
    a.text(edges[1] * 0.35, 0.18 * max(n_ticks), "$n = f(t)$", fontsize=12, ha="left", va="center")

    # подписи осей — горизонтально у концов осей, как в методичке
    a.annotate("$n$, об/мин", xy=(0, 1), xycoords="axes fraction", xytext=(-6, 10), textcoords="offset points",
               ha="right", va="bottom", fontsize=12)
    b.annotate("$M$, Н·м", xy=(0, 1), xycoords="axes fraction", xytext=(-6, 10), textcoords="offset points",
               ha="right", va="bottom", fontsize=12)
    b.annotate("$t$, с", xy=(1, 0), xycoords="axes fraction", xytext=(8, 0), textcoords="offset points",
               ha="left", va="center", fontsize=12)
    b.set_xlim(0, t_end * 1.04)
    raw = t_end / 10
    mag = 10 ** math.floor(math.log10(raw))
    step = min((c * mag for c in (1, 2, 2.5, 5, 10) if c * mag >= raw), default=mag)
    ticks = np.arange(0, t_end * 1.04, step)
    b.set_xticks(ticks)
    b.set_xticklabels([fmt(x, 3) for x in ticks], fontsize=10)
    for ax in (a, b):
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    for ax in (a, b):
        ax.axvspan(t_sw, t_end * 1.04, color="0.94", zorder=0)
    a.text((t_sw + t_end) / 2, 0.45 * max(n_ticks), "установив-\nшийся\nрежим", ha="center", va="center",
           fontsize=9, color="0.25")

    # интервалы t_i: двусторонние стрелки под осью времени, подписи в два ряда, чтобы узкие не сливались
    trans = b.get_xaxis_transform()
    for k, s in enumerate(tr["stages"]):
        x0, x1 = edges[k], edges[k + 1]
        b.annotate("", xy=(x0, -0.15), xytext=(x1, -0.15), xycoords=trans, textcoords=trans,
                   arrowprops=dict(arrowstyle="<|-|>", color="k", lw=1.0, shrinkA=0, shrinkB=0,
                                   mutation_scale=8), annotation_clip=False)
        narrow = (x1 - x0) < 0.06 * t_end
        y_lab = -0.20 if not (narrow and k % 2) else -0.27
        b.text((x0 + x1) / 2, y_lab, rf"$t_{{{k + 1}}}$", transform=trans, ha="center", va="center", fontsize=11,
               clip_on=False)
    fig.tight_layout()
    fig.subplots_adjust(bottom=0.2)
    fig.savefig(outdir / "fig_4_1.png", dpi=300)
    plt.close(fig)


# ---------------------------------------------------------------- раздел 5
BETA_DC = 0.5         # β в L = β·U/(pп·ω·Iном): 0,6…0,25, 0,25 — компенсированные машины (как в образце 0,5)
P_POLES_DC = 2        # число пар полюсов ДПТ: в Прил. 1 не указано, принято 2 (4 полюса), как в образце
SETTLE = 20           # затухание огибающей до 1/20 = 5 %
SETTLE_SHOW = 2.0     # на графике показываем 2·tпп, чтобы был виден установившийся режим


def dc_transient(m: DCMotor, nat):
    """Раздел 5: электромеханический переходный процесс ДПТ при пуске от сети по естественной характеристике."""
    log.section("5. Электромеханический переходный процесс ДПТ (пуск по естественной характеристике)")
    U = log.data("tr5.U", "Напряжение", "U_{ном}", m.U_nom, "В", 3)
    R = log.data("tr5.R", "Сопротивление якоря", "R_{я}", m.R_ya, "Ом", 3)
    I_nom = log.data("tr5.I", "Номинальный ток", "I_{ном}", m.I_nom, "А", 4)
    n_nom = log.data("tr5.n", "Номинальная частота вращения", "n_{ном}", m.n_nom, "об/мин", 4)
    J = log.data("tr5.J", "Момент инерции", "J", m.J, "кг·м²", 3)
    beta = log.data("tr5.beta", "Коэффициент β (0,6…0,25)", r"\beta", BETA_DC, "", 3,
                    note="принят как в образце; 0,25 — для компенсированных машин")
    p_pol = log.data("tr5.p", "Число пар полюсов", "p_{п}", P_POLES_DC, "", 2,
                     note="в Прил. 1 не указано, принято 2 (4 полюса) — уточнить по паспорту Д-812")
    M_nom, M_p, I_p, n0 = nat["M_nom"], nat["M_p"], nat["I_p"], nat["n0"]
    M_c = log.step("tr5.Mc", "Момент сопротивления (как в разд. 3–4)", "M_{с}", "M_{ном}",
                   fmt(M_nom, 5), M_nom, "Н·м", 5, ref="п. 5")
    n_c = log.step("tr5.nc", "Установившаяся частота вращения при Mс", "n_{с}", "n_{ном}",
                   fmt(n_nom), n_nom, "об/мин", 4, ref="п. 5")

    L = log.step("tr5.L", "Индуктивность цепи якоря", "L",
                 r"\frac{9,55 \beta U_{ном}}{p_{п} n_{ном} I_{ном}}",
                 rf"\frac{{9,55 \cdot {fmt(beta)} \cdot {fmt(U)}}}{{{fmt(p_pol)} \cdot {fmt(n_nom)} \cdot {fmt(I_nom)}}}",
                 9.55 * beta * U / (p_pol * n_nom * I_nom), "Гн", 4, ref="п. 5")
    Te = log.step("tr5.Te", "Электромагнитная постоянная времени якоря", "T_{э}", r"\frac{L}{R_{я}}",
                  rf"\frac{{{fmt(L)}}}{{{fmt(R)}}}", L / R, "с", 4, ref="п. 5")
    Tm = log.step("tr5.Tm", "Механическая постоянная времени (nу = nном, Mик = Mп − Mс)", "T_{М}",
                  r"\frac{J n_{ном}}{9,55 (M_{п} - M_{ном})}",
                  rf"\frac{{{fmt(J)} \cdot {fmt(n_nom)}}}{{9,55 \cdot ({fmt(M_p, 5)} - {fmt(M_nom, 5)})}}",
                  J * n_nom / (9.55 * (M_p - M_nom)), "с", 4, ref="п. 5")
    assert abs(Tm - J * n0 / (9.55 * M_p)) / Tm < 1e-9     # = J·n0/(9,55·Mп): характеристика прямая

    dt1 = log.step("tr5.dt", "Длительность 1-й части (якорь неподвижен, I растёт до Iном)", r"\Delta t",
                   r"-T_{э} \ln\left(1 - \frac{I_{ном} R_{я}}{U_{ном}}\right)",
                   rf"-{fmt(Te)} \cdot \ln\left(1 - \frac{{{fmt(I_nom)} \cdot {fmt(R)}}}{{{fmt(U)}}}\right)",
                   -Te * math.log(1 - I_nom * R / U), "с", 4, ref="п. 5",
                   note="из I = U/R·(1 − e^(−t/Tэ)) при I = Iном (M = Mс)")
    D = log.step("tr5.D", "Дискриминант характеристического уравнения", "D",
                 r"\frac{1}{4 T_{э}^2} - \frac{1}{T_{э} T_{М}}",
                 rf"\frac{{1}}{{4 \cdot {fmt(Te)}^2}} - \frac{{1}}{{{fmt(Te)} \cdot {fmt(Tm)}}}",
                 1 / (4 * Te ** 2) - 1 / (Te * Tm), "1/с²", 4, ref="п. 5")
    out = dict(M_c=M_c, n_c=n_c, L=L, Te=Te, Tm=Tm, dt1=dt1, D=D, M_p=M_p, I_p=I_p, n0=n0,
               n_nom=n_nom, I_nom=I_nom, M_nom=M_nom, U=U, R=R, J=J)
    sq = math.sqrt(Te * Tm)
    alpha = 1 / (2 * Te)
    if D < 0:
        out["kind"] = "колебательный"
        log.text("Корни комплексные (D < 0) — колебательный процесс; расчёт по формулам Эйлера",
                 r"p_{1,2} = -\frac{1}{2 T_{э}} \pm j \nu")
        al = log.step("tr5.alpha", "Показатель затухания (Re p)", r"\delta", r"\frac{1}{2 T_{э}}",
                      rf"\frac{{1}}{{2 \cdot {fmt(Te)}}}", alpha, "1/с", 4, ref="п. 5")
        nu = log.step("tr5.nu", "Угловая частота колебаний", r"\nu",
                      r"\sqrt{\frac{1}{T_{э} T_{М}} - \frac{1}{4 T_{э}^2}}",
                      rf"\sqrt{{\frac{{1}}{{{fmt(Te)} \cdot {fmt(Tm)}}} - \frac{{1}}{{4 \cdot {fmt(Te)}^2}}}}",
                      math.sqrt(1 / (Te * Tm) - 1 / (4 * Te ** 2)), "рад/с", 4, ref="п. 5")
        A = log.step("tr5.A", "Амплитуда n", "A", r"\frac{n_{ном}}{\nu \sqrt{T_{э} T_{М}}}",
                     rf"\frac{{{fmt(n_nom)}}}{{{fmt(nu)} \cdot \sqrt{{{fmt(Te)} \cdot {fmt(Tm)}}}}}",
                     n_nom / (nu * sq), "об/мин", 4, ref="п. 5")
        phi0 = log.step("tr5.phi0", "Фаза n: главное значение арксинуса (формула методички)", r"\varphi_0",
                        r"\arcsin\left(-\nu \sqrt{T_{э} T_{М}}\right)",
                        rf"\arcsin\left(-{fmt(nu)} \cdot \sqrt{{{fmt(Te)} \cdot {fmt(Tm)}}}\right)",
                        math.asin(-nu * sq), "рад", 4, ref="п. 5")
        phi = log.step("tr5.phi", "Фаза n: второй корень sin φ = −ν√(TэTМ), при котором n'(0) = 0", r"\varphi",
                       r"-\pi - \varphi_0", rf"-\pi - ({fmt(phi0)})", -math.pi - phi0, "рад", 4, ref="п. 5",
                       note="arcsin даёт только главное значение; якорь трогается без начальной скорости "
                            "(n(0)=0 и n'(0)=0), этому удовлетворяет φ = −π − φ0 (знак cos φ отрицательный)")
        B = log.step("tr5.B", "Амплитуда I", "B", r"\frac{I_{п} - I_{ном}}{\nu T_{э}}",
                     rf"\frac{{{fmt(I_p, 5)} - {fmt(I_nom)}}}{{{fmt(nu)} \cdot {fmt(Te)}}}",
                     (I_p - I_nom) / (nu * Te), "А", 4, ref="п. 5")
        C = log.step("tr5.C", "Амплитуда M", "C", r"\frac{M_{п} - M_{ном}}{\nu T_{э}}",
                     rf"\frac{{{fmt(M_p, 5)} - {fmt(M_nom, 5)}}}{{{fmt(nu)} \cdot {fmt(Te)}}}",
                     (M_p - M_nom) / (nu * Te), "Н·м", 4, ref="п. 5")
        log.text("Начальные фазы тока и момента", r"\psi = 0;\quad \gamma = 0")
        f_n = lambda t: A * np.exp(-alpha * t) * np.sin(nu * t + phi) + n_nom
        f_I = lambda t: B * np.exp(-alpha * t) * np.sin(nu * t) + I_nom
        f_M = lambda t: C * np.exp(-alpha * t) * np.sin(nu * t) + M_nom
        w = 2 * Te * nu
        s1 = w / math.sqrt(1 + w * w)                            # sin(arctan(2·Te·ν))
        t_M = log.step("tr5.tMmax", "Время 1-го максимума M и I (после начала вращения)", "t_{M\\,max}",
                       r"\frac{\arctan(2 T_{э} \nu)}{\nu}",
                       rf"\frac{{\arctan(2 \cdot {fmt(Te)} \cdot {fmt(nu)})}}{{{fmt(nu)}}}",
                       math.atan(w) / nu, "с", 4, ref="п. 5", note="из dM/dt = 0: tg(νt) = 2·Tэ·ν")
        M_max = log.step("tr5.Mmax", "Максимальный момент в переходном процессе", "M_{max}",
                         r"M_{ном} + C e^{-\delta t_{M\,max}} \sin(\nu t_{M\,max})",
                         rf"{fmt(M_nom, 5)} + {fmt(C)} \cdot e^{{-{fmt(alpha)} \cdot {fmt(t_M)}}} \cdot \sin({fmt(nu * t_M)})",
                         float(f_M(t_M)), "Н·м", 5, ref="п. 5")
        I_max = log.step("tr5.Imax", "Максимальный ток в переходном процессе", "I_{max}",
                         r"I_{ном} + B e^{-\delta t_{M\,max}} \sin(\nu t_{M\,max})",
                         rf"{fmt(I_nom)} + {fmt(B)} \cdot e^{{-{fmt(alpha)} \cdot {fmt(t_M)}}} \cdot \sin({fmt(nu * t_M)})",
                         float(f_I(t_M)), "А", 5, ref="п. 5")
        t_n = (math.atan(w) - phi) / nu
        if t_n <= 0:
            t_n += math.pi / nu
        n_max = log.step("tr5.nmax", "Максимальная частота вращения (1-й заброс)", "n_{max}",
                         r"n_{ном} + A e^{-\delta t} \sin(\nu t + \varphi)",
                         rf"{fmt(n_nom)} + {fmt(A)} \cdot e^{{-{fmt(alpha)} \cdot {fmt(t_n)}}} \cdot \sin({fmt(nu * t_n + phi)})",
                         float(f_n(t_n)), "об/мин", 4, ref="п. 5",
                         note=f"при t = {fmt(t_n)} с (tg(νt+φ) = 2·Tэ·ν); n0 = {fmt(n0)} об/мин")
        sigma = log.step("tr5.sigma", "Перерегулирование по скорости", r"\sigma",
                         r"\frac{n_{max} - n_{ном}}{n_{ном}} \cdot 100",
                         rf"\frac{{{fmt(n_max)} - {fmt(n_nom)}}}{{{fmt(n_nom)}}} \cdot 100",
                         (n_max - n_nom) / n_nom * 100, "%", 4, ref="п. 5")
        out.update(alpha=alpha, nu=nu, A=A, phi=phi, B=B, C=C, t_Mmax=t_M, M_max=M_max, I_max=I_max,
                   t_nmax=t_n, n_max=n_max, sigma_pct=sigma)
    else:
        out["kind"] = "апериодический"
        log.text("Корни вещественные (D ≥ 0) — апериодический процесс; формулы с гиперболическим синусом",
                 r"\varepsilon = \sqrt{\frac{1}{4 T_{э}^2} - \frac{1}{T_{э} T_{М}}}")
        eps = log.step("tr5.eps", "Параметр ε", r"\varepsilon", r"\sqrt{\frac{1}{4 T_{э}^2} - \frac{1}{T_{э} T_{М}}}",
                       rf"\sqrt{{{fmt(D)}}}", math.sqrt(D), "1/с", 4, ref="п. 5")
        A = log.step("tr5.A", "Амплитуда n", "A", r"\frac{n_{ном}}{\varepsilon \sqrt{T_{э} T_{М}}}",
                     rf"\frac{{{fmt(n_nom)}}}{{{fmt(eps)} \cdot \sqrt{{{fmt(Te)} \cdot {fmt(Tm)}}}}}",
                     n_nom / (eps * sq), "об/мин", 4, ref="п. 5")
        phi = log.step("tr5.phi", "Начальная фаза n", r"\varphi", r"\mathrm{arsh}\left(-\varepsilon \sqrt{T_{э} T_{М}}\right)",
                       rf"\mathrm{{arsh}}\left(-{fmt(eps)} \cdot \sqrt{{{fmt(Te)} \cdot {fmt(Tm)}}}\right)",
                       math.asinh(-eps * sq), "рад", 4, ref="п. 5")
        B = log.step("tr5.B", "Амплитуда I", "B", r"\frac{I_{п} - I_{ном}}{\varepsilon T_{э}}",
                     rf"\frac{{{fmt(I_p, 5)} - {fmt(I_nom)}}}{{{fmt(eps)} \cdot {fmt(Te)}}}",
                     (I_p - I_nom) / (eps * Te), "А", 4, ref="п. 5")
        C = log.step("tr5.C", "Амплитуда M", "C", r"\frac{M_{п} - M_{ном}}{\varepsilon T_{э}}",
                     rf"\frac{{{fmt(M_p, 5)} - {fmt(M_nom, 5)}}}{{{fmt(eps)} \cdot {fmt(Te)}}}",
                     (M_p - M_nom) / (eps * Te), "Н·м", 4, ref="п. 5")
        f_n = lambda t: A * np.exp(-alpha * t) * np.sinh(eps * t + phi) + n_nom
        f_I = lambda t: B * np.exp(-alpha * t) * np.sinh(eps * t) + I_nom
        f_M = lambda t: C * np.exp(-alpha * t) * np.sinh(eps * t) + M_nom
        out.update(alpha=alpha, eps=eps, A=A, phi=phi, B=B, C=C)

    t_set = log.step("tr5.tset", "Время переходного процесса (затухание огибающей до 5 %)", "t_{пп}",
                     r"2 T_{э} \ln 20", rf"2 \cdot {fmt(Te)} \cdot \ln {SETTLE}",
                     2 * Te * math.log(SETTLE), "с", 4, ref="п. 5")
    out["t_set"] = t_set

    # --- проверка: численное решение той же линейной модели (RK4) ---
    #   Tэ·dM/dt = Mп(1 − n/n0) − M;   (J/9,55)·dn/dt = M − Mс;   M(0) = Mс, n(0) = 0
    dt = 2e-5
    steps = int(1.2 * t_set / dt)
    x = np.array([M_nom, 0.0])

    def rhs(v):
        return np.array([(M_p * (1 - v[1] / n0) - v[0]) / Te, 9.55 * (v[0] - M_c) / J])

    tt = np.arange(steps + 1) * dt
    num = np.empty((steps + 1, 2))
    num[0] = x
    for k in range(steps):
        k1 = rhs(x)
        k2 = rhs(x + dt / 2 * k1)
        k3 = rhs(x + dt / 2 * k2)
        k4 = rhs(x + dt * k3)
        x = x + dt / 6 * (k1 + 2 * k2 + 2 * k3 + k4)
        num[k + 1] = x
    dev_n = float(np.max(np.abs(f_n(tt) - num[:, 1]))) / n_nom * 100
    dev_M = float(np.max(np.abs(f_M(tt) - num[:, 0]))) / float(np.max(np.abs(f_M(tt) - M_nom))) * 100
    print(f"      Проверка численным решением (RK4): max|Δn|/nном = {dev_n:.2f} %, "
          f"max|ΔM|/амплитуды = {dev_M:.2f} %")
    out.update(dev_n_pct=dev_n, dev_M_pct=dev_M)

    # --- кривые для графика (в одной шкале времени: 1-я часть + 2-я часть) ---
    t1 = np.linspace(0, dt1, 60)
    t2 = np.linspace(0, SETTLE_SHOW * t_set, 6000)
    curves = dict(
        t=np.concatenate([t1, dt1 + t2]),
        n=np.concatenate([np.zeros_like(t1), f_n(t2)]),
        M=np.concatenate([M_p * (1 - np.exp(-t1 / Te)) * (M_nom / (M_p * (1 - math.exp(-dt1 / Te)))), f_M(t2)]),
        I=np.concatenate([U / R * (1 - np.exp(-t1 / Te)), f_I(t2)]))
    return out, curves


def plot_dc_transient(tr, curves, outdir):
    """Рис. 5 методички: а) механическая характеристика; в) колебательный (или б) апериодический) процесс."""
    n0, n_c, M_c, M_p, dt1 = tr["n0"], tr["n_c"], tr["M_c"], tr["M_p"], tr["dt1"]
    t, n, M = curves["t"], curves["n"], curves["M"]
    fig = plt.figure(figsize=(14, 8))
    gs = fig.add_gridspec(2, 2, width_ratios=[1, 1.5], hspace=0.12, wspace=0.28)
    arrow = dict(arrowstyle="-|>", color="k", lw=1.2, shrinkA=0, shrinkB=0)

    # а) механическая характеристика
    a = fig.add_subplot(gs[:, 0])
    a.plot([0, M_p], [n0, 0], "k", lw=1.6)
    a.plot(M, n, color="0.35", lw=1.0)   # динамическая характеристика n(M) за время переходного процесса
    a.annotate("динамическая\nхарактеристика", xy=(float(M[np.argmax(n)]), float(n.max())), xytext=(0.45 * M_p, 0.95 * n.max()), fontsize=9, color="0.3", arrowprops=dict(arrowstyle="->", lw=0.7, color="0.3"))
    a.plot([0, M_p], [n0, n0], color="gray", ls=(0, (5, 4)), lw=0.9)
    a.plot([0, M_p], [n_c, n_c], color="gray", ls=(0, (5, 4)), lw=0.9)
    a.plot([M_c, M_c], [0, n0], "k", lw=1.0)
    a.plot([M_c], [n_c], "ko", ms=4)
    a.set_xlim(0, 1.08 * M_p)
    a.set_ylim(0, 1.12 * max(n0, n.max()))
    for sp in ("top", "right"):
        a.spines[sp].set_visible(False)
    a.set_yticks([n_c, n0])
    a.set_yticklabels([])
    a.text(-0.01 * M_p, n0 + 0.025 * n0, f"$n_0$ = {fmt(n0, 4)}", ha="right", va="center", fontsize=11)
    a.text(-0.01 * M_p, n_c - 0.045 * n0, f"$n_c$ = {fmt(n_c, 4)}", ha="right", va="center", fontsize=11)
    a.set_xticks([M_c, M_p])
    a.set_xticklabels([f"$M_с$ = {fmt(M_c, 3)}", f"$M_п$ = {fmt(M_p, 5)}"], fontsize=10)
    a.annotate("$n$, об/мин", xy=(0, 1), xycoords="axes fraction", xytext=(-6, 10), textcoords="offset points",
               ha="right", va="bottom", fontsize=12)
    a.annotate("$M$, Н·м", xy=(1, 0), xycoords="axes fraction", xytext=(8, 0), textcoords="offset points",
               ha="left", va="center", fontsize=12)
    a.set_title("а)", loc="left", fontsize=12, style="italic")

    # в) переходный процесс: n(t), M(t) на общей оси времени
    b = fig.add_subplot(gs[0, 1])
    c = fig.add_subplot(gs[1, 1], sharex=b)
    b.plot(t, n, "k", lw=1.4)
    b.axhline(n0, color="gray", ls=(0, (5, 4)), lw=0.9)
    b.axhline(n_c, color="gray", ls="-", lw=0.7)
    b.set_ylim(min(0, n.min()) - 0.02 * n.max(), 1.08 * n.max())
    b.set_yticks([n_c])
    b.set_yticklabels([])
    b.text(-0.012 * t.max(), n_c - 0.05 * n.max(), f"$n_c$ = {fmt(n_c, 4)}", ha="right", va="center", fontsize=10)
    b.text(-0.012 * t.max(), n0 + 0.05 * n.max(), f"$n_0$ = {fmt(n0, 4)}", ha="right", va="center", fontsize=10)
    b.text(0.55 * t.max(), 0.12 * n.max(), "$n = f(t)$", fontsize=12)
    c.plot(t, M, "k", lw=1.4)
    c.axhline(M_c, color="k", lw=0.7)
    ymax = 1.1 * M.max()
    ymin = min(M.min(), 0) * 1.1
    c.set_ylim(ymin, ymax)
    c.set_yticks([M_c])
    c.set_yticklabels([])
    c.text(-0.012 * t.max(), M_c, f"$M_с$ = {fmt(M_c, 3)}", ha="right", va="center", fontsize=10)
    c.annotate(f"$M_п$ = {fmt(M_p, 5)} Н·м — выше шкалы", xy=(0.02 * t.max(), ymax), xytext=(0.12 * t.max(), ymax * 0.97),
               ha="left", va="top", fontsize=10, arrowprops=dict(arrowstyle="->", lw=0.8))
    c.text(0.30 * t.max(), 0.55 * ymax, "$M = f(t)$", fontsize=12)
    for ax in (b, c):
        ax.axvline(dt1, color="gray", ls=(0, (5, 4)), lw=0.8)
        ax.axvline(dt1 + tr["t_set"], color="gray", ls=(0, (5, 4)), lw=0.8)
        ax.axvspan(dt1 + tr["t_set"], t.max(), color="0.94", zorder=0)
        for sp in ("top", "right"):
            ax.spines[sp].set_visible(False)
        ax.grid(True, alpha=0.25)
    plt.setp(b.get_xticklabels(), visible=False)
    b.text(dt1 + tr["t_set"] + 0.02 * t.max(), 0.22 * n.max(), "установив-\nшийся\nрежим", fontsize=9, color="0.25", va="center")
    b.text(0.5 * (dt1 + tr["t_set"]), 0.30 * n.max(), "переходный\nпроцесс", fontsize=9, color="0.25", ha="center", va="center")
    c.annotate(r"$t_{пп}$", xy=(dt1 + tr["t_set"], 0), xycoords=c.get_xaxis_transform(), xytext=(0, -36), textcoords="offset points", ha="center", fontsize=11)
    c.set_xlim(0, t.max())
    c.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _: fmt(x, 3)))
    c.annotate(r"$\Delta t$", xy=(dt1, ymin), xytext=(dt1 + 0.03 * t.max(), ymin + 0.05 * (ymax - ymin)),
               ha="left", fontsize=11, arrowprops=dict(arrowstyle="->", lw=0.8))
    c.annotate("$t$, с", xy=(1, 0), xycoords="axes fraction", xytext=(8, 0), textcoords="offset points",
               ha="left", va="center", fontsize=12)
    b.annotate("$n$, об/мин", xy=(0, 1), xycoords="axes fraction", xytext=(-6, 8), textcoords="offset points",
               ha="right", va="bottom", fontsize=12)
    c.annotate("$M$, Н·м", xy=(0, 1), xycoords="axes fraction", xytext=(-6, 8), textcoords="offset points",
               ha="right", va="bottom", fontsize=12)
    lab = "в)" if tr["kind"] == "колебательный" else "б)"
    b.set_title(lab, loc="left", fontsize=12, style="italic")
    fig.savefig(outdir / "fig_5_1.png", dpi=300, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------- main
def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    dc, ac = DC_MOTORS[VARIANT], AC_MOTORS[VARIANT]
    print(f"Вариант {VARIANT}: ДПТ {dc.type}, АД {ac.type}")

    dc_nat, dc_arts = dc_run(dc, VARIANT)
    ac_res, ac_arts = ac_run(ac, VARIANT)
    plot_dc(dc, dc_nat, dc_arts, OUT_DIR)
    plot_ac(ac, ac_res, ac_arts, OUT_DIR)
    st = ac_start_diagram(ac, ac_res)
    tr = ac_transient(ac, ac_res, st)
    plot_start_diagram(st, OUT_DIR)
    plot_transient(st, tr, OUT_DIR, ac.n_nom)
    dc5, dc5_curves = dc_transient(dc, dc_nat)
    plot_dc_transient(dc5, dc5_curves, OUT_DIR)
    import csv
    with open(OUT_DIR / "table_4.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Ступень", "n_нач", "n_кон", "n_уст", "T_М", "t"])
        for x in tr["stages"]:
            w.writerow([x["i"]] + [str(round(x[k], 4)).replace(".", ",") for k in ("n_start", "n_end", "n_y", "T", "t")])

    results = {
        "variant": VARIANT,
        "dc": {"motor": dc.__dict__, "natural": dc_nat, "artificial": dc_arts},
        "ac": {"motor": ac.__dict__, **ac_res, "artificial": ac_arts},
        "start": st,
        "transient": tr,
        "dc_transient": dc5,
    }
    (OUT_DIR / "results.json").write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    log.save(OUT_DIR)
    print(f"\nГотово: {OUT_DIR}  (шагов расчёта: {log.n_step}, данных: {log.n_data})")


if __name__ == "__main__":
    main()
