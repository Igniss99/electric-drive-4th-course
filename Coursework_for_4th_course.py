"""Курсовая работа «Электрический привод» (ГУМРФ). Разделы 1–2: характеристики ДПТ и АД.

Запуск:  python Coursework_for_4th_course.py
Выход:   out/variant_<N>/results.json, fig_*.png
Нумерация пунктов — по методичке (Metodichka.docx).
"""
import json
import math
from dataclasses import dataclass
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

VARIANT = 10
OUT_DIR = Path(__file__).parent / "out" / f"variant_{VARIANT}"

# ---------------------------------------------------------------- константы
F1 = 50.0          # частота сети, Гц
U_LINE = 380.0     # линейное напряжение АД, В (Прил. 3)
U_PHASE = 220.0    # фазное напряжение, В (принято в образце: 380/√3 ≈ 220)
M_PHASES = 3       # число фаз
K_SAFETY = 0.95    # коэффициент 0,95 в k = 0,95·Uф/E2ф (методичка, п. 2.1)
SYNC_STEP = 3000   # ряд синхронных скоростей: 60·f1/pп, об/мин


def rpm_to_rad(n):
    return 2 * math.pi * n / 60


def rad_to_rpm(w):
    return 60 * w / (2 * math.pi)


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


DC_MOTORS = {10: DCMotor("Д-812", 38, 565, 192, 0.0140, 7.0, 220)}
AC_MOTORS = {10: ACMotor("4МТН225L6", 55, 970, 0.039, 0.13, 0.029, 0.135, 290, 1.02)}

# Прил. 2 и 4: множители по группам вариантов (7–12 → 2-я группа).
# Здесь только группа 7–12; остальные группы добавить при расчёте других вариантов.
DC_ART = {  # (название, Ф, R, U) — множители к номинальным
    "7-12": [("Ф = 0,8Фном", 0.8, 1.0, 1.0),
             ("R = 2,0Rя", 1.0, 2.0, 1.0),
             ("U = 1,2Uном", 1.0, 1.0, 1.2)],
}
AC_ART = {  # (название, ключ параметра, множитель)
    "7-12": [("U = 0,8Uном", "U", 0.8),
             ("r1 = 2,0r1", "r1", 2.0),
             ("xкз = 2,0xкз", "xkz", 2.0),
             ("f1 = 2,0f1", "f1", 2.0),
             ("pп = 3pп", "p", 3.0)],
}


def group_of(variant):
    for lo, hi in [(1, 6), (7, 12), (13, 18), (19, 24), (25, 29)]:
        if lo <= variant <= hi:
            return f"{lo}-{hi}"


# ---------------------------------------------------------------- раздел 1
def dc_natural(m: DCMotor):
    """п. 1.1: естественные характеристики ДПТ."""
    P = m.P_kW * 1e3
    w_nom = rpm_to_rad(m.n_nom)
    M_nom = P / w_nom
    cPhi = (m.U_nom - m.I_nom * m.R_ya) / m.n_nom        # В·мин/об
    kPhi = M_nom / m.I_nom                               # Н·м/А (по методичке)
    n0 = m.U_nom / cPhi
    I_p = m.U_nom / m.R_ya
    M_p = I_p * kPhi
    return dict(P=P, w_nom=w_nom, M_nom=M_nom, cPhi_nom=cPhi, kPhi_nom=kPhi,
                n0=n0, I_p=I_p, M_p=M_p)


def dc_artificial(m: DCMotor, nat, phi, rk, uk):
    """п. 1.2: искусственная характеристика при изменённых Ф, R, U."""
    U = uk * m.U_nom
    R = rk * m.R_ya
    n0 = U / (nat["cPhi_nom"] * phi)
    I_p = U / R
    M_p = U * nat["kPhi_nom"] * phi / R
    return dict(U=U, R=R, phi=phi, n0=n0, I_p=I_p, M_p=M_p)


def dc_run(m: DCMotor, variant):
    nat = dc_natural(m)
    arts = []
    for name, phi, rk, uk in DC_ART[group_of(variant)]:
        arts.append(dict(name=name, **dc_artificial(m, nat, phi, rk, uk)))
    # проверки
    n_check = (m.U_nom - m.I_nom * m.R_ya) / nat["cPhi_nom"]
    assert abs(n_check - m.n_nom) < 1e-6
    assert abs(nat["M_p"] / nat["M_nom"] - nat["I_p"] / m.I_nom) < 1e-9
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
    """pп по n0: ближайшая синхронная скорость выше nном, кратная 3000/pп."""
    p = 1
    while 60 * f1 / p > n_nom:
        p += 1
    return p - 1


def ac_params(m: ACMotor):
    """Приведение ротора к статору, xкз, a (п. 2.1)."""
    E2f = m.E2k / math.sqrt(3)
    k = K_SAFETY * U_PHASE / E2f
    r2p = k ** 2 * m.r2
    x2p = k ** 2 * m.x2
    return dict(E2f=E2f, k=k, r2p=r2p, x2p=x2p, xkz=m.x1 + x2p)


def ac_klos(U, r1, r2p, xkz, f1, p):
    """Mmax, sk (двигательный и генераторный), n0, a — по формулам п. 2.1."""
    n0 = 60 * f1 / p
    w0 = 2 * math.pi * f1 / p
    root = math.hypot(r1, xkz)
    Mmax_d = M_PHASES * U ** 2 / (2 * w0 * (r1 + root))
    Mmax_g = M_PHASES * U ** 2 / (2 * w0 * (r1 - root))        # < 0
    sk_d = r2p / root
    sk_g = -sk_d
    a = r1 / r2p
    return dict(n0=n0, w0=w0, Mmax_d=Mmax_d, Mmax_g=Mmax_g, sk_d=sk_d, sk_g=sk_g, a=a)


def klos(s, Mmax, sk, a):
    """M(s) по формуле Клосса с учётом r1 (п. 2.1)."""
    s = np.asarray(s, dtype=float)
    return 2 * Mmax * (1 + a * sk) / (s / sk + sk / s + 2 * a * sk)


def ac_char(U, r1, xkz, f1, p, r2p):
    c = ac_klos(U, r1, r2p, xkz, f1, p)
    c["M_start"] = float(klos(1.0, c["Mmax_d"], c["sk_d"], c["a"]))
    return c


def ac_run(m: ACMotor, variant):
    prm = ac_params(m)
    p = ac_pole_pairs(m.n_nom)
    base = dict(U=U_PHASE, r1=m.r1, xkz=prm["xkz"], f1=F1, p=p)
    nat = ac_char(base["U"], base["r1"], base["xkz"], base["f1"], base["p"], prm["r2p"])
    s_nom = (nat["n0"] - m.n_nom) / nat["n0"]
    M_nom = float(klos(s_nom, nat["Mmax_d"], nat["sk_d"], nat["a"]))
    arts = []
    for name, key, mult in AC_ART[group_of(variant)]:
        q = dict(base)
        q[key] = base[key] * mult
        c = ac_char(q["U"], q["r1"], q["xkz"], q["f1"], int(round(q["p"])), prm["r2p"])
        arts.append(dict(name=name, key=key, mult=mult, **c))
    # проверка Клосса: в sk момент равен Mmax
    assert abs(float(klos(nat["sk_d"], nat["Mmax_d"], nat["sk_d"], nat["a"])) - nat["Mmax_d"]) < 1e-6
    return dict(params=prm, p=p, s_nom=s_nom, M_nom_klos=M_nom, natural=nat), arts


def plot_ac(m, res, arts, outdir):
    nat = res["natural"]
    s = np.concatenate([np.linspace(-2, -1e-3, 800), np.linspace(1e-3, 2, 800)])

    def curve(c):
        M = np.where(s > 0, klos(s, c["Mmax_d"], c["sk_d"], c["a"]),
                     klos(s, c["Mmax_g"], c["sk_g"], c["a"]))
        return M, c["n0"] * (1 - s)

    def fmt(ax):
        ax.set(xlabel="M, Н·м", ylabel="n, об/мин")
        ax.axhline(0, color="gray", lw=.5)
        ax.axvline(0, color="gray", lw=.5)
        ax.grid(True, alpha=.4)
        ax.legend()

    fig, ax = plt.subplots(figsize=(7, 5))
    ax.plot(*curve(nat), "k", label="естественная")
    fmt(ax)
    fig.tight_layout()
    fig.savefig(outdir / "fig_2_1.png", dpi=300)
    plt.close(fig)

    for i, art in enumerate(arts, 1):
        fig, ax = plt.subplots(figsize=(7, 5))
        ax.plot(*curve(nat), "k", label="естественная")
        ax.plot(*curve(art), "C3", label="искусственная: " + art["name"])
        fmt(ax)
        fig.tight_layout()
        fig.savefig(outdir / f"fig_2_2_{i}.png", dpi=300)
        plt.close(fig)


# ---------------------------------------------------------------- main
def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    dc, ac = DC_MOTORS[VARIANT], AC_MOTORS[VARIANT]

    dc_nat, dc_arts = dc_run(dc, VARIANT)
    ac_res, ac_arts = ac_run(ac, VARIANT)
    plot_dc(dc, dc_nat, dc_arts, OUT_DIR)
    plot_ac(ac, ac_res, ac_arts, OUT_DIR)

    results = {
        "variant": VARIANT,
        "dc": {"motor": dc.__dict__, "natural": dc_nat, "artificial": dc_arts},
        "ac": {"motor": ac.__dict__, **ac_res, "artificial": ac_arts},
    }
    (OUT_DIR / "results.json").write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Вариант {VARIANT}: {dc.type}, {ac.type}")
    print("ДПТ естеств.:", {k: round(v, 4) for k, v in dc_nat.items()})
    for a in dc_arts:
        print("  ", a["name"], {k: round(v, 2) for k, v in a.items() if k != "name"})
    print("АД параметры:", {k: round(v, 4) for k, v in ac_res["params"].items()}, "pп =", ac_res["p"])
    print("АД естеств.:", {k: round(v, 4) for k, v in ac_res["natural"].items()})
    for a in ac_arts:
        print("  ", a["name"], {k: round(v, 4) for k, v in a.items() if k not in ("name", "key")})


if __name__ == "__main__":
    main()
