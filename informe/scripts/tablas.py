"""Genera figuras/tablas-datos.tex a partir de Practica3_pendulos_acoplados.xlsx.

Las tablas del informe se escriben a partir de la hoja de cálculo para que
no haya errores de transcripción. Uso:  python3 scripts/tablas.py
"""
import math
import pathlib
import openpyxl

RAIZ = pathlib.Path(__file__).resolve().parents[2]
XLSX = RAIZ / "Practica3_pendulos_acoplados.xlsx"
OUT = RAIZ / "informe" / "tablas"
ws = openpyxl.load_workbook(XLSX, data_only=True).active

LEN = [25, 35, 45, 55, 65]
STARTS = ["E", "J", "O", "T", "Y"]
ci = openpyxl.utils.column_index_from_string
cl = openpyxl.utils.get_column_letter


def serie(start, k, r0, r1):
    c0 = ci(start)
    pc, tc = cl(c0 + k - 1), cl(c0 + k)
    pts = [(ws[f"{pc}{r}"].value, ws[f"{tc}{r}"].value) for r in range(r0, r1 + 1)]
    return [(p, t) for p, t in pts if t is not None]


def num(x, d=0):
    s = f"{x:.{d}f}".replace("-", "−")
    return s


def prange(ps):
    lo, hi = min(ps), max(ps)
    f = lambda v: f"{v:.2f}".replace(".", ",").replace("-", "$-$")
    return f(lo) if lo == hi else f"{f(hi)} a {f(lo)}"


def celda_t(t, marca=False):
    s = f"\\num{{{t:.0f}}}"
    return s + "\\rlap{$^{\\dagger}$}" if marca else s


lines = ["% Generado por scripts/tablas.py a partir de Practica3_pendulos_acoplados.xlsx. No editar a mano.", ""]

# ---------------------------------------------------------------- modos 1 y 2
def tabla_modo(nombre, filas_datos, longitudes, etiqueta, titulo):
    out = [f"\\begin{{table}}[ht]", "\\centering", f"\\caption{{{titulo}}}", f"\\label{{{etiqueta}}}",
           "\\renewcommand{\\arraystretch}{1.18}\\setlength{\\tabcolsep}{4.2pt}",
           "\\begin{tabular}{@{}l c *{6}{r} c r@{}}", "\\toprule",
           "$\\ell$ & Péndulo & \\multicolumn{6}{c}{Tiempo de los máximos, $t_{i}$ (\\si{\\milli\\second})} & $p$ (\\si{\\volt}) & $T$ (\\si{\\milli\\second}) \\\\",
           "\\cmidrule(lr){3-8}",
           " & & $t_1$ & $t_2$ & $t_3$ & $t_4$ & $t_5$ & $t_6$ & & \\\\", "\\midrule"]
    for li in longitudes:
        for k, lado in ((1, 1), (3, 2)):
            pts = serie(STARTS[li], k, *filas_datos)
            ts = [t for _, t in pts]
            ps = [p for p, _ in pts]
            T = (ts[-1] - ts[0]) / (len(ts) - 1)
            marca = [(nombre == 1 and li == 1 and lado == 2 and i == 3) for i in range(6)]
            cel = " & ".join(celda_t(t, m) for t, m in zip(ts, marca))
            lab = f"\\multirow{{2}}{{*}}{{$\\ell_{{{li + 1}}}=\\SI{{{LEN[li]}}}{{\\centi\\meter}}$}}" if lado == 1 else ""
            out.append(f"{lab} & {lado} & {cel} & {prange(ps)} & \\num{{{T:.1f}}} \\\\")
        if li != longitudes[-1]:
            out.append("\\addlinespace[2pt]")
    out += ["\\bottomrule", "\\end{tabular}"]
    return out


lines.append("%%FILE modo1")
lines += tabla_modo(1, (3, 8), [0, 1], "tab:modo1",
                    "Modo en fase: instantes de los seis primeros máximos de cada péndulo, "
                    "altura de los picos y periodo $T=(t_6-t_1)/5$. Este modo solo se registró para "
                    "$\\ell_1$ y $\\ell_2$, pues $\\omega_0$ no depende de $\\ell$.")
lines += ["\\par\\vspace{3pt}{\\footnotesize\\color{grispizarra}$^{\\dagger}$\\,En el cuaderno de laboratorio "
          "figuraba \\SI{19824}{\\milli\\second}, errata de transcripción (cifras intercambiadas): con ese valor los "
          "intervalos vecinos serían de \\num{2568} y \\SI{1468}{\\milli\\second}, frente a los "
          "$\\approx\\SI{2000}{\\milli\\second}$ del resto.\\par}", "\\end{table}", ""]
lines.append("%%FILE modo2")
lines += tabla_modo(2, (13, 18), [0, 1, 2, 3, 4], "tab:modo2",
                    "Modo en contrafase: instantes de los seis primeros máximos de cada péndulo, "
                    "altura de los picos y periodo $T=(t_6-t_1)/5$ para las cinco longitudes de acoplamiento.")
lines += ["\\end{table}", ""]

# ---------------------------------------------------------------- modo 3
ENV_USO = [(0, 1), (0, 1), (1, 2), (1, 3), (1, 3)]  # índices (primero, último) de los máximos usados
env_rows, port_rows, resumen = [], [], []
for li in range(5):
    env = serie(STARTS[li], 1, 24, 29)
    port = serie(STARTS[li], 3, 24, 29)
    te = [t for _, t in env]
    tp = [t for _, t in port]
    a, b = ENV_USO[li]
    Tb = (te[b] - te[a]) / (b - a)
    Tp = (tp[-1] - tp[0]) / 5
    cel = []
    for i in range(4):
        if i < len(te):
            s = f"\\num{{{te[i]:.0f}}}"
            cel.append(s if a <= i <= b else f"\\textcolor{{grispizarra!70}}{{({s})}}")
        else:
            cel.append("—")
    env_rows.append(f"$\\ell_{{{li + 1}}}=\\SI{{{LEN[li]}}}{{\\centi\\meter}}$ & " + " & ".join(cel)
                    + f" & {prange([p for p, _ in env])} & {b - a} & \\num{{{Tb:.0f}}} \\\\")
    port_rows.append(f"$\\ell_{{{li + 1}}}=\\SI{{{LEN[li]}}}{{\\centi\\meter}}$ & "
                     + " & ".join(f"\\num{{{t:.0f}}}" for t in tp)
                     + f" & {prange([p for p, _ in port])} & \\num{{{Tp:.1f}}} \\\\")
    resumen.append((Tb, Tp))

lines.append("%%FILE modo3")
lines += ["\\begin{table}[ht]", "\\centering",
          "\\caption{Modo pulsante. Arriba, instantes de los máximos de la envolvente y periodo de "
          "pulsación $T_{\\mathrm{b}}=(t_{\\text{últ}}-t_{\\text{pri}})/n$, con $n$ el número de intervalos "
          "empleados (entre paréntesis, el máximo descartado). Abajo, seis máximos consecutivos de la portadora "
          "y su periodo $T_{+}=(t_6-t_1)/5$.}",
          "\\label{tab:modo3}",
          "\\renewcommand{\\arraystretch}{1.18}\\setlength{\\tabcolsep}{4.6pt}",
          "\\begin{tabular}{@{}l *{4}{r} c c r@{}}", "\\toprule",
          "\\multicolumn{8}{@{}l}{\\scshape\\color{azulprof}Envolvente (máximos de amplitud)} \\\\",
          "$\\ell$ & \\multicolumn{4}{c}{$t_{i}$ (\\si{\\milli\\second})} & $p$ (\\si{\\volt}) & $n$ & $T_{\\mathrm{b}}$ (\\si{\\milli\\second}) \\\\",
          "\\cmidrule(lr){2-5}", " & $t_1$ & $t_2$ & $t_3$ & $t_4$ & & & \\\\", "\\midrule"]
lines += env_rows
lines += ["\\bottomrule", "\\end{tabular}", "\\par\\vspace{8pt}",
          "\\begin{tabular}{@{}l *{6}{r} c r@{}}", "\\toprule",
          "\\multicolumn{9}{@{}l}{\\scshape\\color{azulprof}Portadora (máximos consecutivos)} \\\\",
          "$\\ell$ & \\multicolumn{6}{c}{$t_{i}$ (\\si{\\milli\\second})} & $p$ (\\si{\\volt}) & $T_{+}$ (\\si{\\milli\\second}) \\\\",
          "\\cmidrule(lr){2-7}", " & $t_1$ & $t_2$ & $t_3$ & $t_4$ & $t_5$ & $t_6$ & & \\\\", "\\midrule"]
lines += port_rows
lines += ["\\bottomrule", "\\end{tabular}", "\\end{table}", ""]

# ---------------------------------------------------------------- resumen de periodos y frecuencias
lines.append("%%FILE frecuencias")
lines += ["\\begin{table}[ht]", "\\centering",
          "\\caption{Periodos y frecuencias angulares para cada longitud de acoplamiento. $T_1$ es la media de "
          "los dos péndulos en contrafase; $\\omega_1=2\\pi/T_1$, $\\omega_{-}=\\pi/T_{\\mathrm{b}}$ y "
          "$\\omega_{+}=2\\pi/T_{+}$. Las incertidumbres se tratan en la \\cref{sec:ajustes}.}",
          "\\label{tab:frecuencias}",
          "\\renewcommand{\\arraystretch}{1.22}\\setlength{\\tabcolsep}{5pt}",
          "\\begin{tabular}{@{}l c c c c c c c c@{}}", "\\toprule",
          "$\\ell$ & $\\ell^{2}$ (\\si{\\meter\\squared}) & $T_1$ (\\si{\\second}) & $\\omega_1$ (\\si{\\radian\\per\\second}) & "
          "$\\omega_1^{2}$ (\\si{\\radian\\squared\\per\\second\\squared}) & $T_{\\mathrm{b}}$ (\\si{\\second}) & "
          "$\\omega_{-}$ (\\si{\\radian\\per\\second}) & $T_{+}$ (\\si{\\second}) & $\\omega_{+}$ (\\si{\\radian\\per\\second}) \\\\",
          "\\midrule"]
for li in range(5):
    p1 = serie(STARTS[li], 1, 13, 18)
    p2 = serie(STARTS[li], 3, 13, 18)
    T1 = ((p1[-1][1] - p1[0][1]) / 5 + (p2[-1][1] - p2[0][1]) / 5) / 2 / 1000
    w1 = 2 * math.pi / T1
    Tb, Tp = resumen[li][0] / 1000, resumen[li][1] / 1000
    l2 = (LEN[li] / 100) ** 2
    lines.append(f"$\\ell_{{{li + 1}}}$ & \\num{{{l2:.4f}}} & \\num{{{T1:.4f}}} & \\num{{{w1:.4f}}} & \\num{{{w1**2:.3f}}} & "
                 f"\\num{{{Tb:.2f}}} & \\num{{{math.pi / Tb:.4f}}} & \\num{{{Tp:.4f}}} & \\num{{{2 * math.pi / Tp:.4f}}} \\\\")
lines += ["\\bottomrule", "\\end{tabular}", "\\end{table}", ""]

# ---------------------------------------------------------------- momento de inercia
T0 = []
for li in (0, 1):
    for k in (1, 3):
        pts = serie(STARTS[li], k, 3, 8)
        T0.append((pts[-1][1] - pts[0][1]) / 5 / 1000)
Tm = sum(T0) / 4
w0 = 2 * math.pi / Tm
I = 1 * 9.81 * 1 / w0 ** 2
rows = []
lines.append("%%FILE omega0")
for (li, lado), T in zip([(1, 1), (1, 2), (2, 1), (2, 2)], T0):
    rows.append(f"$\\ell_{li}=\\SI{{{LEN[li - 1]}}}{{\\centi\\meter}}$ & {lado} & \\num{{{T:.4f}}} & \\num{{{2 * math.pi / T:.4f}}} \\\\")
lines += ["\\begin{table}[ht]", "\\centering",
          "\\caption{Periodos del modo en fase y frecuencia propia de cada medida. La última fila recoge el "
          "periodo medio y la frecuencia $\\omega_0=2\\pi/\\mean{T}_0$ que empleamos en adelante.}",
          "\\label{tab:omega0}", "\\renewcommand{\\arraystretch}{1.22}",
          "\\begin{tabular}{@{}l c c c@{}}", "\\toprule",
          "Longitud & Péndulo & $T_0$ (\\si{\\second}) & $2\\pi/T_0$ (\\si{\\radian\\per\\second}) \\\\", "\\midrule"]
lines += rows
lines += ["\\midrule", f"\\multicolumn{{2}}{{@{{}}l}}{{\\scshape Media}} & \\num{{{Tm:.4f}}} & \\num{{{w0:.4f}}} \\\\",
          "\\bottomrule", "\\end{tabular}", "\\end{table}", ""]
# macros con los valores, para citarlos en el texto
lines.append("%%FILE valores")
lines += [f"\\newcommand{{\\valTcero}}{{\\num{{{Tm:.4f}}}}}", f"\\newcommand{{\\valWcero}}{{\\num{{{w0:.4f}}}}}",
          f"\\newcommand{{\\valWceroDos}}{{\\num{{{w0 ** 2:.4f}}}}}", f"\\newcommand{{\\valI}}{{\\num{{{I:.4f}}}}}"]
OUT.mkdir(exist_ok=True)
cab, actual, partes = lines[0], None, {}
for ln in lines[1:]:
    if ln.startswith("%%FILE "):
        actual = ln.split()[1]
        partes[actual] = [cab, ""]
    elif actual:
        partes[actual].append(ln)
for nombre, cont in partes.items():
    (OUT / f"{nombre}.tex").write_text("\n".join(cont).rstrip() + "\n", encoding="utf-8")
print("ok", OUT, "T0", Tm, "w0", w0, "I", I)
