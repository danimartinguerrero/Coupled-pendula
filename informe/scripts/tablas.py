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
    f = lambda v: f"{v:.2f}".replace("-", "$-$")
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
          "\\caption{Modo pulsante. Arriba, instantes de los nodos de la envolvente (amplitud nula, "
          "sobre la línea de equilibrio) y periodo de "
          "pulsación $T_{\\mathrm{b}}=(t_{\\text{últ}}-t_{\\text{pri}})/n$, con $n$ el número de intervalos "
          "empleados (entre paréntesis, el nodo descartado). Abajo, seis máximos consecutivos de la portadora "
          "y su periodo $T_{+}=(t_6-t_1)/5$.}",
          "\\label{tab:modo3}",
          "\\renewcommand{\\arraystretch}{1.18}\\setlength{\\tabcolsep}{4.6pt}",
          "\\begin{tabular}{@{}l *{4}{r} c c r@{}}", "\\toprule",
          "\\multicolumn{8}{@{}l}{\\scshape\\color{azulprof}Envolvente (nodos)} \\\\",
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

# ---------------------------------------------------------------- incertidumbres (sección 3.5)
RES_T = 4e-3                 # resolución temporal: todos los instantes son múltiplos de 4 ms
T95_3 = 3.182                # factor de Student al 95 % con nu = 3


def redondea(v, u):
    """Valor e incertidumbre con el criterio del informe: una cifra significativa
    (dos si la primera es un 1) y el valor a la misma posición decimal."""
    e = math.floor(math.log10(abs(u)))
    cifras = 2 if int(u / 10 ** e) == 1 else 1
    dec = max(cifras - 1 - e, 0)
    return f"{v:.{dec}f}", f"{u:.{dec}f}"


def pm(v, u):
    a, b = redondea(v, u)
    return f"\\num{{{a} \\pm {b}}}"


def ajuste(x, y):
    n = len(x)
    mx, my = sum(x) / n, sum(y) / n
    sxx = sum((xi - mx) ** 2 for xi in x)
    a = sum((xi - mx) * (yi - my) for xi, yi in zip(x, y)) / sxx
    b = my - a * mx
    res = [yi - (a * xi + b) for xi, yi in zip(x, y)]
    s = math.sqrt(sum(r * r for r in res) / (n - 2))
    r2 = 1 - sum(r * r for r in res) / sum((yi - my) ** 2 for yi in y)
    return a, s / math.sqrt(sxx), b, s * math.sqrt(1 / n + mx ** 2 / sxx), r2, res


uT = math.sqrt(2) / 5 * RES_T            # periodo con 6 máximos
uT1 = uT / math.sqrt(2)                  # media de los dos péndulos en contrafase
X, W1, UW1, WM, UWM, WP, UWP, TB, TP, T1S = [], [], [], [], [], [], [], [], [], []
for li in range(5):
    p1 = serie(STARTS[li], 1, 13, 18)
    p2 = serie(STARTS[li], 3, 13, 18)
    T1 = ((p1[-1][1] - p1[0][1]) / 5 + (p2[-1][1] - p2[0][1]) / 5) / 2 / 1000
    Tb, Tp = resumen[li][0] / 1000, resumen[li][1] / 1000
    n = ENV_USO[li][1] - ENV_USO[li][0]
    uTb = math.sqrt(2) * Tp / n          # u(t_nodo) ~ T+
    X.append((LEN[li] / 100) ** 2); T1S.append(T1); TB.append(Tb); TP.append(Tp)
    W1.append(2 * math.pi / T1); UW1.append(2 * math.pi * uT1 / T1 ** 2)
    WM.append(math.pi / Tb); UWM.append(math.pi * uTb / Tb ** 2)
    WP.append(2 * math.pi / Tp); UWP.append(2 * math.pi * uT / Tp ** 2)
W1C = [w ** 2 for w in W1]
UW1C = [2 * w * u for w, u in zip(W1, UW1)]

lines.append("%%FILE frecuencias")
lines += ["\\begin{table}[ht]", "\\centering",
          "\\caption{Periodos y frecuencias angulares para cada longitud de acoplamiento, con sus incertidumbres "
          "típicas (\\cref{sec:incertidumbres}). $T_1$ es la media de los dos péndulos en contrafase; "
          "$\\omega_1=2\\pi/T_1$, $\\omega_{-}=\\pi/T_{\\mathrm{b}}$ y $\\omega_{+}=2\\pi/T_{+}$.}",
          "\\label{tab:frecuencias}",
          "\\renewcommand{\\arraystretch}{1.25}\\setlength{\\tabcolsep}{3.6pt}\\small",
          "\\begin{tabular}{@{}l c c c c c c c c@{}}", "\\toprule",
          "$\\ell$ & $\\ell^{2}$ & $T_1$ & $\\omega_1$ & $\\omega_1^{2}$ & $T_{\\mathrm{b}}$ & $\\omega_{-}$ & $T_{+}$ & $\\omega_{+}$ \\\\",
          " & (\\si{\\meter\\squared}) & (\\si{\\second}) & (\\si{\\radian\\per\\second}) & "
          "(\\si{\\radian\\squared\\per\\second\\squared}) & (\\si{\\second}) & (\\si{\\radian\\per\\second}) & "
          "(\\si{\\second}) & (\\si{\\radian\\per\\second}) \\\\",
          "\\midrule"]
for li in range(5):
    lines.append(f"$\\ell_{{{li + 1}}}$ & \\num{{{X[li]:.4f}}} & \\num{{{T1S[li]:.4f}}} & {pm(W1[li], UW1[li])} & "
                 f"{pm(W1C[li], UW1C[li])} & \\num{{{TB[li]:.2f}}} & {pm(WM[li], UWM[li])} & "
                 f"\\num{{{TP[li]:.4f}}} & {pm(WP[li], UWP[li])} \\\\")
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
# ---------------------------------------------------------------- incertidumbre de omega0 e I (cuestión 1)
sT0 = math.sqrt(sum((t - Tm) ** 2 for t in T0) / 3)
uTm = math.sqrt((sT0 / 2) ** 2 + uT ** 2)
uw0 = 2 * math.pi * uTm / Tm ** 2
uI = 2 * I * uTm / Tm
G, M, LP = 9.81, 1.0, 1.0

# ---------------------------------------------------------------- ajustes lineales (cuestión 2)
FITS = {}
for clave, y in (("w1", W1C), ("wm", WM), ("wp", WP)):
    a, ua, b, ub, r2, res = ajuste(X, y)
    if clave == "w1":
        k = I * a / 2
        uk = math.hypot(I / 2 * ua, a / 2 * uI)
    else:
        k = 2 * M * G * LP * a / w0
        uk = 2 * M * G * LP / w0 * math.hypot(ua, a / w0 * uw0)
    FITS[clave] = dict(a=a, ua=ua, b=b, ub=ub, r2=r2, res=res, k=k, uk=uk, Uk=T95_3 * uk)

lines.append("%%FILE ajustes")
filas = [("w1", "$\\omega_1^{2}$ frente a $\\ell^{2}$", f"$\\omega_0^{{2}}=\\num{{{w0 ** 2:.2f}}}$"),
         ("wm", "$\\omega_{-}$ frente a $\\ell^{2}$", "$0$"),
         ("wp", "$\\omega_{+}$ frente a $\\ell^{2}$", f"$\\omega_0=\\num{{{w0:.3f}}}$")]
lines += ["\\begin{table}[ht]", "\\centering",
          "\\caption{Parámetros de los ajustes lineales $y = a\\,\\ell^2 + b$ y constante del muelle $k$ deducida de cada uno.}",
          "\\label{tab:ajustes_lineales}",
          "\\renewcommand{\\arraystretch}{1.3}\\setlength{\\tabcolsep}{4.5pt}\\small",
          "\\begin{tabular}{@{}l c c c c c c@{}}", "\\toprule",
          "Ajuste & $a\\pm u(a)$ & $b\\pm u(b)$ & Ordenada teórica & $R^{2}$ & $k\\pm u(k)$ & $k\\pm U(k)$ \\\\",
          " & & & & & (\\si{\\newton\\per\\meter}) & (\\si{\\newton\\per\\meter}) \\\\", "\\midrule"]
for clave, nombre, teo in filas:
    f = FITS[clave]
    lines.append(f"{nombre} & {pm(f['a'], f['ua'])} & {pm(f['b'], f['ub'])} & {teo} & \\num{{{f['r2']:.4f}}} & "
                 f"{pm(f['k'], f['uk'])} & {pm(f['k'], f['Uk'])} \\\\")
lines += ["\\bottomrule", "\\end{tabular}",
          "\\par\\vspace{3pt}{\\footnotesize\\color{grispizarra}Unidades: $a$ en "
          "\\si{\\radian\\squared\\per\\second\\squared\\per\\meter\\squared} para $\\omega_1^{2}$ y en "
          "\\si{\\radian\\per\\second\\per\\meter\\squared} para $\\omega_{\\pm}$; $b$ y la ordenada teórica, "
          "en las de $y$. $u$: incertidumbre típica; $U=t_{95,3}\\,u$, con $t_{95,3}=\\num{3.182}$.\\par}",
          "\\end{table}", ""]

# figuras de los ajustes, con residuos
ETIQ = {"w1": ("\\omega_1^{2}", "\\si{\\radian\\squared\\per\\second\\squared}", W1C, UW1C, "%.3f", "%.3f"),
        "wm": ("\\omega_{-}", "\\si{\\radian\\per\\second}", WM, UWM, "%.4f", "%.4f"),
        "wp": ("\\omega_{+}", "\\si{\\radian\\per\\second}", WP, UWP, "%.4f", "%.4f")}
NOMBRE = {"w1": "Uno", "wm": "Menos", "wp": "Mas"}
for clave, (sim, uni, y, uy, fa, fb) in ETIQ.items():
    f = FITS[clave]
    dm = "\\datos" + NOMBRE[clave]
    lines.append(f"%%FILE fig-ajuste-{clave}")
    tabla = "\n".join(f"      {x:.4f} {yy:.6f} {u:.6f} {r:.6f}" for x, yy, u, r in zip(X, y, uy, f["res"]))
    ecuacion = (fa % f["a"]) + "\\,x+" + (fb % f["b"])
    lines += [
        "\\begin{tikzpicture}",
        "  \\pgfplotstableread{",
        "      x y uy r",
        tabla,
        f"  }}{dm}",
        "  \\begin{groupplot}[group style={group size=1 by 2, vertical sep=6pt,",
        "                      xlabels at=edge bottom, xticklabels at=edge bottom},",
        "                    datos, width=0.82\\linewidth, xmin=0.03, xmax=0.46,",
        "                    xlabel={$\\ell^{2}$ (\\si{\\meter\\squared})},",
        "                    xtick={0.1,0.2,0.3,0.4}, xticklabel style={/pgf/number format/fixed},",
        "                    enlarge y limits=0.15,",
        "                    yticklabel style={/pgf/number format/fixed, /pgf/number format/precision=3}]",
        f"  \\nextgroupplot[height=5.4cm, ylabel={{${sim}$ ({uni})}}, legend pos=north west,",
        "                 legend style={font=\\scriptsize}]",
        f"    \\addplot[azulprof, line width=1.0pt, domain=0.035:0.455, samples=2] {{{f['a']:.6f}*x+{f['b']:.6f}}};",
        f"    \\addlegendentry{{Ajuste: $y={ecuacion}$, $R^{{2}}=\\num{{{f['r2']:.4f}}}$}}",
        f"    \\addplot[puntoA, error bars/.cd, y dir=both, y explicit,",
        "              error bar style={line width=0.6pt, azulprof}, error mark options={rotate=90, mark size=2pt, azulprof}]",
        f"      table[x=x, y=y, y error=uy] {{{dm}}};",
        "    \\addlegendentry{Datos experimentales}",
        "    \\node[anchor=south east, font=\\small, text=grispizarra] at (rel axis cs:0.99,0.03) {(a)};",
        f"  \\nextgroupplot[height=3.0cm, ylabel={{Residuo}}]",
        "    \\draw[grispizarra, line width=0.5pt, densely dashed] (axis cs:0.03,0) -- (axis cs:0.46,0);",
        "    \\node[anchor=north west, font=\\small, text=grispizarra] at (rel axis cs:0.01,0.97) {(b)};",
        f"    \\addplot[puntoA, error bars/.cd, y dir=both, y explicit,",
        "              error bar style={line width=0.6pt, azulprof}, error mark options={rotate=90, mark size=2pt, azulprof}]",
        f"      table[x=x, y=r, y error=uy] {{{dm}}};",
        "  \\end{groupplot}",
        "\\end{tikzpicture}"]

# coherencia entre modos: omega_- y omega_+ medidos frente a (omega_1 -+ omega_0)/2
lines.append("%%FILE fig-coherencia")
filas_c = []
for li in range(5):
    um = math.hypot(UW1[li], uw0) / 2
    filas_c.append(f"      {X[li]:.4f} {WM[li]:.6f} {UWM[li]:.6f} {(W1[li] - w0) / 2:.6f} {um:.6f} "
                   f"{WP[li]:.6f} {UWP[li]:.6f} {(W1[li] + w0) / 2:.6f}")
lines += [
    "\\begin{tikzpicture}",
    "  \\pgfplotstableread{",
    "      x wm uwm tm utm wp uwp tp",
    "\n".join(filas_c),
    "  }\\datoscoh",
    "  \\begin{groupplot}[group style={group size=2 by 1, horizontal sep=1.6cm},",
    "                    datos, width=0.47\\textwidth, height=5.4cm, xmin=0.03, xmax=0.46,",
    "                    xlabel={$\\ell^{2}$ (\\si{\\meter\\squared})},",
    "                    legend pos=north west, legend style={font=\\scriptsize},",
    "                    xtick={0.1,0.2,0.3,0.4}, xticklabel style={/pgf/number format/fixed},",
    "                    yticklabel style={/pgf/number format/fixed, /pgf/number format/precision=3}]",
    "  \\nextgroupplot[ylabel={$\\omega_{-}$ (\\si{\\radian\\per\\second})},",
    "                 title={\\footnotesize\\scshape (a)\\; frecuencia de modulación}, title style={text=grispizarra}]",
    "    \\addplot[puntoA, error bars/.cd, y dir=both, y explicit, error bar style={azulprof}]",
    "      table[x=x, y=wm, y error=uwm] {\\datoscoh};",
    "    \\addlegendentry{$\\omega_{-}=\\pi/T_{\\mathrm{b}}$ (pulsante)}",
    "    \\addplot[puntoB, error bars/.cd, y dir=both, y explicit, error bar style={azulacero}]",
    "      table[x=x, y=tm, y error=utm] {\\datoscoh};",
    "    \\addlegendentry{$(\\omega_1-\\omega_0)/2$ (fase y contrafase)}",
    "  \\nextgroupplot[ylabel={$\\omega_{+}$ (\\si{\\radian\\per\\second})},",
    "                 title={\\footnotesize\\scshape (b)\\; frecuencia portadora}, title style={text=grispizarra}]",
    "    \\addplot[puntoA, error bars/.cd, y dir=both, y explicit, error bar style={azulprof}]",
    "      table[x=x, y=wp, y error=uwp] {\\datoscoh};",
    "    \\addlegendentry{$\\omega_{+}=2\\pi/T_{+}$ (pulsante)}",
    "    \\addplot[puntoB, error bars/.cd, y dir=both, y explicit, error bar style={azulacero}]",
    "      table[x=x, y=tp, y error=utm] {\\datoscoh};",
    "    \\addlegendentry{$(\\omega_1+\\omega_0)/2$ (fase y contrafase)}",
    "  \\end{groupplot}",
    "\\end{tikzpicture}"]

# robustez: ajustes de control con menos longitudes
lines.append("%%FILE robustez")
lines += ["\\begin{table}[ht]", "\\centering",
          "\\caption{Constante del muelle obtenida con cada método al repetir el ajuste sin las longitudes "
          "mayores. Con tres puntos el ajuste tiene un solo grado de libertad, así que estos valores solo "
          "indican tendencias.}",
          "\\label{tab:robustez}", "\\renewcommand{\\arraystretch}{1.25}",
          "\\begin{tabular}{@{}l c c c@{}}", "\\toprule",
          "Longitudes ajustadas & $k$ de $\\omega_1^{2}$ (\\si{\\newton\\per\\meter}) & "
          "$k$ de $\\omega_{-}$ (\\si{\\newton\\per\\meter}) & $k$ de $\\omega_{+}$ (\\si{\\newton\\per\\meter}) \\\\",
          "\\midrule"]
for m, etiqueta in ((5, "$\\ell_1$ a $\\ell_5$"), (4, "$\\ell_1$ a $\\ell_4$"), (3, "$\\ell_1$ a $\\ell_3$")):
    ks = []
    for clave, y in (("w1", W1C), ("wm", WM), ("wp", WP)):
        a = ajuste(X[:m], y[:m])[0]
        ks.append(I * a / 2 if clave == "w1" else 2 * M * G * LP * a / w0)
    lines.append(f"{etiqueta} & " + " & ".join(f"\\num{{{k:.2f}}}" for k in ks) + " \\\\")
lines += ["\\bottomrule", "\\end{tabular}", "\\end{table}", ""]

lines.append("%%FILE valores")
lines += [f"\\newcommand{{\\valTcero}}{{\\num{{{Tm:.4f}}}}}", f"\\newcommand{{\\valWcero}}{{\\num{{{w0:.4f}}}}}",
          f"\\newcommand{{\\valWceroDos}}{{\\num{{{w0 ** 2:.4f}}}}}", f"\\newcommand{{\\valI}}{{\\num{{{I:.4f}}}}}"]
for clave, nombre in (("w1", "Uno"), ("wm", "Menos"), ("wp", "Mas")):
    f = FITS[clave]
    lines += [f"\\newcommand{{\\valK{nombre}}}{{{pm(f['k'], f['Uk'])}}}"]
lines += [f"\\newcommand{{\\valIU}}{{{pm(I, T95_3 * uI)}}}", f"\\newcommand{{\\valWceroU}}{{{pm(w0, T95_3 * uw0)}}}"]
print("u(T0m)", uTm, "u(w0)", uw0, "u(I)", uI)
for c, f in FITS.items():
    print(c, {k: (round(v, 5) if isinstance(v, float) else None) for k, v in f.items() if k != "res"}, [round(r, 5) for r in f["res"]])
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
