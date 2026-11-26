"""Reconstruye RESONANCIA.xlsx con formato formal y fórmulas para T, Ω y β.

Uso: python3 build.py RESONANCIA_original.xlsx RESONANCIA.xlsx (desde esta carpeta)."""
import io, sys, math
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.chart.data_source import NumDataSource, NumRef
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DFont
from openpyxl.drawing.line import LineProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L

from openpyxl.chart.title import Title
from openpyxl.chart.text import Text
from openpyxl.drawing.text import RegularTextRun
sys.path.insert(0, '.')
from defaults import propone

SRC, OUT = sys.argv[1], sys.argv[2]

# ---------------------------------------------------------------- paleta
MARINO, COBALTO, PETROLEO = '0B1F3A', '1F5A96', '12667A'
TURQUESA, AGUAMARINA, CIELO = '17A2A6', '86D3CF', 'B8DCEF'
PIZARRA, PERLA, FONDO = '44556B', 'E2EBF3', 'EFF6FB'
EDITABLE = 'FFF4CC'
FUENTE = 'Cambria'

def f(size=10, bold=False, italic=False, color=MARINO):
    return Font(name=FUENTE, size=size, bold=bold, italic=italic, color=color)
def fill(c):
    return PatternFill('solid', start_color=c, end_color=c)
fino = Side(style='thin', color=CIELO)
medio = Side(style='medium', color=MARINO)
CEN = Alignment(horizontal='center', vertical='center')
IZQ = Alignment(horizontal='left', vertical='center', indent=1)
DER = Alignment(horizontal='right', vertical='center')

def _cp(size, bold=False, color='0B1F3A'):
    return CharacterProperties(latin=DFont(typeface='Cambria'), sz=size, b=bold, solidFill=color)

def titulo(texto, size=1200, bold=True):
    cp = _cp(size, bold)
    rt = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=texto)])])
    return Title(tx=Text(rich=rt), overlay=False)

def txpr(size=800, color='44556B'):
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=_cp(size, color=color)),
                                 endParaRPr=_cp(size, color=color))])

def estiliza(ch, t, xt, yt, xfmt, yfmt):
    ch.title = titulo(t, 1150)
    ch.x_axis.title = titulo(xt, 900, False)
    ch.y_axis.title = titulo(yt, 900, False)
    ch.x_axis.txPr = txpr()
    ch.y_axis.txPr = txpr()
    ch.legend.txPr = txpr(850, '0B1F3A')
    ch.x_axis.number_format = xfmt
    ch.y_axis.number_format = yfmt
    ch.x_axis.majorGridlines = None
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill='E2EBF3'))
    for ax in (ch.x_axis, ch.y_axis):
        ax.spPr = GraphicalProperties(ln=LineProperties(solidFill='44556B', w=6350))
    ch.roundedCorners = False
    ch.legend.position = 'b'
    ch.legend.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False

def rich(*parts, size=10, bold=False, italic=False, color=MARINO):
    """parts: str o (str, 'sub'|'sup')"""
    out = []
    for p in parts:
        txt, va = (p, None) if isinstance(p, str) else p
        va = {'sub': 'subscript', 'sup': 'superscript'}.get(va, va)
        out.append(TextBlock(InlineFont(rFont=FUENTE, sz=size, b=bold, i=italic,
                                        color=color, vertAlign=va), txt))
    return CellRichText(out)

# ---------------------------------------------------------------- datos
src = load_workbook(SRC)
ORDEN = [('Resonancia valor -3', 'Ω −3', '−3'), ('Resonancia valor -2', 'Ω −2', '−2'),
         ('Resonancia valor -1', 'Ω −1', '−1'), ('Resonancia valor central', 'Ω R', 'R'),
         ('Resonancia valor +1', 'Ω +1', '+1'), ('Resonancia valor +2', 'Ω +2', '+2'),
         ('Resonancia valor +3', 'Ω +3', '+3')]
AMORT = [('SIN PASO', '0', 'sin puente'), ('PASO 4', '4', 'puente 4'), ('PASO 6', '6', 'puente 6')]

def lee(ws):
    blocks = [r for r in range(1, ws.max_row + 1)
              if isinstance(ws.cell(r, 2).value, str) and 'PASO' in ws.cell(r, 2).value]
    res = {}
    for i, b in enumerate(blocks):
        end = blocks[i + 1] - 1 if i + 1 < len(blocks) else ws.max_row
        rows = [[ws.cell(r, c).value for c in range(2, 7)]
                for r in range(b + 2, end + 1) if ws.cell(r, 3).value is not None]
        res[ws.cell(b, 2).value] = rows
    imgs = sorted(ws._images, key=lambda im: im.anchor._from.row)
    return res, [im._data() for im in imgs]

wb = Workbook()
res = wb.active
res.title = 'Resumen'
res.sheet_view.showGridLines = False

NMAX = 21            # filas de datos por bloque
BLOQUE = NMAX + 5    # altura de cada bloque en filas
FILA0 = 5            # primera fila del primer bloque

# celdas de parámetros en Resumen (se definen más abajo; se fijan aquí)
P_DELTA, P_UT, P_TOL = "Resumen!$D$6", "Resumen!$D$7", "Resumen!$D$8"

refs = {}   # (paso, amort) -> dict de celdas

for src_name, name, paso in ORDEN:
    datos, imgs = lee(src[src_name])
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    anchos = {'A': 2.5, 'B': 5, 'C': 9, 'D': 9, 'E': 9, 'F': 9, 'G': 10, 'H': 5, 'I': 5,
              'J': 6, 'K': 2.5, 'L': 15, 'M': 11, 'N': 12, 'O': 2.5}
    for k, v in anchos.items():
        ws.column_dimensions[k].width = v
    # título
    ws.merge_cells('B1:N1')
    ws['B1'] = rich('Oscilación forzada · ', ('Ω', None),
                    f' paso {paso}' if paso != 'R' else ' en resonancia', size=15, bold=True)
    ws['B1'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[1].height = 26
    ws.merge_cells('B2:N2')
    ws['B2'] = 'Análisis de picos · Δ = 0,1 s · misma Ω en los tres registros'
    ws['B2'].font = f(9.5, italic=True, color=PIZARRA)
    for c in range(2, 15):
        ws.cell(3, c).border = Border(top=Side(style='thin', color=TURQUESA))

    for k, (clave, sub, rotulo) in enumerate(AMORT):
        filas = datos[clave]
        h = np.array([r[3] for r in filas], float)
        tt = np.array([r[1] for r in filas], float)
        n0, excl, meseta = propone(tt, h)
        r0 = FILA0 + k * BLOQUE          # banda del bloque
        rh = r0 + 1                      # cabeceras
        rs, re_ = rh + 1, rh + NMAX      # datos
        # banda
        ws.merge_cells(start_row=r0, start_column=2, end_row=r0, end_column=10)
        c = ws.cell(r0, 2, rich(('γ', None), (sub, 'sub'), f'  ·  {rotulo}', size=11, bold=True,
                                color='FFFFFF'))
        for cc in range(2, 11):
            ws.cell(r0, cc).fill = fill(MARINO)
        c.alignment = IZQ
        ws.row_dimensions[r0].height = 20
        # cabeceras
        cab = [rich('n', italic=True), rich(('t', None), ('ent', 'sub'), ' (s)'),
               rich('t', ('n', 'sub'), ' (s)'), rich(('t', None), ('sal', 'sub'), ' (s)'),
               rich('φ', ('n', 'sub'), ' (°)'), rich('∫φ d', 't', ' (°·s)'),
               rich('✕'), rich('w', italic=True), rich('nota')]
        for j, v in enumerate(cab):
            c = ws.cell(rh, 2 + j, v)
            c.alignment = CEN
            c.fill = fill(PERLA)
            c.border = Border(bottom=Side(style='thin', color=MARINO))
        ws.cell(rh, 8).comment = Comment('1 = excluir pico', '')
        ws.cell(rh, 9).comment = Comment('1 si n ≥ n₀ y no excluido', '')
        # ficha (columnas L:N)
        ws.merge_cells(start_row=r0, start_column=12, end_row=r0, end_column=14)
        c = ws.cell(r0, 12, rich('Estado estacionario', size=10.5, bold=True, color='FFFFFF'))
        for cc in range(12, 15):
            ws.cell(r0, cc).fill = fill(PETROLEO)
        c.alignment = IZQ

        R = lambda col: f'${col}${rs}:${col}${re_}'
        n_, t_, phi, w = R('B'), R('D'), R('F'), R('I')
        F = {}  # fila de cada magnitud en la ficha
        ficha = [
            ('n0', rich('n', ('0', 'sub')), n0, '1.er máx. estac.'),
            ('N', rich('N', italic=True), f'=SUM({w})', ''),
            ('beta', rich('β = ', 'φ̄', ('n', 'sub')), None, '°'),
            ('s', rich('s', ('φ', 'sub')), None, '°'),
            ('ub', rich('u(β) = s/√N'), None, '°'),
            ('delta', rich('δ = Δφ/β'), None, 'deriva'),
            ('T', rich('T', italic=True), None, 's'),
            ('uT', rich('u(T)'), None, 's'),
            ('Om', rich('Ω = 2π/T'), None, 'rad/s'),
            ('uOm', rich('u(Ω)'), None, 'rad/s'),
            ('est', rich('estado'), None, ''),
            ('tbar', rich('t̄', ('n', 'sub'), size=8.5, color=PIZARRA), None, 'aux.'),
            ('nbar', rich('n̄', size=8.5, color=PIZARRA), None, 'aux.'),
            ('Sxx', rich('Σw(n−n̄)²', size=8.5, color=PIZARRA), None, 'aux.'),
            ('sres', rich('s', ('res', 'sub'), size=8.5, color=PIZARRA), None, 'aux.'),
        ]
        for i, (key, lab, _, _u) in enumerate(ficha):
            F[key] = r0 + 1 + i
        a = lambda key: f'$M${F[key]}'
        form = {
            'beta': f'=IF({a("N")}>0,SUMPRODUCT({w},{phi})/{a("N")},"")',
            's': f'=IF({a("N")}>1,SQRT(SUMPRODUCT({w},({phi}-{a("beta")})^2)/({a("N")}-1)),"")',
            'ub': f'=IF({a("N")}>1,{a("s")}/SQRT({a("N")}),"")',
            'tbar': f'=IF({a("N")}>0,SUMPRODUCT({w},{t_})/{a("N")},"")',
            'nbar': f'=IF({a("N")}>0,SUMPRODUCT({w},{n_})/{a("N")},"")',
            'Sxx': f'=SUMPRODUCT({w},({n_}-{a("nbar")})^2)',
            'T': f'=IF({a("N")}>2,SUMPRODUCT({w},{n_}-{a("nbar")},{t_}-{a("tbar")})/{a("Sxx")},"")',
            'sres': f'=IF({a("N")}>2,SQRT(SUMPRODUCT({w},({t_}-{a("tbar")}-{a("T")}*({n_}-{a("nbar")}))^2)/({a("N")}-2)),"")',
            'uT': f'=IF({a("N")}>2,SQRT(MAX({a("sres")}^2,{P_UT}^2)/{a("Sxx")}),"")',
            'Om': f'=IF({a("N")}>2,2*PI()/{a("T")},"")',
            'uOm': f'=IF({a("N")}>2,{a("Om")}*{a("uT")}/{a("T")},"")',
            'delta': f'=IF({a("N")}>2,SUMPRODUCT({w},{t_}-{a("tbar")},{phi}-{a("beta")})'
                     f'/SUMPRODUCT({w},({t_}-{a("tbar")})^2)*({a("N")}-1)*{a("T")}/{a("beta")},"")',
            'est': f'=IF({a("N")}<4,"⚠ N < 4",IF(ABS({a("delta")})<={P_TOL},"✓ estacionario",'
                   f'IF({a("delta")}>0,"↗ transitorio","↘ transitorio")))',
        }
        fmts = {'n0': '0', 'N': '0', 'beta': '0.0', 's': '0.0', 'ub': '0.0', 'delta': '+0%;−0%;0%',
                'T': '0.0000', 'uT': '0.0000', 'Om': '0.0000', 'uOm': '0.0000',
                'tbar': '0.00', 'nbar': '0.0', 'Sxx': '0.0', 'sres': '0.000'}
        for key, lab, val, unidad in ficha:
            r = F[key]
            ws.cell(r, 12, lab).alignment = IZQ
            v = val if key in ('n0', 'N') else form[key]
            c = ws.cell(r, 13, v)
            c.number_format = fmts.get(key, 'General')
            c.alignment = DER if key != 'est' else Alignment(horizontal='left', vertical='center')
            aux = key in ('tbar', 'nbar', 'Sxx', 'sres')
            c.font = f(8.5 if aux else 10, bold=key in ('beta', 'Om'), color=PIZARRA if aux else MARINO)
            u = ws.cell(r, 14, unidad)
            u.font = f(8.5, italic=True, color=PIZARRA)
            for cc in (12, 13, 14):
                ws.cell(r, cc).border = Border(bottom=fino)
            if key in ('beta', 'ub', 'Om', 'uOm'):
                for cc in (12, 13, 14):
                    ws.cell(r, cc).fill = fill(FONDO)
        # n0 editable
        c = ws.cell(F['n0'], 13)
        c.fill = fill(EDITABLE)
        c.font = f(10, bold=True, color=COBALTO)
        c.comment = Comment('|δ| ≤ 10 % desde n₀' if meseta else 'sin meseta: últimos 5', '')
        ws.cell(F['est'], 13).font = f(10, bold=True, color=PETROLEO)
        # datos
        for i in range(NMAX):
            r = rs + i
            if i < len(filas):
                te, tn, ts, al, ar = filas[i]
                n = i + 1
                vals = [n, te, tn, ts, al, ar]
                for j, v in enumerate(vals):
                    c = ws.cell(r, 2 + j, v)
                    c.number_format = ['0', '0.0', '0.0', '0.0', '0.000', '0.000'][j]
                    c.alignment = CEN
                    c.font = f(10, italic=(j == 0), color=PIZARRA if j in (0, 1, 3, 5) else MARINO)
                x = ws.cell(r, 8, 1 if n in excl else None)
                x.alignment = CEN
                x.fill = fill(EDITABLE)
                x.font = f(10, bold=True, color=COBALTO)
                c = ws.cell(r, 9, f'=IF(AND(B{r}>=$M${F["n0"]},H{r}<>1),1,0)')
                c.alignment = CEN
                c.font = f(10, color=TURQUESA, bold=True)
                nota = ''
                if n in excl:
                    nota = '✱'
                elif i == 0 and te <= 0.15:
                    nota = '◂'
                c = ws.cell(r, 10, nota)
                c.alignment = CEN
                c.font = f(10, color=PETROLEO)
            for cc in range(2, 11):
                ws.cell(r, cc).border = Border(bottom=Side(style='hair', color=CIELO))
        for cc in range(2, 11):
            ws.cell(re_, cc).border = Border(bottom=Side(style='thin', color=MARINO))
        # filas usadas: fondo suave
        ws.conditional_formatting.add(
            f'B{rs}:G{re_}', FormulaRule(formula=[f'$I{rs}=1'], fill=fill(FONDO),
                                         font=Font(name=FUENTE, bold=False)))
        # captura de pantalla
        if k < len(imgs):
            img = XLImage(io.BytesIO(imgs[k]))
            ratio = img.height / img.width
            img.width = 560
            img.height = round(560 * ratio)
            ws.add_image(img, f'P{r0}')
        refs[(paso, sub)] = {kk: f"'{name}'!$M${F[kk]}" for kk in ('N', 'beta', 'ub', 'Om', 'uOm', 'est', 'delta')}
        refs[(paso, sub)]['rango'] = (name, rs, re_)

    # leyenda corta al pie
    rl = FILA0 + 3 * BLOQUE
    ws.merge_cells(start_row=rl, start_column=2, end_row=rl, end_column=14)
    ws.cell(rl, 2, '◂ primer pico truncado · ✱ excluido · w = 1: usado').font = f(8.5, italic=True, color=PIZARRA)

    # gráfica φn(tn) de los tres registros
    ch = ScatterChart()
    ch.style = 13
    ch.height, ch.width = 8, 17
    colores = [MARINO, COBALTO, TURQUESA]
    marcas = ['circle', 'square', 'triangle']
    for k, (clave, sub, rotulo) in enumerate(AMORT):
        nm, rs, re_ = refs[(paso, sub)]['rango']
        xs = Reference(ws, min_col=4, min_row=rs, max_row=re_)
        ys = Reference(ws, min_col=6, min_row=rs, max_row=re_)
        s = Series(ys, xs, title=f'γ{"₀₄₆"[k]} · {rotulo}')
        s.marker = Marker(symbol=marcas[k], size=6)
        s.marker.graphicalProperties = GraphicalProperties(solidFill=colores[k])
        s.marker.graphicalProperties.line.solidFill = colores[k]
        s.graphicalProperties.line.solidFill = colores[k]
        s.graphicalProperties.line.width = 15000
        s.smooth = False
        ch.series.append(s)
    estiliza(ch, 'φₙ frente a tₙ · evolución hacia el estado estacionario', 'tₙ (s)', 'φₙ (°)', '0', '0')
    ch.x_axis.scaling.min = 0
    ch.y_axis.scaling.min = 0
    ws.add_chart(ch, f'B{rl + 2}')
    ws.freeze_panes = 'A4'

# ---------------------------------------------------------------- Resumen
ws = res
for k, v in {'A': 2.5, 'B': 9, 'C': 22, 'D': 10, 'E': 9, 'F': 2}.items():
    ws.column_dimensions[k].width = v
ws.merge_cells('B1:V1')
ws['B1'] = rich('Péndulo de Pohl · curvas de resonancia ', ('β', None), '(', 'Ω', ')', size=16, bold=True)
ws.row_dimensions[1].height = 28
ws.merge_cells('B2:V2')
ws['B2'] = 'Práctica 4 · 7 frecuencias × 3 amortiguamientos · β y Ω con los máximos estacionarios'
ws['B2'].font = f(9.5, italic=True, color=PIZARRA)
for c in range(2, 23):
    ws.cell(3, c).border = Border(top=Side(style='thin', color=TURQUESA))

# parámetros
ws.merge_cells('B5:E5')
ws['B5'] = rich('Parámetros', bold=True, color='FFFFFF')
for c in range(2, 6):
    ws.cell(5, c).fill = fill(PETROLEO)
par = [
    (6, rich('Δ'), 'intervalo de muestreo', 0.1, 's', '0.0'),
    (7, rich('u(t', ('n', 'sub'), ')'), 'Δ/(2√3)', '=D6/(2*SQRT(3))', 's', '0.000'),
    (8, rich('|δ|', ('máx', 'sub')), 'umbral estacionario', 0.10, '', '0%'),
    (9, rich('ε', ('β', 'sub')), 'sesgo de φₙ ≤ (ΩΔ)²/8', '=(3.37*D6)^2/8', '', '0.0%'),
]
for r, s, d, v, u, fm in par:
    ws.cell(r, 2, s).alignment = IZQ
    ws.cell(r, 3, d).font = f(9, italic=True, color=PIZARRA)
    c = ws.cell(r, 4, v)
    c.number_format = fm
    c.alignment = DER
    c.font = f(10, bold=True, color=COBALTO if not str(v).startswith('=') else MARINO)
    if not str(v).startswith('='):
        c.fill = fill(EDITABLE)
    ws.cell(r, 5, u).font = f(9, italic=True, color=PIZARRA)
    for cc in range(2, 6):
        ws.cell(r, cc).border = Border(bottom=fino)
ws['D9'].comment = Comment('con Ω ≈ 3,37 rad/s', '')

# oscilación libre: entradas para Ω_R
ws.merge_cells('B11:C11')
ws['B11'] = rich('Oscilación libre → ', 'Ω', ('R', 'sub'), ' = √(ω', ('0', 'sub'), '² − 2γ²)',
                 bold=True, color='FFFFFF')
for c in range(2, 7):
    ws.cell(11, c).fill = fill(PETROLEO)
for c, t in ((4, 'valor'), (5, 'u')):
    ws.cell(11, c, t).font = f(9, italic=True, color='FFFFFF')
    ws.cell(11, c).alignment = CEN
hdr = [rich(''), rich('valor'), rich('u'), rich('')]
libre = [(12, rich('ω', ('0', 'sub')), 'rad/s', 'frecuencia natural'),
         (13, rich('γ', ('0', 'sub')), 's⁻¹', 'sin puente'),
         (14, rich('γ', ('4', 'sub')), 's⁻¹', 'puente 4'),
         (15, rich('γ', ('6', 'sub')), 's⁻¹', 'puente 6')]
ws.cell(11, 6).value = None
for r, s, u, d in libre:
    ws.cell(r, 2, s).alignment = IZQ
    ws.cell(r, 3, d).font = f(9, italic=True, color=PIZARRA)
    for cc in (4, 5):
        c = ws.cell(r, cc)
        c.fill = fill(EDITABLE)
        c.number_format = '0.0000'
        c.font = f(10, bold=True, color=COBALTO)
        c.alignment = DER
    ws.cell(r, 6, u).font = f(9, italic=True, color=PIZARRA)
    for cc in range(2, 7):
        ws.cell(r, cc).border = Border(bottom=fino)
ws.column_dimensions['F'].width = 7
ws['D12'].comment = Comment('de la oscilación libre (cuestiones 1 y 2)', '')
for i, (r, g) in enumerate([(17, 13), (18, 14), (19, 15)]):
    ws.cell(r, 2, rich('Ω', ('R', 'sub'), ',', ('γ', None), ("046"[i], 'sub'))).alignment = IZQ
    ws.cell(r, 3, 'u: ec. (3.7)').font = f(8.5, italic=True, color=PIZARRA)
    c = ws.cell(r, 4, f'=IF(OR($D$12="",D{g}=""),"",SQRT($D$12^2-2*D{g}^2))')
    c.number_format = '0.0000'
    c.font = f(10, bold=True)
    c.alignment = DER
    c = ws.cell(r, 5, f'=IF(D{r}="","",SQRT(($D$12*$E$12)^2+(2*D{g}*E{g})^2)/D{r})')
    c.number_format = '0.0000'
    c.alignment = DER
    ws.cell(r, 6, 'rad/s').font = f(9, italic=True, color=PIZARRA)
    for cc in range(2, 7):
        ws.cell(r, cc).border = Border(bottom=fino)
        ws.cell(r, cc).fill = fill(FONDO)

# tabla principal (columnas H:V)
T0 = 5
ws.column_dimensions['G'].width = 3
ws.column_dimensions['H'].width = 7
grupos = [('0', 'sin puente', 9), ('4', 'puente 4', 14), ('6', 'puente 6', 19)]
ws.merge_cells(start_row=T0, start_column=8, end_row=T0 + 1, end_column=8)
c = ws.cell(T0, 8, rich('paso', bold=True, color='FFFFFF'))
c.alignment = CEN
for rr in (T0, T0 + 1):
    ws.cell(rr, 8).fill = fill(MARINO)
sub_cab = [rich('Ω (rad/s)'), rich('u(Ω)'), rich('β (°)'), rich('u(β)'), rich('estado')]
for sub, rot, c0 in grupos:
    ws.merge_cells(start_row=T0, start_column=c0, end_row=T0, end_column=c0 + 4)
    c = ws.cell(T0, c0, rich('γ', (sub, 'sub'), f' · {rot}', bold=True, color='FFFFFF'))
    c.alignment = CEN
    for cc in range(c0, c0 + 5):
        ws.cell(T0, cc).fill = fill(MARINO if sub == '0' else (COBALTO if sub == '4' else PETROLEO))
        c = ws.cell(T0 + 1, cc, sub_cab[cc - c0])
        c.alignment = CEN
        c.fill = fill(PERLA)
        c.border = Border(bottom=Side(style='thin', color=MARINO))
    for cc, wdt in zip(range(c0, c0 + 5), (9.5, 8, 7.5, 6.5, 14)):
        ws.column_dimensions[L(cc)].width = wdt
ws.cell(T0 + 1, 8).border = Border(bottom=Side(style='thin', color=MARINO))
pasos = [p for _, _, p in ORDEN]
for i, p in enumerate(pasos):
    r = T0 + 2 + i
    c = ws.cell(r, 8, 'R' if p == 'R' else p)
    c.alignment = CEN
    c.font = f(10, bold=True, color=MARINO)
    for sub, rot, c0 in grupos:
        rf = refs[(p, sub)]
        for j, key in enumerate(('Om', 'uOm', 'beta', 'ub', 'est')):
            c = ws.cell(r, c0 + j, f'={rf[key]}')
            c.number_format = ['0.0000', '0.0000', '0.0', '0.0', 'General'][j]
            c.alignment = CEN if key == 'est' else DER
            c.font = f(10, bold=key in ('Om', 'beta'), color=MARINO if key != 'est' else PETROLEO)
            if key in ('uOm', 'ub'):
                c.font = f(9, color=PIZARRA)
    for cc in range(8, 24):
        ws.cell(r, cc).border = Border(bottom=Side(style='hair', color=CIELO))
    if p == 'R':
        for cc in range(8, 24):
            ws.cell(r, cc).fill = fill(FONDO)
rlast = T0 + 2 + len(pasos) - 1
for cc in range(8, 24):
    ws.cell(rlast, cc).border = Border(bottom=Side(style='thin', color=MARINO))
# separación visual entre grupos
for c0 in (14, 19):
    for r in range(T0, rlast + 1):
        b = ws.cell(r, c0).border
        ws.cell(r, c0).border = Border(left=Side(style='thin', color=CIELO), bottom=b.bottom)

# anotaciones (muy sintéticas)
rn = rlast + 2
ws.merge_cells(start_row=rn, start_column=8, end_row=rn, end_column=23)
ws.cell(rn, 8, rich('Anotaciones', bold=True, color='FFFFFF')).alignment = IZQ
for cc in range(8, 24):
    ws.cell(rn, cc).fill = fill(PETROLEO)
notas = [
    '⚠  Ω +2: γ₆ ≡ γ₄ (registro duplicado) → repetir.',
    '↘  γ₄, γ₆: transitorio al poner el puente → β solo de la meseta (n ≥ n₀).',
    '↘  γ₆ en −2, −1 y R: sin meseta → β por exceso.',
    '↗  γ₀: φₙ crece al arrancar · ✱ excluidos (171°, 146°).',
    'Ω(γ₄, γ₆) − Ω(γ₀) ≈ +0,04 rad/s a motor fijo → β frente a Ω medida, no frente al paso.',
    'Barrido γ₀: ΔΩ ≈ 0,02 rad/s ≪ 2γ.',
]
for i, t in enumerate(notas):
    r = rn + 1 + i
    ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=23)
    c = ws.cell(r, 8, t)
    c.font = f(9.5, color=MARINO)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    for cc in range(8, 24):
        ws.cell(r, cc).border = Border(bottom=Side(style='hair', color=CIELO))

# leyenda
rl = rn + len(notas) + 2
ws.merge_cells(start_row=rl, start_column=8, end_row=rl, end_column=23)
ws.cell(rl, 8, '✓ |δ| ≤ 10 % · ↗ ↘ deriva de φₙ · ⚠ revisar · ✱ excluido · δ = (dφ/dt)(N − 1)T/β').font = f(8.5, italic=True, color=PIZARRA)

# datos para Ω_R (líneas verticales en la gráfica), apartados en Y:Z
for i, r in enumerate((17, 18, 19)):
    base = 6 + 2 * i
    for j in range(2):
        ws.cell(base + j, 25, f'=IF($D${r}="","",$D${r})').number_format = '0.0000'
        ws.cell(base + j, 26, 0 if j == 0 else 140)
        ws.cell(base + j, 25).font = f(8, color='B0B8C4')
        ws.cell(base + j, 26).font = f(8, color='B0B8C4')
ws.cell(5, 25, 'gráfica Ω_R').font = f(8, italic=True, color='B0B8C4')
ws.column_dimensions['X'].width = 4
# gráfica de resonancia β(Ω)
ch = ScatterChart()
ch.style = 13
ch.height, ch.width = 10.5, 21
colores = [MARINO, COBALTO, TURQUESA]
marcas = ['circle', 'square', 'triangle']
r1, r2 = T0 + 2, rlast
for k, (sub, rot, c0) in enumerate(grupos):
    xs = Reference(ws, min_col=c0, min_row=r1, max_row=r2)
    ys = Reference(ws, min_col=c0 + 2, min_row=r1, max_row=r2)
    s = Series(ys, xs, title=f'γ{"₀₄₆"[k]} · {rot}')
    s.marker = Marker(symbol=marcas[k], size=7)
    s.marker.graphicalProperties = GraphicalProperties(solidFill=colores[k])
    s.marker.graphicalProperties.line.solidFill = 'FFFFFF'
    s.graphicalProperties.line.noFill = True
    eref = f"Resumen!${L(c0 + 3)}${r1}:${L(c0 + 3)}${r2}"
    s.errBars = ErrorBars(errDir='y', errBarType='both', errValType='cust', noEndCap=False,
                          plus=NumDataSource(numRef=NumRef(f=eref)),
                          minus=NumDataSource(numRef=NumRef(f=eref)))
    s.errBars.spPr = GraphicalProperties(ln=LineProperties(solidFill=colores[k], w=9525))
    ch.series.append(s)
for i in range(3):
    base = 6 + 2 * i
    xs = Reference(ws, min_col=25, min_row=base, max_row=base + 1)
    ys = Reference(ws, min_col=26, min_row=base, max_row=base + 1)
    s = Series(ys, xs, title=f'Ω_R(γ{"₀₄₆"[i]}) teórica')
    s.marker = Marker(symbol='none')
    s.graphicalProperties.line.solidFill = colores[i]
    s.graphicalProperties.line.dashStyle = 'dash'
    s.graphicalProperties.line.width = 12700
    ch.series.append(s)
estiliza(ch, 'Curvas de resonancia · amplitud estacionaria β frente a Ω', 'Ω (rad/s)', 'β (°)', '0.00', '0')
ch.y_axis.scaling.min = 0
ch.y_axis.scaling.max = 140
ch.x_axis.scaling.min = 3.32
ch.x_axis.scaling.max = 3.42
ws.add_chart(ch, 'H' + str(rl + 2))

# fuente por defecto en todas las celdas sin estilo explícito
for sh in wb.worksheets:
    for row in sh.iter_rows():
        for c in row:
            if c.value is not None and c.font.name != FUENTE:
                c.font = f(10)
    sh.sheet_properties.tabColor = MARINO if sh.title == 'Resumen' else (
        TURQUESA if sh.title == 'Ω R' else COBALTO)
for sh in wb.worksheets:
    sh.page_setup.orientation = 'landscape'
    sh.page_setup.paperSize = sh.PAPERSIZE_A4
    sh.sheet_properties.pageSetUpPr.fitToPage = True
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0 if sh.title != 'Resumen' else 1
    if sh.title == 'Resumen':
        sh.print_area = 'A1:W50'
    sh.print_options.horizontalCentered = True
    sh.page_margins.left = sh.page_margins.right = 0.4
    sh.page_margins.top = sh.page_margins.bottom = 0.5
# símbolos en modo texto (evita que se muestren como emoji)
def vs(t):
    for sym in ('⚠', '↘', '↗'):
        t = t.replace(sym + '\ufe0e', sym).replace(sym, sym + '\ufe0e')
    return t
for sh in wb.worksheets:
    for row in sh.iter_rows():
        for c in row:
            if isinstance(c.value, str):
                c.value = vs(c.value)
            elif isinstance(c.value, CellRichText):
                for b in c.value:
                    if isinstance(b, TextBlock):
                        b.text = vs(b.text)
# metadatos del archivo original
for k in ('creator', 'lastModifiedBy', 'created'):
    setattr(wb.properties, k, getattr(src.properties, k))
wb.active = 0
wb.save(OUT)
print('ok')
