"""Recorta, endereza y mejora las fotos del laboratorio para el informe.

Uso, desde la raíz del repositorio:  python3 informe/scripts/procesar_fotos.py
"""
import glob, pathlib
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
RAIZ = pathlib.Path(__file__).resolve().parents[2]
SRC = str(RAIZ) + "/"                      # fotos originales (WhatsApp Image ...jpeg)
OUT = str(RAIZ / "informe" / "figuras") + "/"
fs = sorted(glob.glob(SRC + "WhatsApp*.jpeg"))

def load(i):
    return ImageOps.exif_transpose(Image.open(fs[i])).convert("RGB")

def enhance(im, contrast=1.08, sharp=1.6, cutoff=0.6, wb=True):
    if wb:  # balance de blancos: normaliza cada canal por separado
        r, g, b = im.split()
        r, g, b = [ImageOps.autocontrast(c, cutoff=cutoff) for c in (r, g, b)]
        im = Image.merge("RGB", (r, g, b))
    im = ImageEnhance.Contrast(im).enhance(contrast)
    im = im.filter(ImageFilter.UnsharpMask(radius=2, percent=90, threshold=2))
    return ImageEnhance.Sharpness(im).enhance(sharp)

def quad(im, pts, size):
    """pts en coordenadas de la previsualización de 900 px de lado mayor."""
    s = max(im.size) / 900
    flat = []
    for (x, y) in pts:          # TL, BL, BR, TR (orden de PIL)
        flat += [x * s, y * s]
    return im.transform(size, Image.QUAD, flat, resample=Image.BICUBIC)

def save(im, name, width=None, q=92):
    if width and im.width < width:
        h = round(im.height * width / im.width)
        im = im.resize((width, h), Image.LANCZOS)
    im.save(OUT + name, quality=q, optimize=True)
    print(name, im.size)

# --- montaje general y detalle del muelle (foto 3) ---
im = load(3)
s = im.width / 675
crop = im.crop((int(150*s), int(0*s), int(640*s), int(890*s)))
save(enhance(crop, contrast=1.05, sharp=1.3, cutoff=0.3), "montaje-foto.jpg")
det = im.crop((int(230*s), int(380*s), int(470*s), int(520*s)))
save(enhance(det, contrast=1.05, sharp=1.5, cutoff=0.3), "montaje-muelle.jpg", width=1400)

# --- capturas de PHYWE Measure ---
jobs = {
    1:  ("senal-sin-suavizar.jpg", [(152, 4), (184, 462), (836, 432), (834, 2)], (1500, 900)),
    5:  ("senal-suavizada.jpg",    [(62, 42), (66, 640), (872, 628), (880, 60)], (1500, 1060)),
    4:  ("analisis-picos.jpg",     [(182, 150), (182, 446), (790, 466), (822, 204)], (1500, 720)),
    9:  ("picos-contrafase.jpg",   [(52, 20), (60, 606), (850, 612), (870, 26)], (1500, 1080)),
    7:  ("pulsante-L1.jpg",        [(104, 22), (120, 574), (880, 552), (884, 18)], (1500, 1060)),
    13: ("pulsante-varios.jpg",    [(110, 58), (126, 576), (858, 554), (856, 52)], (1500, 1040)),
    12: ("pulsante-dos-canales.jpg", [(72, 64), (100, 620), (880, 610), (876, 50)], (1500, 1060)),
    8:  ("picos-envolvente.jpg",   [(10, 0), (24, 640), (880, 632), (884, 0)], (1500, 1080)),
}
for i, (name, pts, size) in jobs.items():
    save(enhance(quad(load(i), pts, size)), name)

# menú Análisis → Suavizar (foto vertical)
im = load(2)
s = im.width / 675
save(enhance(im.crop((int(0*s), int(60*s), int(500*s), int(640*s)))), "menu-suavizar.jpg")

# anotación manuscrita (girada 90º para leerla en horizontal)
im = load(10)
s = im.width / 506
paper = im.crop((int(0*s), int(80*s), int(506*s), int(760*s))).rotate(90, expand=True)
g = ImageOps.grayscale(paper)
g = ImageOps.autocontrast(g, cutoff=1.5)
g = ImageEnhance.Contrast(g).enhance(1.35)
save(g.convert("RGB").filter(ImageFilter.UnsharpMask(radius=2, percent=80)), "nota-manuscrita.jpg")
