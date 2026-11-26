"""Recorta, endereza y mejora las fotos del laboratorio para el informe.

Uso, desde la raíz del repositorio:  python3 informe-pohl/scripts/procesar_fotos.py
Las coordenadas de las esquinas están en píxeles de la foto original
(ya girada según su orientación EXIF).
"""
import pathlib
from PIL import Image, ImageOps, ImageEnhance, ImageFilter

RAIZ = pathlib.Path(__file__).resolve().parents[2]
OUT = RAIZ / "informe-pohl" / "figuras"

def foto(nombre):
    ruta = RAIZ / f"WhatsApp Image 2026-11-26 at {nombre}.jpeg"
    return ImageOps.exif_transpose(Image.open(ruta)).convert("RGB")

def mejora(im, contraste=1.06, nitidez=1.35, corte=0.4, balance=True):
    if balance:  # balance de blancos: normaliza cada canal por separado
        im = Image.merge("RGB", [ImageOps.autocontrast(c, cutoff=corte) for c in im.split()])
    im = ImageEnhance.Contrast(im).enhance(contraste)
    im = im.filter(ImageFilter.UnsharpMask(radius=2, percent=80, threshold=2))
    return ImageEnhance.Sharpness(im).enhance(nitidez)

def endereza(im, esquinas, tam):
    """esquinas: sup. izq., inf. izq., inf. dcha., sup. dcha. (orden de PIL)."""
    plano = [c for p in esquinas for c in p]
    return im.transform(tam, Image.QUAD, plano, resample=Image.BICUBIC)

def guarda(im, nombre, ancho=None, q=92):
    if ancho and im.width < ancho:  # reescalado con Lanczos para ganar definición al imprimir
        im = im.resize((ancho, round(im.height * ancho / im.width)), Image.LANCZOS)
        im = im.filter(ImageFilter.UnsharpMask(radius=1.2, percent=50, threshold=1))
    im.save(OUT / nombre, quality=q, optimize=True, subsampling=0)
    print(nombre, im.size)

# --- montaje: péndulo de Pohl con motor, rectificador y fuente ----------------
im = foto("16.07.55 (5)")
guarda(mejora(im.crop((0, 180, 1500, 1800)), contraste=1.04, nitidez=1.3), "montaje-pohl.jpg", ancho=2100)

# --- material auxiliar --------------------------------------------------------
guarda(mejora(foto("16.07.54 (1)").crop((0, 80, 1500, 1500))), "cobra3.jpg", ancho=1600)
guarda(mejora(foto("16.07.55").crop((0, 380, 1500, 1880))), "puente.jpg", ancho=1600)
guarda(mejora(foto("16.07.54").crop((60, 450, 1500, 1950))), "fuente-rectificador.jpg", ancho=1600)

# --- capturas de PHYWE Measure (corrección de perspectiva) --------------------
capturas = {
    "16.07.55 (1)": ("registro-libre.jpg",
                     [(205, 255), (262, 1200), (1790, 1200), (1845, 195)], (2400, 1500)),
    "16.07.54 (6)": ("picos-debil.jpg",
                     [(100, 55), (110, 1395), (1925, 1385), (1935, 265)], (2400, 1650)),
    "16.07.55 (3)": ("picos-medio.jpg",
                     [(172, 265), (190, 1385), (1730, 1378), (1712, 190)], (2400, 1700)),
    "16.07.55 (4)": ("picos-fuerte.jpg",
                     [(228, 290), (238, 1370), (1655, 1368), (1675, 212)], (2400, 1750)),
}
for nombre, (salida, esquinas, tam) in capturas.items():
    guarda(mejora(endereza(foto(nombre), esquinas, tam), contraste=1.08, nitidez=1.5, corte=0.6), salida)
