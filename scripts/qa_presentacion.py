"""
Control de calidad de la presentación sin depender de LibreOffice.

Comprueba, diapositiva a diapositiva:
  1. Que ninguna forma se salga de los límites de la diapositiva.
  2. Que ningún texto desborde su caja (estimación tipográfica).
  3. Que no haya solapamientos entre cajas de texto.
  4. Que se respeten los márgenes mínimos.
"""
import math
import sys
from pptx import Presentation
from pptx.util import Emu

RUTA = sys.argv[1] if len(sys.argv) > 1 else "/home/user/DANISANTELMO/Alcopalet_Defensa.pptx"
MARGEN_MIN = 0.5          # pulgadas
EMU = 914400.0

# Ancho medio de carácter como fracción del cuerpo, por familia
ANCHO_CAR = {"Calibri": 0.479, "Cambria": 0.503}
INTERLINEA = 1.22


def pulg(v):
    return None if v is None else v / EMU


def alto_necesario(texto, ancho_pulg, cuerpo_pt, fuente, negrita):
    """Estima el alto que ocupará el texto, en pulgadas."""
    if not texto.strip():
        return 0.0
    factor = ANCHO_CAR.get(fuente, 0.49) * (1.045 if negrita else 1.0)
    ancho_pt = max(ancho_pulg * 72 - 4, 10)
    por_linea = max(int(ancho_pt / (factor * cuerpo_pt)), 1)
    lineas = 0
    for parrafo in texto.split("\n"):
        lineas += max(1, math.ceil(len(parrafo) / por_linea))
    return lineas * cuerpo_pt * INTERLINEA / 72


def rect(sh):
    try:
        return (pulg(sh.left), pulg(sh.top), pulg(sh.width), pulg(sh.height))
    except TypeError:
        return None


prs = Presentation(RUTA)
W = pulg(prs.slide_width)
H = pulg(prs.slide_height)
print(f"Diapositiva: {W:.2f} x {H:.2f} pulgadas · {len(prs.slides.__iter__.__self__._sldIdLst)} slides\n")

problemas = []

for n, slide in enumerate(prs.slides, 1):
    cajas = []
    for sh in slide.shapes:
        r = rect(sh)
        if r is None:
            continue
        x, y, w, h = r

        # 1. límites de la diapositiva
        if x < -0.02 or y < -0.02 or x + w > W + 0.02 or y + h > H + 0.02:
            problemas.append(f"S{n}  FUERA DE LÍMITES  {sh.shape_type} en "
                             f"({x:.2f},{y:.2f}) {w:.2f}x{h:.2f}")

        if not sh.has_text_frame:
            continue
        texto = sh.text_frame.text
        if not texto.strip():
            continue

        # tamaño y fuente del primer run con formato
        cuerpo, fuente, negrita = 14, "Calibri", False
        for p in sh.text_frame.paragraphs:
            for run in p.runs:
                if run.font.size:
                    cuerpo = run.font.size.pt
                if run.font.name:
                    fuente = run.font.name
                negrita = bool(run.font.bold)
                break
            break

        nec = alto_necesario(texto, w, cuerpo, fuente, negrita)
        if nec > h + 0.06:
            problemas.append(
                f"S{n}  DESBORDE  «{texto[:42].replace(chr(10),' ')}…» "
                f"necesita {nec:.2f}\" y tiene {h:.2f}\" ({cuerpo:.0f}pt)")

        # 4. margen respecto al borde
        if x < MARGEN_MIN - 0.02 or x + w > W - MARGEN_MIN + 0.02:
            problemas.append(f"S{n}  MARGEN  «{texto[:32]}…» x={x:.2f} ancho={w:.2f}")

        cajas.append((x, y, w, h, texto[:30]))

    # 3. solapamientos entre cajas de texto
    for i in range(len(cajas)):
        for j in range(i + 1, len(cajas)):
            ax, ay, aw, ah, at = cajas[i]
            bx, by, bw, bh, bt = cajas[j]
            sx = min(ax + aw, bx + bw) - max(ax, bx)
            sy = min(ay + ah, by + bh) - max(ay, by)
            if sx > 0.06 and sy > 0.06:
                problemas.append(
                    f"S{n}  SOLAPE  «{at.replace(chr(10),' ')}» x "
                    f"«{bt.replace(chr(10),' ')}» ({sx:.2f}\" x {sy:.2f}\")")

if problemas:
    print(f"{len(problemas)} incidencia(s):\n")
    for p in problemas:
        print("  " + p)
else:
    print("Sin incidencias: nada fuera de límites, sin desbordes ni solapes.")
