"""
Inyecta en un .xlsx el valor calculado de cada fórmula, conservando la fórmula.

openpyxl escribe las fórmulas sin valor cacheado. Excel de escritorio recalcula al
abrir, pero los visores (móvil, vista previa, Excel Online, Google Sheets) muestran
la celda vacía. Este script evalúa el libro con la librería `formulas` y escribe el
resultado dentro del XML como <v>, que es lo que esos visores leen.

Uso:  python3 scripts/inyectar_valores.py <archivo.xlsx>
"""
import re
import shutil
import sys
import warnings
import zipfile
from pathlib import Path

warnings.filterwarnings("ignore")
import formulas

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def evaluar(ruta):
    """Devuelve {(HOJA_MAYUS, 'B8'): valor} para todas las celdas calculadas."""
    xl = formulas.ExcelModel().loads(str(ruta)).finish()
    sol = xl.calculate()
    out, errores = {}, []
    for clave, celda in sol.items():
        if "]" not in clave:
            continue
        ref = clave.split("]", 1)[1].replace("'", "")
        if "!" not in ref:
            continue
        hoja, coord = ref.rsplit("!", 1)
        if not re.fullmatch(r"[A-Z]+[0-9]+", coord):
            continue
        try:
            valor = celda.value[0, 0]
        except Exception:
            continue
        out[(hoja.upper(), coord)] = valor
        if isinstance(valor, str) and valor.startswith("#"):
            errores.append((hoja, coord, valor))
    return out, errores


def mapa_hojas(z):
    """Relaciona el nombre de cada hoja con su fichero xl/worksheets/sheetN.xml.

    Los atributos de <Relationship> no vienen en orden fijo, y el Target puede
    llevar barra inicial, así que se extrae cada atributo por separado.
    """
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")

    id_destino = {}
    for etiqueta in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'\bId="([^"]+)"', etiqueta)
        destino = re.search(r'\bTarget="([^"]+)"', etiqueta)
        if rid and destino:
            id_destino[rid.group(1)] = destino.group(1)

    salida = {}
    for etiqueta in re.findall(r"<sheet\b[^>]*/?>", wb):
        nombre = re.search(r'\bname="([^"]+)"', etiqueta)
        rid = re.search(r'\br:id="([^"]+)"', etiqueta)
        if not (nombre and rid):
            continue
        destino = id_destino.get(rid.group(1), "").lstrip("/")
        if not destino:
            continue
        if not destino.startswith("xl/"):
            destino = "xl/" + destino
        salida[nombre.group(1).upper()] = destino
    return salida


def escapar(texto):
    return (texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def inyectar_hoja(xml, valores_hoja):
    """Añade <v> a cada <c> que tenga <f> y del que conozcamos el resultado."""
    inyectadas = 0

    def sustituir(m):
        nonlocal inyectadas
        celda = m.group(0)
        coord = re.search(r'\br="([A-Z]+[0-9]+)"', celda)
        if not coord or coord.group(1) not in valores_hoja:
            return celda
        if "<f" not in celda:
            return celda
        valor = valores_hoja[coord.group(1)]
        if valor is None:
            return celda
        # openpyxl deja un <v /> vacío tras la fórmula: hay que retirarlo antes
        # de escribir el valor real, o el XML queda con dos elementos <v>.
        celda = re.sub(r"<v\b[^>]*/>", "", celda)
        celda = re.sub(r"<v\b[^>]*>.*?</v>", "", celda, flags=re.DOTALL)
        if isinstance(valor, bool):
            nuevo, tipo = ("1" if valor else "0"), "b"
        elif isinstance(valor, (int, float)):
            if valor != valor or valor in (float("inf"), float("-inf")):
                return celda
            nuevo, tipo = repr(float(valor)), None
        else:
            texto = str(valor)
            if texto == "" or texto.startswith("#"):
                return celda
            nuevo, tipo = escapar(texto), "str"
        celda = re.sub(r"(</f>)", r"\1<v>" + nuevo + "</v>", celda, count=1)
        if tipo:
            if re.search(r'\bt="[^"]*"', celda):
                celda = re.sub(r'\bt="[^"]*"', f't="{tipo}"', celda, count=1)
            else:
                celda = celda.replace("<c ", f'<c t="{tipo}" ', 1)
        inyectadas += 1
        return celda

    # Las celdas vacías con estilo se escriben autocerradas (<c r="C8" s="46"/>).
    # Hay que reconocerlas ANTES que la forma con contenido, o el patrón se las
    # traga junto con la celda siguiente y esa se queda sin valor.
    xml = re.sub(r"<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>", sustituir, xml, flags=re.DOTALL)
    return xml, inyectadas


def procesar(ruta):
    ruta = Path(ruta)
    print(f"Evaluando {ruta.name} ...")
    valores, errores = evaluar(ruta)
    print(f"  celdas calculadas : {len(valores)}")
    print(f"  errores de fórmula: {len(errores)}")
    for h, c, v in errores[:20]:
        print(f"     {h}!{c} -> {v}")
    if errores:
        print("  ABORTADO: corrige los errores antes de inyectar.")
        return False

    copia = ruta.with_suffix(".bak.xlsx")
    shutil.copy2(ruta, copia)

    with zipfile.ZipFile(copia) as z:
        hojas = mapa_hojas(z)
        contenido = {n: z.read(n) for n in z.namelist()}

    total = 0
    for nombre_hoja, fichero in hojas.items():
        if fichero not in contenido:
            continue
        celdas = {c: v for (h, c), v in valores.items() if h == nombre_hoja}
        if not celdas:
            continue
        xml = contenido[fichero].decode("utf-8")
        xml, n = inyectar_hoja(xml, celdas)
        contenido[fichero] = xml.encode("utf-8")
        total += n
        print(f"  {nombre_hoja:<16} {n:>4} valores inyectados")

    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, datos in contenido.items():
            z.writestr(nombre, datos)

    copia.unlink()
    print(f"  TOTAL: {total} valores inyectados en {ruta.name}\n")
    return True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("Uso: python3 scripts/inyectar_valores.py <archivo.xlsx> [...]")
    for arg in sys.argv[1:]:
        procesar(arg)
