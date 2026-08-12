"""
Construye el modelo financiero del caso DTI-1399 ALCOPALET (Programa Lydes 2026).

Todos los datos de entrada proceden literalmente del enunciado y sus tres anexos.
Los resultados se calculan con fórmulas de Excel, nunca con valores precalculados,
para que el modelo recalcule al mover cualquier hipótesis.

Convenio de colores (estándar de modelización financiera):
    AZUL   = dato de entrada tomado del caso (editable)
    NEGRO  = fórmula calculada en la propia hoja
    VERDE  = enlace a otra hoja del libro
    AMARILLO (relleno) = hipótesis propia del analista, no dada por el caso
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Estilos
# --------------------------------------------------------------------------
FUENTE = "Arial"

TITULO = Font(name=FUENTE, size=14, bold=True, color="1F3864")
SUBTITULO = Font(name=FUENTE, size=9, italic=True, color="595959")
SECCION = Font(name=FUENTE, size=11, bold=True, color="FFFFFF")
CABECERA = Font(name=FUENTE, size=10, bold=True, color="FFFFFF")
ENTRADA = Font(name=FUENTE, size=10, color="0000FF")
FORMULA = Font(name=FUENTE, size=10, color="000000")
ENLACE = Font(name=FUENTE, size=10, color="008000")
ETIQUETA = Font(name=FUENTE, size=10, color="000000")
NOTA = Font(name=FUENTE, size=8, italic=True, color="808080")
TOTAL = Font(name=FUENTE, size=10, bold=True, color="000000")
TOTAL_ENL = Font(name=FUENTE, size=10, bold=True, color="008000")
DESTACADO = Font(name=FUENTE, size=11, bold=True, color="C00000")

F_SECCION = PatternFill("solid", fgColor="1F3864")
F_CABECERA = PatternFill("solid", fgColor="4472C4")
F_HIPOTESIS = PatternFill("solid", fgColor="FFFF00")
F_TOTAL = PatternFill("solid", fgColor="D9E2F3")
F_ALERTA = PatternFill("solid", fgColor="FCE4E4")
F_OK = PatternFill("solid", fgColor="E2EFDA")

linea_fina = Side(style="thin", color="BFBFBF")
BORDE = Border(left=linea_fina, right=linea_fina, top=linea_fina, bottom=linea_fina)
BORDE_SUP = Border(top=Side(style="thin", color="404040"))

EUR = '#,##0 "€";(#,##0) "€";"-"'
EUR2 = '#,##0.00 "€";(#,##0.00) "€";"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
NUM = '#,##0;(#,##0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'

wb = Workbook()


def titulo_hoja(ws, texto, subtitulo=None, ancho=8):
    ws["A1"] = texto
    ws["A1"].font = TITULO
    if subtitulo:
        ws["A2"] = subtitulo
        ws["A2"].font = SUBTITULO
    ws.sheet_view.showGridLines = False


def seccion(ws, fila, texto, ancho=6):
    ws.cell(fila, 1, texto).font = SECCION
    for c in range(1, ancho + 1):
        ws.cell(fila, c).fill = F_SECCION
    return fila + 1


def cabecera(ws, fila, valores, col_ini=1):
    for i, v in enumerate(valores):
        c = ws.cell(fila, col_ini + i, v)
        c.font = CABECERA
        c.fill = F_CABECERA
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    return fila + 1


def fila_dato(ws, fila, etiqueta, valor, fmt=None, fuente=ENTRADA, nota=None,
              col_etiqueta=1, col_valor=2, col_nota=3, hipotesis=False):
    ws.cell(fila, col_etiqueta, etiqueta).font = ETIQUETA
    c = ws.cell(fila, col_valor, valor)
    c.font = fuente
    if fmt:
        c.number_format = fmt
    if hipotesis:
        c.fill = F_HIPOTESIS
    if nota:
        n = ws.cell(fila, col_nota, nota)
        n.font = NOTA
    return fila + 1


def anchos(ws, mapa):
    for col, w in mapa.items():
        ws.column_dimensions[col].width = w


# ==========================================================================
# HOJA: DATOS
# ==========================================================================
ws = wb.active
ws.title = "Datos"
titulo_hoja(ws, "ALCOPALET (DTI-1399) — Datos e hipótesis del caso",
            "Programa Lydes 2026. Azul = dato del enunciado · Negro = fórmula · "
            "Amarillo = hipótesis propia del analista (no dada por el caso).")
anchos(ws, {"A": 46, "B": 16, "C": 74})

R = {}  # referencias nombradas -> "Datos!$B$n"


def dato(fila, clave, etiqueta, valor, fmt=None, fuente=ENTRADA, nota=None, hip=False):
    R[clave] = f"Datos!$B${fila}"
    return fila_dato(ws, fila, etiqueta, valor, fmt, fuente, nota, hipotesis=hip)


f = 4
f = seccion(ws, f, "A. PARÁMETROS GENERALES DE CÁLCULO", 3)
f = dato(f, "meses", "Meses productivos al año", 11, NUM,
         nota="Nota final del Anexo 3 del caso.")
f = dato(f, "dias_mes", "Días productivos al mes", 22, NUM,
         nota="Nota final del Anexo 3 del caso.")
f = dato(f, "dias_anio", "Días productivos al año", f"={R['meses']}*{R['dias_mes']}",
         NUM, FORMULA, "Calculado: 11 × 22 = 242 días.")
f = dato(f, "horas", "Horas por jornada", 8, NUM,
         nota="Turno único de 7:00 a 15:00. Esteban descarta turnos de tarde y noche.")
f = dato(f, "coste_mo", "Coste medio de mano de obra (€/año)", 38000, EUR,
         nota="Sueldos y cotizaciones sociales. Apartado (d) del caso.")
f = dato(f, "coste_desp", "Coste medio de despido (€/trabajador)", 30000, EUR,
         nota="Elevado por la antigüedad de la plantilla. Apartado (d) del caso.")
f = dato(f, "interes", "Tipo de interés de la financiación", 0.075, PCT,
         nota="Ofrecido por la entidad financiera habitual. Apartado (c) del caso.")
f += 1

f = seccion(ws, f, "B. CONTRATO PERSÁN", 3)
f = dato(f, "vol_persan", "Volumen anual comprometido (palets)", 90000, NUM,
         nota="90 mil unidades/año.")
f = dato(f, "precio_persan", "Precio de venta unitario (€)", 10.00, EUR2,
         nota="PVunit = 10 € FIJO durante los 5 años. Sin cláusula de revisión.")
f = dato(f, "anios_persan", "Duración del contrato (años)", 5, NUM)
f = dato(f, "dias_cobro", "Plazo de cobro (días)", 180, NUM,
         nota="Pago a 180 días.")
f = dato(f, "reten", "Retén permanente exigido (palets)", 1000, NUM,
         nota="Dos camiones en nave para garantizar servicio en menos de 24 horas.")
f += 1

f = seccion(ws, f, "C. LÍNEA AUTOMÁTICA (instalada en 2019)", 3)
f = dato(f, "linea_hora", "Capacidad nominal (palets/hora)", 65, NUM)
f = dato(f, "linea_rend", "Rendimiento considerado", 0.75, PCT,
         nota="Alcopalet la considera al 75% por su mantenimiento complejo y constante.")
f = dato(f, "linea_oper", "Operarios necesarios", 2, NUM)
f = dato(f, "linea_coste", "Coste de adquisición (€)", 360000, EUR,
         nota="Nota ** del Anexo 1.")
f = dato(f, "linea_amort_anios", "Periodo de amortización (años)", 10, NUM)
f = dato(f, "linea_amort", "Amortización anual (€)",
         f"={R['linea_coste']}/{R['linea_amort_anios']}", EUR, FORMULA,
         "36.000 € frente a los 36.143 € del Anexo 1: la diferencia son otros inmovilizados menores.")
f = dato(f, "linea_cambio_min", "Cambio de formato — mínimo (horas)", 4, NUM2)
f = dato(f, "linea_cambio_max", "Cambio de formato — máximo (horas)", 6, NUM2,
         nota="Parada total de la línea. Es la gran rigidez de esta instalación (38 metros).")
f += 1

f = seccion(ws, f, "D. FABRICACIÓN MANUAL (palet nuevo)", 3)
f = dato(f, "man_oper", "Operarios", 8, NUM)
f = dato(f, "man_dia", "Producción total (palets/día)", 500, NUM)
f = dato(f, "man_prod", "Productividad (palets/operario/día)",
         f"={R['man_dia']}/{R['man_oper']}", NUM2, FORMULA,
         "Dato derivado, necesario para valorar la sustitución por el robot.")
f += 1

f = seccion(ws, f, "E. LÍNEA DE PALET USADO (Palet Ojeda)", 3)
f = dato(f, "usa_oper", "Operarios", 7, NUM)
f = dato(f, "usa_prod", "Productividad (palets/persona/día)", 155, NUM)
f = dato(f, "usa_techo_mes", "Techo logístico (palets/mes)", 30000, NUM,
         nota="Techo por FALTA DE ESPACIO, no por falta de demanda. Clave del caso.")
f = dato(f, "usa_techo_anio", "Techo logístico (palets/año)",
         f"={R['usa_techo_mes']}*12", NUM, FORMULA)
f += 1

f = seccion(ws, f, "F. NUEVO ROBOT PARA PALETS", 3)
f = dato(f, "rob_inv", "Inversión (€)", 40000, EUR)
f = dato(f, "rob_dia", "Producción (palets/jornada)", 400, NUM)
f = dato(f, "rob_rend", "Rendimiento considerado", 0.90, PCT,
         nota="Frente al 75% de la línea de 2019: sin electrónica compleja ni rodamientos.")
f = dato(f, "rob_oper", "Operarios necesarios", 1, NUM,
         nota="Un solo operario dedicado en exclusiva a la máquina (nota 1 del Anexo 3).")
f = dato(f, "rob_cambio", "Cambio de formato (minutos)", 25, NUM,
         nota="Entre 20 y 30 minutos, frente a 4-6 HORAS de la línea de 2019.")
f += 1

f = seccion(ws, f, "G. INVERSIONES ADICIONALES SI SE ACEPTA PERSÁN", 3)
f = dato(f, "campa_mes", "Alquiler almacén/campa (€/mes)", 5000, EUR,
         nota="Terreno colindante. Apartado (a) del caso.")
f = dato(f, "campa_anio", "Alquiler almacén/campa (€/año)",
         f"={R['campa_mes']}*12", EUR, FORMULA)
f = dato(f, "trailer", "Nuevo tráiler y remolque (€)", 150000, EUR,
         nota="La logística actual ya está muy saturada. Apartado (b) del caso.")
f = dato(f, "leasing_mes", "Leasing máquina + camión (€/mes)", 4000, EUR,
         nota="Cuota conjunta para ambos activos. Apartado (c) del caso.")
f = dato(f, "leasing_anio", "Leasing máquina + camión (€/año)",
         f"={R['leasing_mes']}*12", EUR, FORMULA)
f = dato(f, "leasing_anios", "Duración del leasing (años)", 5, NUM)
f = dato(f, "inv_total", "Inversión total financiada (€)",
         f"={R['rob_inv']}+{R['trailer']}", EUR, FORMULA,
         "40.000 € (robot) + 150.000 € (tráiler). La cuota de 4.000 €/mes × 60 meses = 240.000 €.")
f += 1

f = seccion(ws, f, "H. HIPÓTESIS PROPIAS DEL ANALISTA (no dadas por el caso)", 3)
ws.cell(f, 1, "Estas celdas NO proceden del enunciado. Son los supuestos que hay que "
              "declarar y defender ante el jurado.").font = NOTA
f += 1
f = dato(f, "monetiz", "% de mano de obra liberada que se monetiza", 0.00, PCT, FORMULA,
         "CRÍTICO. Esteban NO quiere despidos. Si el personal liberado por el robot no se "
         "reasigna a producción vendible, el ahorro teórico de mano de obra es cero.", hip=True)
f = dato(f, "precio_reneg", "Precio Persán renegociado objetivo (€/ud)", 12.50, EUR2, FORMULA,
         "Hipótesis de negociación. Ver hoja Persan para el precio de equilibrio exacto.", hip=True)
f = dato(f, "dias_cobro_reneg", "Plazo de cobro renegociado (días)", 90, NUM, FORMULA,
         "Hipótesis de negociación: reducir de 180 a 90 días libera la mitad del circulante.", hip=True)
f = dato(f, "infl_madera", "Inflación anual del precio de la madera", 0.00, PCT, FORMULA,
         "Escenario base 0%. La madera es el 83% del precio Persán y el contrato NO tiene "
         "cláusula de revisión durante 5 años.", hip=True)
f = dato(f, "dividendo", "Dividendo extraordinario planteado (€)", 200000, EUR, FORMULA,
         "El caso no cuantifica el importe de la casa. Cifra de trabajo, a ajustar.", hip=True)

FIN_DATOS = f

# ==========================================================================
# HOJA: PyG 2025 (Anexo 1)
# ==========================================================================
ws = wb.create_sheet("PyG2025")
titulo_hoja(ws, "Anexo 1 — Cuenta de pérdidas y ganancias consolidada 2025",
            "Los subtotales se recalculan con fórmulas y se contrastan contra el valor "
            "publicado en el anexo (columna de diferencia).")
anchos(ws, {"A": 44, "B": 16, "C": 12, "D": 16, "E": 12, "F": 50})

f = 4
f = cabecera(ws, f, ["Concepto", "Importe 2025", "% s/ Ventas",
                     "Valor del anexo", "Diferencia", "Comentario"])

PYG = {}


def pyg(fila, clave, etiqueta, valor, anexo=None, negrita=False, comentario=None,
        es_formula=False):
    PYG[clave] = f"PyG2025!$B${fila}"
    c_et = ws.cell(fila, 1, etiqueta)
    c_va = ws.cell(fila, 2, valor)
    c_va.number_format = EUR
    c_va.font = TOTAL if negrita else (FORMULA if es_formula else ENTRADA)
    c_et.font = TOTAL if negrita else ETIQUETA
    c_pc = ws.cell(fila, 3, f"=B{fila}/$B${VENTAS_FILA}")
    c_pc.number_format = PCT
    c_pc.font = FORMULA
    if negrita:
        for col in range(1, 6):
            ws.cell(fila, col).fill = F_TOTAL
    if anexo is not None:
        ws.cell(fila, 4, anexo).number_format = EUR
        ws.cell(fila, 4).font = ENTRADA
        d = ws.cell(fila, 5, f"=B{fila}-D{fila}")
        d.number_format = EUR
        d.font = FORMULA
    if comentario:
        ws.cell(fila, 6, comentario).font = NOTA
    return fila + 1


VENTAS_FILA = f
f = pyg(f, "ventas", "VENTAS NETAS", 4585267, negrita=True)
f = pyg(f, "var_exist", "Variación de existencias", 376062,
        comentario="Aumento de existencias: es beneficio contable, pero NO es caja.")
f = pyg(f, "aprov", "Aprovisionamientos", -3068506,
        comentario="Compras del ejercicio. El consumo real es Aprovisionamientos − Variación.")
f = pyg(f, "mb", "MARGEN BRUTO", f"=SUM(B{VENTAS_FILA}:B{f-1})", 1892822,
        negrita=True, es_formula=True)
MB_FILA = f - 1
f = pyg(f, "pers_prod", "Gastos de Personal — Producción", -481660)
f = pyg(f, "transp", "Gastos de Transporte", -268218,
        comentario="Coste total de los 2 camiones actuales, YA al 100% de su capacidad.")
f = pyg(f, "pers_estr", "Gastos de Personal — Estructura/Admin.", -311292)
f = pyg(f, "otros", "Otros Costes de Explotación", -378979,
        comentario="El 60% corresponde a la línea de palet nuevo (apartado (e) del caso).")
f = pyg(f, "ebitda", "EBITDA", f"=B{MB_FILA}+SUM(B{MB_FILA+1}:B{f-1})", 452674,
        negrita=True, es_formula=True)
EBITDA_FILA = f - 1
f = pyg(f, "amort", "Amortización del Inmovilizado", -36143,
        comentario="Línea de 2019: 360.000 € a 10 años = 36.000 €/año.")
f = pyg(f, "ebit", "RESULTADO DE EXPLOTACIÓN (EBIT)",
        f"=B{EBITDA_FILA}+B{EBITDA_FILA+1}", 416531, negrita=True, es_formula=True)
EBIT_FILA = f - 1
f = pyg(f, "gfin", "Gastos Financieros", -44046)
f = pyg(f, "bai", "RESULTADO ANTES DE IMPUESTOS", f"=B{EBIT_FILA}+B{EBIT_FILA+1}",
        372485, negrita=True, es_formula=True)
BAI_FILA = f - 1
f = pyg(f, "imp", "Impuestos (estimado)", -93121)
f = pyg(f, "rn", "RESULTADO NETO DEL EJERCICIO", f"=B{BAI_FILA}+B{BAI_FILA+1}",
        279364, negrita=True, es_formula=True)

f += 1
f = seccion(ws, f, "RATIOS Y LECTURA DE CAJA", 6)
f = fila_dato(ws, f, "Tipo impositivo efectivo", f"={PYG['imp']}/-{PYG['bai']}", PCT,
              FORMULA, "Coherente con un tipo del 25%.")
f = fila_dato(ws, f, "Margen EBITDA sobre ventas", f"={PYG['ebitda']}/{PYG['ventas']}",
              PCT, FORMULA)
f = fila_dato(ws, f, "Margen neto sobre ventas", f"={PYG['rn']}/{PYG['ventas']}",
              PCT, FORMULA)
f += 1
ws.cell(f, 1, "Conversión del EBITDA en caja (aproximación)").font = SECCION
for c in range(1, 7):
    ws.cell(f, c).fill = F_SECCION
f += 1
f = fila_dato(ws, f, "EBITDA", f"={PYG['ebitda']}", EUR, ENLACE)
f = fila_dato(ws, f, "(−) Aumento de existencias", f"=-{PYG['var_exist']}", EUR, ENLACE,
              "Salida de caja: se ha fabricado y comprado más de lo vendido.")
f = fila_dato(ws, f, "(−) Gastos financieros", f"={PYG['gfin']}", EUR, ENLACE)
f = fila_dato(ws, f, "(−) Impuestos", f"={PYG['imp']}", EUR, ENLACE)
CAJA_FILA = f
c = ws.cell(f, 1, "CAJA OPERATIVA APROXIMADA 2025")
c.font = TOTAL
cv = ws.cell(f, 2, f"=B{f-4}+B{f-3}+B{f-2}+B{f-1}")
cv.number_format = EUR
cv.font = TOTAL
for col in range(1, 7):
    ws.cell(f, col).fill = F_ALERTA
ws.cell(f, 6, "Frente a un beneficio contable de 279.364 €. El beneficio no es caja "
              "disponible: el argumento central frente a la petición de dividendo.").font = NOTA
CAJA_OP = f"PyG2025!$B${f}"
f += 2
ws.cell(f, 1, "Nota: el aumento de existencias puede ser puntual (acopio de madera gallega). "
              "Aun así, muestra que 2025 generó mucha menos caja libre de la que sugiere el "
              "beneficio.").font = NOTA

# ==========================================================================
# HOJA: LINEAS (Anexo 2)
# ==========================================================================
ws = wb.create_sheet("Lineas")
titulo_hoja(ws, "Anexo 2 — Rentabilidad por líneas de negocio 2025",
            "Comprobación de que las dos líneas suman exactamente el consolidado del Anexo 1.")
anchos(ws, {"A": 40, "B": 18, "C": 18, "D": 18, "E": 16, "F": 46})

f = 4
f = cabecera(ws, f, ["Concepto", "ALCOPALET (Nuevo)", "PALET OJEDA (Usado)",
                     "TOTAL", "Anexo 1", "Diferencia"])
LIN = {}
fila_v = f
LIN["v_nuevo"] = f"Lineas!$B${f}"
LIN["v_usado"] = f"Lineas!$C${f}"
ws.cell(f, 1, "Ventas netas (€)").font = ETIQUETA
for col, val in ((2, 2645224), (3, 1940042)):
    c = ws.cell(f, col, val)
    c.font = ENTRADA
    c.number_format = EUR
ws.cell(f, 4, f"=B{f}+C{f}").font = TOTAL
ws.cell(f, 4).number_format = EUR
ws.cell(f, 5, f"={PYG['ventas']}").font = ENLACE
ws.cell(f, 5).number_format = EUR
ws.cell(f, 6, f"=D{f}-E{f}").font = FORMULA
ws.cell(f, 6).number_format = EUR
f += 1

fila_r = f
LIN["r_nuevo"] = f"Lineas!$B${f}"
LIN["r_usado"] = f"Lineas!$C${f}"
ws.cell(f, 1, "Resultado neto (€)").font = ETIQUETA
for col, val in ((2, 158882), (3, 120482)):
    c = ws.cell(f, col, val)
    c.font = ENTRADA
    c.number_format = EUR
ws.cell(f, 4, f"=B{f}+C{f}").font = TOTAL
ws.cell(f, 4).number_format = EUR
ws.cell(f, 5, f"={PYG['rn']}").font = ENLACE
ws.cell(f, 5).number_format = EUR
ws.cell(f, 6, f"=D{f}-E{f}").font = FORMULA
ws.cell(f, 6).number_format = EUR
f += 1

fila_u = f
LIN["u_nuevo"] = f"Lineas!$B${f}"
LIN["u_usado"] = f"Lineas!$C${f}"
ws.cell(f, 1, "Unidades vendidas (palets)").font = ETIQUETA
for col, val in ((2, 174504), (3, 260330)):
    c = ws.cell(f, col, val)
    c.font = ENTRADA
    c.number_format = NUM
ws.cell(f, 4, f"=B{f}+C{f}").font = TOTAL
ws.cell(f, 4).number_format = NUM
f += 2

f = seccion(ws, f, "INDICADORES DERIVADOS", 6)
for etiqueta, formula, fmt, nota in [
    ("Precio medio de venta (€/palet)", f"=B{fila_v}/B{fila_u}", EUR2,
     "Coincide con el ingreso medio del Anexo 3 (15,16 € y 7,45 €): los anexos son coherentes."),
    ("Resultado neto por palet (€)", f"=B{fila_r}/B{fila_u}", EUR2, None),
    ("Margen neto sobre ventas", f"=B{fila_r}/B{fila_v}", PCT, None),
    ("Peso sobre ventas totales", f"=B{fila_v}/$D${fila_v}", PCT, None),
    ("Peso sobre resultado total", f"=B{fila_r}/$D${fila_r}", PCT,
     "El palet usado aporta el 43% del beneficio con el 42% de las ventas."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    for col, letra in ((2, "B"), (3, "C")):
        c = ws.cell(f, col, formula.replace("B", letra) if letra != "B" else formula)
        c.font = FORMULA
        c.number_format = fmt
    if nota:
        ws.cell(f, 6, nota).font = NOTA
    f += 1
PRECIO_NUEVO = f"Lineas!$B${f-5}"

# ==========================================================================
# HOJA: ESCANDALLO (Anexo 3)
# ==========================================================================
ws = wb.create_sheet("Escandallo")
titulo_hoja(ws, "Anexo 3 — Escandallo de costes unitarios (€/palet)",
            "El coste y el margen se recalculan sumando los componentes, en lugar de copiar "
            "el total del anexo. Así se detecta cualquier redondeo.")
anchos(ws, {"A": 38, "B": 17, "C": 17, "D": 19, "E": 15, "F": 52})

f = 4
f = cabecera(ws, f, ["Concepto de coste", "Manual actual (2025)", "Palet usado (2025)",
                     "Proyecto Persán", "", "Notas del anexo"])
ESC = {}
fila_ing = f
for clave, etiqueta, vals, nota in [
    ("ing", "Ingreso (medio)", (15.16, 7.45, 10.00),
     "El precio Persán está un 34% por debajo del precio medio actual del palet nuevo."),
    ("mp", "Materia prima (madera y clavos)", (8.27, 4.80, 8.27),
     "IDÉNTICA en manual y Persán: el robot no reduce el coste de la madera."),
    ("mod", "Mano de obra directa (MOD)", (2.18, 1.02, 0.42),
     "(1) 90.000 ud/año con un solo operario: 38.000 € / 90.000 = 0,42 €."),
    ("tr", "Gastos de transporte", (0.62, 0.62, 0.53),
     "(2) Considera el nuevo tráiler propio a plena carga en las rutas de Persán."),
    ("otros", "Otros costes de explotación variables", (1.30, 0.58, 1.30),
     "(3) Varían proporcionalmente al volumen (electricidad, mantenimiento, etc.)."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    ESC[clave] = {}
    for i, (col, linea) in enumerate(((2, "man"), (3, "usa"), (4, "per"))):
        c = ws.cell(f, col, vals[i])
        c.font = ENTRADA
        c.number_format = EUR2
        ESC[clave][linea] = f"Escandallo!${get_column_letter(col)}${f}"
    ws.cell(f, 6, nota).font = NOTA
    f += 1

fila_coste = f
ws.cell(f, 1, "COSTE DE PRODUCCIÓN UNITARIO").font = TOTAL
for col in (2, 3, 4):
    L = get_column_letter(col)
    c = ws.cell(f, col, f"=SUM({L}{fila_ing+1}:{L}{f-1})")
    c.font = TOTAL
    c.number_format = EUR2
for col in range(1, 7):
    ws.cell(f, col).fill = F_TOTAL
ws.cell(f, 6, "El anexo publica 12,36 € en manual; la suma exacta da 12,37 €. "
              "Redondeo de 0,01 €, sin efecto material.").font = NOTA
ESC["coste"] = {l: f"Escandallo!${get_column_letter(c)}${f}"
                for c, l in ((2, "man"), (3, "usa"), (4, "per"))}
f += 1

fila_margen = f
ws.cell(f, 1, "MARGEN DE CONTRIBUCIÓN UNITARIO").font = TOTAL
for col in (2, 3, 4):
    L = get_column_letter(col)
    c = ws.cell(f, col, f"={L}{fila_ing}-{L}{fila_coste}")
    c.font = TOTAL
    c.number_format = EUR2
for col in range(1, 5):
    ws.cell(f, col).fill = F_TOTAL
ws.cell(f, 4).fill = F_ALERTA
ws.cell(f, 4).font = DESTACADO
ws.cell(f, 6, "HALLAZGO CENTRAL: el margen del proyecto Persán es NEGATIVO. "
              "Ni siquiera cubre los costes variables.").font = Font(
    name=FUENTE, size=9, italic=True, bold=True, color="C00000")
ESC["margen"] = {l: f"Escandallo!${get_column_letter(c)}${f}"
                 for c, l in ((2, "man"), (3, "usa"), (4, "per"))}
f += 1

ws.cell(f, 1, "Margen sobre precio de venta").font = ETIQUETA
for col in (2, 3, 4):
    L = get_column_letter(col)
    c = ws.cell(f, col, f"={L}{fila_margen}/{L}{fila_ing}")
    c.font = FORMULA
    c.number_format = PCT
f += 1
ws.cell(f, 1, "Peso de la materia prima sobre el precio").font = ETIQUETA
for col in (2, 3, 4):
    L = get_column_letter(col)
    c = ws.cell(f, col, f"={L}{fila_ing+1}/{L}{fila_ing}")
    c.font = FORMULA
    c.number_format = PCT
ws.cell(f, 6, "En Persán la madera se come el 83% del precio. Sin cláusula de revisión "
              "a 5 años, es el riesgo principal del contrato.").font = NOTA
f += 2

# --- Reconciliación con el Anexo 1 ---
f = seccion(ws, f, "RECONCILIACIÓN DEL ESCANDALLO CON LA CUENTA DE RESULTADOS", 6)
f = cabecera(ws, f, ["Partida", "Según escandallo", "Según Anexo 1", "Diferencia",
                     "% desvío", "Lectura"])
for etiqueta, form_esc, form_pyg, nota in [
    ("Materia prima (consumo real)",
     f"={ESC['mp']['man']}*{LIN['u_nuevo']}+{ESC['mp']['usa']}*{LIN['u_usado']}",
     f"=-({PYG['aprov']}+{PYG['var_exist']})",
     "Consumo = Aprovisionamientos − Variación de existencias. Cuadra al 0,01%."),
    ("Gastos de transporte",
     f"={ESC['tr']['man']}*{LIN['u_nuevo']}+{ESC['tr']['usa']}*{LIN['u_usado']}",
     f"=-{PYG['transp']}", "Cuadra."),
    ("Otros costes de explotación",
     f"={ESC['otros']['man']}*{LIN['u_nuevo']}+{ESC['otros']['usa']}*{LIN['u_usado']}",
     f"=-{PYG['otros']}", "Cuadra."),
    ("Mano de obra directa",
     f"={ESC['mod']['man']}*{LIN['u_nuevo']}+{ESC['mod']['usa']}*{LIN['u_usado']}",
     f"=-{PYG['pers_prod']}",
     "NO cuadra: el escandallo recoge 17 operarios × 38.000 € = 646.000 €, mientras "
     "el Anexo 1 imputa 481.660 € a Producción. Parte de la plantilla directa está "
     "clasificada en Estructura. No invalida el escandallo."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, form_esc); c.font = FORMULA; c.number_format = EUR
    c = ws.cell(f, 3, form_pyg); c.font = ENLACE; c.number_format = EUR
    c = ws.cell(f, 4, f"=B{f}-C{f}"); c.font = FORMULA; c.number_format = EUR
    c = ws.cell(f, 5, f"=IF(C{f}=0,0,(B{f}-C{f})/C{f})"); c.font = FORMULA; c.number_format = PCT
    ws.cell(f, 6, nota).font = NOTA
    f += 1

f += 1
ws.cell(f, 1, "Comprobación de plantilla directa implícita en el escandallo").font = SECCION
for c in range(1, 7):
    ws.cell(f, c).fill = F_SECCION
f += 1
for etiqueta, formula, nota in [
    ("Operarios implícitos — palet nuevo",
     f"={ESC['mod']['man']}*{LIN['u_nuevo']}/{R['coste_mo']}",
     "= 10 operarios: los 8 de manual + los 2 de la línea automática. Coherente."),
    ("Operarios implícitos — palet usado",
     f"={ESC['mod']['usa']}*{LIN['u_usado']}/{R['coste_mo']}",
     "= 7 operarios. Coincide exactamente con el dato del caso. El escandallo es fiable."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, formula); c.font = FORMULA; c.number_format = NUM2
    ws.cell(f, 6, nota).font = NOTA
    f += 1

# ==========================================================================
# HOJA: CAPACIDAD
# ==========================================================================
ws = wb.create_sheet("Capacidad")
titulo_hoja(ws, "Capacidad productiva — el dato que el caso no da hecho",
            "El caso da rendimientos por hora y por jornada, pero nunca la capacidad anual. "
            "Sin ella no se puede saber si Persán cabe en la fábrica.")
anchos(ws, {"A": 46, "B": 16, "C": 14, "D": 14, "E": 14, "F": 56})

f = 4
f = seccion(ws, f, "A. CAPACIDAD ACTUAL DE LA LÍNEA DE PALET NUEVO", 6)
f = cabecera(ws, f, ["Recurso", "Palets/día", "Días/año", "Palets/año", "Operarios", "Cálculo"])

fila_linea = f
ws.cell(f, 1, "Línea automática (2019)").font = ETIQUETA
c = ws.cell(f, 2, f"={R['linea_hora']}*{R['linea_rend']}*{R['horas']}")
c.font = FORMULA; c.number_format = NUM
c = ws.cell(f, 3, f"={R['dias_anio']}"); c.font = ENLACE; c.number_format = NUM
c = ws.cell(f, 4, f"=B{f}*C{f}"); c.font = FORMULA; c.number_format = NUM
c = ws.cell(f, 5, f"={R['linea_oper']}"); c.font = ENLACE; c.number_format = NUM
ws.cell(f, 6, "65 palets/h × 75% de rendimiento × 8 h de jornada.").font = NOTA
f += 1

fila_manual = f
ws.cell(f, 1, "Fabricación manual").font = ETIQUETA
c = ws.cell(f, 2, f"={R['man_dia']}"); c.font = ENLACE; c.number_format = NUM
c = ws.cell(f, 3, f"={R['dias_anio']}"); c.font = ENLACE; c.number_format = NUM
c = ws.cell(f, 4, f"=B{f}*C{f}"); c.font = FORMULA; c.number_format = NUM
c = ws.cell(f, 5, f"={R['man_oper']}"); c.font = ENLACE; c.number_format = NUM
ws.cell(f, 6, "500 palets/día entre 8 operarios = 62,5 palets por operario y día.").font = NOTA
f += 1

fila_cap_act = f
ws.cell(f, 1, "CAPACIDAD ACTUAL TOTAL").font = TOTAL
c = ws.cell(f, 2, f"=B{fila_linea}+B{fila_manual}"); c.font = TOTAL; c.number_format = NUM
c = ws.cell(f, 4, f"=D{fila_linea}+D{fila_manual}"); c.font = TOTAL; c.number_format = NUM
c = ws.cell(f, 5, f"=E{fila_linea}+E{fila_manual}"); c.font = TOTAL; c.number_format = NUM
for col in range(1, 7):
    ws.cell(f, col).fill = F_TOTAL
CAP_ACTUAL = f"Capacidad!$D${f}"
f += 2

f = seccion(ws, f, "B. GRADO DE UTILIZACIÓN ACTUAL", 6)
for etiqueta, formula, fmt, fuente, nota in [
    ("Producción real 2025 (palet nuevo)", f"={LIN['u_nuevo']}", NUM, ENLACE, None),
    ("Capacidad teórica anual", f"={CAP_ACTUAL}", NUM, ENLACE, None),
    ("GRADO DE UTILIZACIÓN", f"=B{f}/B{f+1}", PCT, TOTAL,
     "La fábrica NO está llena: hay recorrido antes de invertir."),
    ("Capacidad ociosa disponible hoy", f"=B{f+1}-B{f}", NUM, TOTAL,
     "Casi 41.000 palets/año que hoy se pueden fabricar sin invertir un euro. "
     "Esto debilita el argumento de que la máquina es imprescindible."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA if fuente != TOTAL else TOTAL
    c = ws.cell(f, 2, formula); c.font = fuente; c.number_format = fmt
    if nota:
        ws.cell(f, 6, nota).font = NOTA
    f += 1
CAP_OCIOSA = f"Capacidad!$B${f-1}"
f += 1

f = seccion(ws, f, "C. EL ROBOT Y EL ENCAJE DEL CONTRATO PERSÁN", 6)
fila_rob = f
for etiqueta, formula, fmt, fuente, nota in [
    ("Robot — producción efectiva (palets/día)",
     f"={R['rob_dia']}*{R['rob_rend']}", NUM, FORMULA, "400 palets/jornada × 90% de rendimiento."),
    ("Robot — capacidad anual (palets)",
     f"=B{f}*{R['dias_anio']}", NUM, TOTAL, None),
    ("Persán — necesidad diaria (palets/día)",
     f"={R['vol_persan']}/{R['dias_anio']}", NUM2, FORMULA, "90.000 / 242 días."),
    ("Persán — necesidad anual (palets)", f"={R['vol_persan']}", NUM, ENLACE, None),
    ("DÉFICIT/SUPERÁVIT del robot frente a Persán",
     f"=B{f+1}-B{f+3}", NUM, DESTACADO,
     "El robot en solitario NO llega a las 90.000 unidades: se queda un 3,2% corto. "
     "Habría que completarlo con las otras líneas o con días extra."),
    ("Déficit en % del compromiso", f"=B{f+4}/B{f+3}", PCT, FORMULA, None),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, formula); c.font = fuente; c.number_format = fmt
    if nota:
        ws.cell(f, 6, nota).font = NOTA
    f += 1
ROB_ANUAL = f"Capacidad!$B${fila_rob+1}"
f += 1

f = seccion(ws, f, "D. ¿CABE PERSÁN EN LA FÁBRICA? (con y sin robot)", 6)
f = cabecera(ws, f, ["Escenario", "Capacidad anual", "Demanda anual", "Holgura",
                     "¿Encaja?", "Lectura"])
for etiqueta, cap, dem, nota in [
    ("Sin robot, sin Persán", f"={CAP_ACTUAL}", f"={LIN['u_nuevo']}",
     "Situación actual: sobra capacidad."),
    ("Sin robot, CON Persán", f"={CAP_ACTUAL}", f"={LIN['u_nuevo']}+{R['vol_persan']}",
     "NO cabe: faltan unas 49.000 unidades. Hace falta capacidad nueva sí o sí."),
    ("Con robot, CON Persán", f"={CAP_ACTUAL}+{ROB_ANUAL}",
     f"={LIN['u_nuevo']}+{R['vol_persan']}",
     "Cabe con holgura. Técnicamente el proyecto es viable; el problema es económico."),
    ("Con robot, SIN Persán", f"={CAP_ACTUAL}+{ROB_ANUAL}", f"={LIN['u_nuevo']}",
     "Enorme capacidad ociosa: el robot solo tiene sentido si hay demanda que lo llene."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, cap); c.font = FORMULA; c.number_format = NUM
    c = ws.cell(f, 3, dem); c.font = FORMULA; c.number_format = NUM
    c = ws.cell(f, 4, f"=B{f}-C{f}"); c.font = TOTAL; c.number_format = NUM
    c = ws.cell(f, 5, f'=IF(D{f}>=0,"SÍ","NO")'); c.font = TOTAL
    c.alignment = Alignment(horizontal="center")
    ws.cell(f, 6, nota).font = NOTA
    f += 1
f += 1

f = seccion(ws, f, "E. LÍNEA DE PALET USADO: CRECIMIENTO BLOQUEADO POR ESPACIO", 6)
fila_usa = f
for etiqueta, formula, fmt, fuente, nota in [
    ("Producción actual (palets/año)", f"={LIN['u_usado']}", NUM, ENLACE, None),
    ("Producción actual (palets/mes)", f"=B{f}/12", NUM, FORMULA, None),
    ("Techo logístico (palets/año)", f"={R['usa_techo_anio']}", NUM, ENLACE,
     "30.000 unidades/mes. El límite es el ESPACIO, no la demanda."),
    ("Recorrido disponible (palets/año)", f"=B{f+2}-B{f}", NUM, TOTAL, None),
    ("Contribución del recorrido (€/año)",
     f"=B{f+3}*{ESC['margen']['usa']}", EUR, TOTAL,
     "Con margen de 0,43 €/ud. Por sí solo no paga la campa de 60.000 €/año, "
     "pero sí reparte su coste con una línea que SÍ gana dinero."),
    ("Operarios necesarios para alcanzar el techo",
     f"={R['usa_techo_anio']}/({R['dias_anio']}*{R['usa_prod']})", NUM2, FORMULA, None),
    ("Operarios adicionales frente a los 7 actuales",
     f"=B{f+5}-{R['usa_oper']}", NUM2, TOTAL,
     "OPORTUNIDAD: el robot libera personal de palet nuevo; esta línea necesita personal "
     "y solo le falta espacio. La campa de Persán resolvería ambas cosas."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, formula); c.font = fuente; c.number_format = fmt
    if nota:
        ws.cell(f, 6, nota).font = NOTA
    f += 1

# ==========================================================================
# HOJA: PERSAN
# ==========================================================================
ws = wb.create_sheet("Persan")
titulo_hoja(ws, "Análisis económico del contrato Persán",
            "90.000 unidades/año · 10 €/unidad fijo durante 5 años · cobro a 180 días · "
            "retén permanente de 1.000 palets.")
anchos(ws, {"A": 48, "B": 18, "C": 18, "D": 14, "E": 12, "F": 54})

f = 4
f = seccion(ws, f, "A. CUENTA DE RESULTADOS INCREMENTAL DEL CONTRATO (año completo)", 6)
fila_ing_p = f
for etiqueta, formula, fmt, fuente, negrita, nota in [
    ("Ingresos (90.000 ud × 10 €)", f"={R['vol_persan']}*{R['precio_persan']}", EUR, FORMULA, False, None),
    ("Costes variables (90.000 ud × 10,52 €)",
     f"=-{R['vol_persan']}*{ESC['coste']['per']}", EUR, FORMULA, False, None),
    ("MARGEN DE CONTRIBUCIÓN", f"=B{f}+B{f+1}", EUR, TOTAL, True,
     "Negativo ANTES de cualquier coste fijo. Ningún volumen arregla esto."),
    ("Alquiler de la campa", f"=-{R['campa_anio']}", EUR, ENLACE, False, "5.000 €/mes."),
    ("Leasing de máquina + camión", f"=-{R['leasing_anio']}", EUR, ENLACE, False, "4.000 €/mes."),
    ("RESULTADO ANTES DE COSTE DE CIRCULANTE", f"=B{f+2}+B{f+3}+B{f+4}", EUR, TOTAL, True, None),
]:
    ws.cell(f, 1, etiqueta).font = TOTAL if negrita else ETIQUETA
    c = ws.cell(f, 2, formula); c.font = fuente; c.number_format = fmt
    if negrita:
        for col in range(1, 7):
            ws.cell(f, col).fill = F_TOTAL
    if nota:
        ws.cell(f, 6, nota).font = NOTA
    f += 1
MC_PERSAN = f"Persan!$B${fila_ing_p+2}"
SUBTOTAL_P = f"Persan!$B${fila_ing_p+5}"
f += 1

f = seccion(ws, f, "B. NECESIDADES DE CIRCULANTE (el coste oculto del pago a 180 días)", 6)
fila_circ = f
for etiqueta, formula, fmt, fuente, nota in [
    ("Cuenta a cobrar media (€)",
     f"={R['vol_persan']}*{R['precio_persan']}*{R['dias_cobro']}/365", EUR, FORMULA,
     "Ingresos × 180/365. Dinero de Alcopalet en manos de Persán de forma permanente."),
    ("Retén permanente de 1.000 palets, a coste (€)",
     f"={R['reten']}*{ESC['coste']['per']}", EUR, FORMULA,
     "Stock inmovilizado que exige el cliente para servir en menos de 24 horas."),
    ("CIRCULANTE TOTAL INMOVILIZADO (€)", f"=B{f}+B{f+1}", EUR, TOTAL, None),
    ("Coste financiero anual del circulante (€)",
     f"=-B{f+2}*{R['interes']}", EUR, TOTAL,
     "Al 7,5% ofrecido por el banco. Es un coste real, aunque no aparezca en el escandallo."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, formula); c.font = fuente; c.number_format = fmt
    if nota:
        ws.cell(f, 6, nota).font = NOTA
    f += 1
COSTE_CIRC = f"Persan!$B${fila_circ+3}"
f += 1

f = seccion(ws, f, "C. IMPACTO TOTAL DEL CONTRATO", 6)
fila_imp = f
ws.cell(f, 1, "IMPACTO ANUAL TOTAL ESTIMADO").font = DESTACADO
c = ws.cell(f, 2, f"={SUBTOTAL_P}+{COSTE_CIRC}"); c.font = DESTACADO; c.number_format = EUR
for col in range(1, 7):
    ws.cell(f, col).fill = F_ALERTA
ws.cell(f, 6, "Margen negativo + campa + leasing + coste del circulante.").font = NOTA
IMPACTO_PERSAN = f"Persan!$B${f}"
f += 1
for etiqueta, formula, fmt, nota in [
    ("En % del resultado neto actual de la empresa", f"={IMPACTO_PERSAN}/{PYG['rn']}", PCT,
     "El contrato se llevaría por delante dos tercios del beneficio de TODA la empresa."),
    ("Impacto acumulado a 5 años", f"={IMPACTO_PERSAN}*{R['anios_persan']}", EUR,
     "Sin cláusula de revisión de precio, el daño se repite cada año."),
    ("Resultado neto de la empresa tras Persán", f"={PYG['rn']}+{IMPACTO_PERSAN}", EUR,
     "La empresa entraría en pérdidas o quedaría al borde."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, formula); c.font = TOTAL; c.number_format = fmt
    ws.cell(f, 6, nota).font = NOTA
    f += 1
f += 1

f = seccion(ws, f, "D. PRECIO DE EQUILIBRIO: ¿A CUÁNTO HABRÍA QUE VENDER?", 6)
f = cabecera(ws, f, ["Nivel de cobertura", "Precio mínimo (€/ud)", "Diferencia vs 10 €",
                     "% de subida", "", "Qué cubre"])
fila_be = f
for etiqueta, formula, nota in [
    ("1. Solo costes variables", f"={ESC['coste']['per']}",
     "Madera, mano de obra, transporte y otros variables."),
    ("2. + campa y leasing",
     f"={ESC['coste']['per']}+({R['campa_anio']}+{R['leasing_anio']})/{R['vol_persan']}",
     "Añade los costes fijos que el contrato obliga a asumir."),
    ("3. + coste del circulante (equilibrio real)",
     f"=({R['vol_persan']}*{ESC['coste']['per']}+{R['campa_anio']}+{R['leasing_anio']})"
     f"/({R['vol_persan']}*(1-{R['dias_cobro']}/365*{R['interes']}))",
     "PRECIO DE EQUILIBRIO VERDADERO. Resuelve el precio que hace el resultado igual a cero."),
    ("4. + margen equivalente al negocio actual",
     f"=B{f+1}+{ESC['margen']['man']}",
     "Para que Persán sea tan rentable como el resto del palet nuevo (2,79 €/ud)."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, formula); c.font = TOTAL; c.number_format = EUR2
    c = ws.cell(f, 3, f"=B{f}-{R['precio_persan']}"); c.font = FORMULA; c.number_format = EUR2
    c = ws.cell(f, 4, f"=B{f}/{R['precio_persan']}-1"); c.font = FORMULA; c.number_format = PCT
    ws.cell(f, 6, nota).font = NOTA
    f += 1
ws.cell(fila_be + 2, 1).fill = F_HIPOTESIS
ws.cell(fila_be + 2, 2).fill = F_HIPOTESIS
PRECIO_EQ = f"Persan!$B${fila_be+2}"
f += 1

f = seccion(ws, f, "E. RIESGO DEL PRECIO DE LA MADERA (83% del precio de venta)", 6)
f = cabecera(ws, f, ["Subida del precio de la madera", "Coste unitario (€)",
                     "Margen unitario (€)", "Impacto anual (€)", "", "Lectura"])
for infl in (0.00, 0.05, 0.10, 0.15, 0.20):
    ws.cell(f, 1, infl).font = ENTRADA
    ws.cell(f, 1).number_format = PCT
    c = ws.cell(f, 2, f"={ESC['mp']['per']}*(1+A{f})+{ESC['mod']['per']}+"
                      f"{ESC['tr']['per']}+{ESC['otros']['per']}")
    c.font = FORMULA; c.number_format = EUR2
    c = ws.cell(f, 3, f"={R['precio_persan']}-B{f}"); c.font = FORMULA; c.number_format = EUR2
    c = ws.cell(f, 4, f"=C{f}*{R['vol_persan']}-{R['campa_anio']}-{R['leasing_anio']}"
                      f"+{COSTE_CIRC}")
    c.font = TOTAL; c.number_format = EUR
    f += 1
ws.cell(f, 1, "El contrato fija el precio de venta durante 5 años pero no fija el precio de "
              "la madera. Todo el riesgo de inflación lo asume Alcopalet.").font = NOTA
f += 2

f = seccion(ws, f, "F. ESCENARIO RENEGOCIADO (hipótesis del analista)", 6)
fila_ren = f
for etiqueta, formula, fmt, fuente, nota in [
    ("Precio renegociado (€/ud)", f"={R['precio_reneg']}", EUR2, ENLACE,
     "Editable en la hoja Datos. Comparar siempre con el precio de equilibrio calculado arriba."),
    ("Plazo de cobro renegociado (días)", f"={R['dias_cobro_reneg']}", NUM, ENLACE, None),
    ("Margen de contribución unitario (€)", f"=B{f}-{ESC['coste']['per']}", EUR2, TOTAL, None),
    ("Margen de contribución total (€)", f"=B{f+2}*{R['vol_persan']}", EUR, FORMULA, None),
    ("(−) Campa y leasing", f"=-{R['campa_anio']}-{R['leasing_anio']}", EUR, FORMULA, None),
    ("(−) Coste del circulante",
     f"=-(B{f}*{R['vol_persan']}*B{f+1}/365+{R['reten']}*{ESC['coste']['per']})*{R['interes']}",
     EUR, FORMULA, None),
    ("RESULTADO ANUAL RENEGOCIADO", f"=B{f+3}+B{f+4}+B{f+5}", EUR, TOTAL,
     "Comparar con el impacto del contrato tal como está hoy."),
    ("Mejora frente al contrato actual", f"=B{f+6}-{IMPACTO_PERSAN}", EUR, TOTAL, None),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, formula); c.font = fuente; c.number_format = fmt
    if nota:
        ws.cell(f, 6, nota).font = NOTA
    f += 1
RES_RENEG = f"Persan!$B${fila_ren+6}"
for col in range(1, 7):
    ws.cell(fila_ren + 6, col).fill = F_OK

# ==========================================================================
# HOJA: ROBOT
# ==========================================================================
ws = wb.create_sheet("Robot")
titulo_hoja(ws, "El robot como decisión independiente del contrato",
            "El caso presenta máquina y contrato como una sola decisión. Separarlos es la "
            "clave del análisis: la tecnología es buena aunque el contrato no lo sea.")
anchos(ws, {"A": 50, "B": 18, "C": 16, "D": 16, "E": 12, "F": 54})

f = 4
f = seccion(ws, f, "A. COMPARATIVA TÉCNICA DE LOS TRES SISTEMAS DE FABRICACIÓN", 6)
f = cabecera(ws, f, ["Indicador", "Manual", "Línea 2019", "Robot nuevo", "", "Lectura"])
fila_comp = f
for etiqueta, fm, fl, fr, fmt, nota in [
    ("Producción efectiva (palets/día)", f"={R['man_dia']}",
     f"={R['linea_hora']}*{R['linea_rend']}*{R['horas']}",
     f"={R['rob_dia']}*{R['rob_rend']}", NUM, None),
    ("Operarios necesarios", f"={R['man_oper']}", f"={R['linea_oper']}",
     f"={R['rob_oper']}", NUM, None),
    ("Productividad (palets/operario/día)", f"=B{f}/B{f+1}", f"=C{f}/C{f+1}",
     f"=D{f}/D{f+1}", NUM2,
     "El robot multiplica por 5,8 la productividad de la fabricación manual."),
    ("Cambio de formato (horas)", "", f"={R['linea_cambio_max']}",
     f"={R['rob_cambio']}/60", NUM2,
     "6 horas de parada frente a menos de media hora: el robot permite series cortas "
     "y atender pedidos a medida, que es donde Alcopalet gana margen."),
    ("Coste de mano de obra por palet (€)",
     f"={R['coste_mo']}*B{f+1}/(B{f}*{R['dias_anio']})",
     f"={R['coste_mo']}*C{f+1}/(C{f}*{R['dias_anio']})",
     f"={R['coste_mo']}*D{f+1}/(D{f}*{R['dias_anio']})", EUR2,
     "Coste a plena capacidad de cada sistema. No coincide con el Anexo 3 (2,18 € y 0,42 €) "
     "porque el anexo promedia manual y línea automática, y porque atribuye al robot 90.000 "
     "unidades cuando su capacidad real es de 87.120."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    for col, formula in ((2, fm), (3, fl), (4, fr)):
        if formula != "":
            c = ws.cell(f, col, formula); c.font = FORMULA; c.number_format = fmt
    if nota:
        ws.cell(f, 6, nota).font = NOTA
    f += 1
f += 1

f = seccion(ws, f, "B. AHORRO TEÓRICO DE MANO DE OBRA", 6)
fila_ah = f
for etiqueta, formula, fmt, fuente, nota in [
    ("Capacidad anual del robot (palets)", f"={ROB_ANUAL}", NUM, ENLACE, None),
    ("Operarios manuales equivalentes",
     f"=B{f}/({R['man_prod']}*{R['dias_anio']})", NUM2, FORMULA,
     "Los que harían falta para fabricar lo mismo a mano, a 62,5 palets/operario/día."),
    ("Operarios que exige el robot", f"={R['rob_oper']}", NUM, ENLACE, None),
    ("Operarios liberados", f"=B{f+1}-B{f+2}", NUM2, TOTAL, None),
    ("AHORRO TEÓRICO BRUTO (€/año)", f"=B{f+3}*{R['coste_mo']}", EUR, TOTAL,
     "Solo se materializa si esas personas pasan a producir algo vendible."),
    ("Cuota anual del leasing, solo la máquina (€)",
     f"=-{R['rob_inv']}*({R['interes']}/12)/(1-(1+{R['interes']}/12)^(-{R['leasing_anios']}*12))*12",
     EUR, FORMULA, "40.000 € al 7,5% a 5 años. Anualidad calculada, no estimada."),
    ("AHORRO TEÓRICO NETO (€/año)", f"=B{f+4}+B{f+5}", EUR, TOTAL, None),
    ("Plazo de recuperación de la inversión (meses)",
     f"={R['rob_inv']}/B{f+4}*12", NUM2, TOTAL,
     "Menos de tres meses. Como inversión aislada es excelente."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, formula); c.font = fuente; c.number_format = fmt
    if nota:
        ws.cell(f, 6, nota).font = NOTA
    f += 1
AHORRO_BRUTO = f"Robot!$B${fila_ah+4}"
f += 1

f = seccion(ws, f, "C. LA RESTRICCIÓN QUE LO CAMBIA TODO: «NO QUIERO DESPIDOS»", 6)
ws.cell(f, 1, 'Esteban afirma expresamente que no quiere despidos y que el turno único es un '
              'valor diferencial para retener a la plantilla. Un ahorro de mano de obra que no '
              'se traduce en menos nóminas ni en más producción vendida NO es un ahorro real.').font = NOTA
f += 2
for etiqueta, formula, fmt, fuente, nota in [
    ("% de mano de obra liberada que se monetiza", f"={R['monetiz']}", PCT, ENLACE,
     "HIPÓTESIS CLAVE, editable en la hoja Datos. 0% = nadie se reasigna ni causa baja."),
    ("Ahorro REAL de mano de obra (€/año)", f"={AHORRO_BRUTO}*{R['monetiz']}", EUR, TOTAL, None),
    ("Coste del leasing de la máquina (€/año)", f"=B{fila_ah+5}", EUR, ENLACE, None),
    ("RESULTADO REAL DEL ROBOT AISLADO (€/año)", f"=B{f+1}+B{f+2}", EUR, TOTAL,
     "Con 0% de monetización el robot solo aporta coste. Su valor depende por completo "
     "de tener dónde colocar a las personas o el producto."),
    ("Vías para monetizar la mano de obra liberada", "", None, None,
     "1) Reasignar a la línea de palet usado, que necesita +2,6 operarios y solo le falta "
     "espacio.  2) Crecer en palet nuevo a precio de mercado.  3) No reponer bajas naturales."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    if formula != "":
        c = ws.cell(f, 2, formula); c.font = fuente; c.number_format = fmt
    if nota:
        ws.cell(f, 6, nota).font = NOTA
    f += 1
f += 1

f = seccion(ws, f, "D. COSTE DE OPORTUNIDAD: ¿A QUÉ DEDICAR LA CAPACIDAD DEL ROBOT?", 6)
f = cabecera(ws, f, ["Destino de la producción del robot", "Precio (€/ud)",
                     "Coste (€/ud)", "Margen (€/ud)", "Margen anual (€)", "Comentario"])
fila_op = f
for etiqueta, precio, nota in [
    ("Fabricar para Persán a 10 €", f"={R['precio_persan']}",
     "Precio del contrato tal como está negociado."),
    ("Fabricar para el mercado actual a 15,16 €", f"={ESC['ing']['man']}",
     "Precio medio real del palet nuevo en 2025 (Tetra Pak, Don Simón, Green Box)."),
    ("Fabricar para Persán renegociado", f"={R['precio_reneg']}",
     "Hipótesis de negociación editable en la hoja Datos."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    c = ws.cell(f, 2, precio); c.font = ENLACE; c.number_format = EUR2
    c = ws.cell(f, 3, f"={ESC['mp']['per']}+{ESC['mod']['per']}+{ESC['tr']['per']}+"
                      f"{ESC['otros']['per']}")
    c.font = FORMULA; c.number_format = EUR2
    c = ws.cell(f, 4, f"=B{f}-C{f}"); c.font = TOTAL; c.number_format = EUR2
    c = ws.cell(f, 5, f"=D{f}*{ROB_ANUAL}"); c.font = TOTAL; c.number_format = EUR
    ws.cell(f, 6, nota).font = NOTA
    f += 1
ws.cell(f, 1, "COSTE DE OPORTUNIDAD (mercado vs Persán)").font = DESTACADO
c = ws.cell(f, 2, f"=E{fila_op+1}-E{fila_op}"); c.font = DESTACADO; c.number_format = EUR
for col in range(1, 7):
    ws.cell(f, col).fill = F_ALERTA
ws.cell(f, 6, "Dedicar el robot a Persán en vez de al mercado ordinario cuesta más de "
              "440.000 €/año de margen. La pregunta correcta no es si comprar la máquina, "
              "sino a qué dedicarla. Requiere demanda: hoy la fábrica está al 81%.").font = NOTA

# ==========================================================================
# HOJA: ESCENARIOS
# ==========================================================================
ws = wb.create_sheet("Escenarios")
titulo_hoja(ws, "Comparación de opciones y proyección del resultado",
            "Cada columna es una de las opciones que el caso pone sobre la mesa. "
            "Todos los importes se recalculan desde las hojas anteriores.")
anchos(ws, {"A": 44, "B": 15, "C": 15, "D": 15, "E": 15, "F": 15, "G": 44})

f = 4
f = cabecera(ws, f, ["Concepto", "E0 Statu quo", "E1 Persán + robot",
                     "E2 Dividendo sin invertir", "E3 Robot sin Persán",
                     "E4 Robot + Persán renegociado", "Comentario"])
fila_e = f

filas_esc = {}
for clave, etiqueta, formulas, fmt, negrita, nota in [
    ("inv", "Inversión comprometida (€)",
     ["0", f"={R['inv_total']}", "0", f"={R['rob_inv']}", f"={R['inv_total']}"],
     EUR, False, "E1 y E4 incluyen el tráiler de 150.000 €."),
    ("vol", "Volumen adicional de palet nuevo (ud)",
     ["0", f"={R['vol_persan']}", "0", "0", f"={R['vol_persan']}"],
     NUM, False, "E3 no añade volumen: sustituye fabricación manual por robot."),
    ("mc", "Margen de contribución adicional (€)",
     ["0", f"={MC_PERSAN}", "0", "0",
      f"=({R['precio_reneg']}-{ESC['coste']['per']})*{R['vol_persan']}"],
     EUR, False, None),
    ("ahorro", "Ahorro real de mano de obra (€)",
     ["0", "0", "0", f"={AHORRO_BRUTO}*{R['monetiz']}", "0"],
     EUR, False, "En E1 y E4 el volumen es nuevo: no libera a nadie, exige contratar 1 operario."),
    ("campa", "Alquiler de la campa (€)",
     ["0", f"=-{R['campa_anio']}", "0", "0", f"=-{R['campa_anio']}"],
     EUR, False, None),
    ("leasing", "Leasing (€)",
     ["0", f"=-{R['leasing_anio']}", "0",
      f"=-{R['rob_inv']}*({R['interes']}/12)/(1-(1+{R['interes']}/12)^(-{R['leasing_anios']}*12))*12",
      f"=-{R['leasing_anio']}"],
     EUR, False, "En E3 solo se financia la máquina: no hace falta tráiler."),
    ("circ", "Coste financiero del circulante (€)",
     ["0", f"={COSTE_CIRC}", "0", "0",
      f"=-({R['precio_reneg']}*{R['vol_persan']}*{R['dias_cobro_reneg']}/365"
      f"+{R['reten']}*{ESC['coste']['per']})*{R['interes']}"],
     EUR, False, "E4 supone además reducir el cobro de 180 a 90 días."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    for i, formula in enumerate(formulas):
        c = ws.cell(f, 2 + i, formula); c.font = FORMULA; c.number_format = fmt
    if nota:
        ws.cell(f, 7, nota).font = NOTA
    filas_esc[clave] = f
    f += 1

fila_impacto = f
ws.cell(f, 1, "IMPACTO ANUAL SOBRE EL RESULTADO (€)").font = TOTAL
for col in range(2, 7):
    L = get_column_letter(col)
    c = ws.cell(f, col, f"=SUM({L}{filas_esc['mc']}:{L}{filas_esc['circ']})")
    c.font = TOTAL; c.number_format = EUR
for col in range(1, 8):
    ws.cell(f, col).fill = F_TOTAL
f += 1

fila_rn = f
ws.cell(f, 1, "RESULTADO NETO PROYECTADO (€)").font = TOTAL
for col in range(2, 7):
    L = get_column_letter(col)
    c = ws.cell(f, col, f"={PYG['rn']}+{L}{fila_impacto}")
    c.font = TOTAL; c.number_format = EUR
for col in range(1, 8):
    ws.cell(f, col).fill = F_TOTAL
ws.cell(f, 7, "Partiendo del resultado neto de 2025 (279.364 €).").font = NOTA
f += 1

ws.cell(f, 1, "Variación sobre el resultado actual").font = ETIQUETA
for col in range(2, 7):
    L = get_column_letter(col)
    c = ws.cell(f, col, f"={L}{fila_impacto}/{PYG['rn']}")
    c.font = FORMULA; c.number_format = PCT
f += 2

f = seccion(ws, f, "IMPACTO EN CAJA DEL PRIMER AÑO", 7)
fila_caja = f
for etiqueta, formulas, nota in [
    ("Impacto en resultado (€)",
     [f"={get_column_letter(c)}{fila_impacto}" for c in range(2, 7)], None),
    ("(−) Circulante inmovilizado por el cobro aplazado (€)",
     ["0", f"=-({R['vol_persan']}*{R['precio_persan']}*{R['dias_cobro']}/365"
      f"+{R['reten']}*{ESC['coste']['per']})", "0", "0",
      f"=-({R['precio_reneg']}*{R['vol_persan']}*{R['dias_cobro_reneg']}/365"
      f"+{R['reten']}*{ESC['coste']['per']})"],
     "Solo el primer año: después el circulante se mantiene, no se repite."),
    ("(−) Dividendo extraordinario (€)",
     ["0", "0", f"=-{R['dividendo']}", "0", "0"],
     "El caso no cuantifica la casa. Cifra de trabajo editable en la hoja Datos."),
]:
    ws.cell(f, 1, etiqueta).font = ETIQUETA
    for i, formula in enumerate(formulas):
        c = ws.cell(f, 2 + i, formula); c.font = FORMULA; c.number_format = EUR
    if nota:
        ws.cell(f, 7, nota).font = NOTA
    f += 1

ws.cell(f, 1, "NECESIDAD DE CAJA DEL PRIMER AÑO (€)").font = TOTAL
for col in range(2, 7):
    L = get_column_letter(col)
    c = ws.cell(f, col, f"=SUM({L}{fila_caja}:{L}{f-1})")
    c.font = TOTAL; c.number_format = EUR
for col in range(1, 8):
    ws.cell(f, col).fill = F_ALERTA
ws.cell(f, 7, "Contrastar con la caja operativa real de 2025, muy inferior al beneficio "
              "contable (ver hoja PyG2025).").font = NOTA
f += 1
ws.cell(f, 1, "Caja operativa generada en 2025 (referencia)").font = ETIQUETA
c = ws.cell(f, 2, f"={CAJA_OP}"); c.font = ENLACE; c.number_format = EUR
f += 2

f = seccion(ws, f, "VALORACIÓN CUALITATIVA FRENTE A LOS CRITERIOS DE DECISIÓN", 7)
f = cabecera(ws, f, ["Criterio", "E0", "E1", "E2", "E3", "E4", "Observación"])
for criterio, vals, nota in [
    ("Rentabilidad del contrato", ["n/a", "NO", "n/a", "n/a", "SÍ"],
     "Solo E4 supera el precio de equilibrio de 12,17 €/ud."),
    ("Sostenibilidad de la caja", ["SÍ", "NO", "DUDOSA", "SÍ", "AJUSTADA"],
     "E1 exige casi 600.000 € de caja el primer año."),
    ("Respeta «sin despidos»", ["SÍ", "SÍ", "SÍ", "SÍ", "SÍ"],
     "Ninguna opción obliga a despedir: la capacidad extra se necesita igualmente."),
    ("Respeta el turno único", ["SÍ", "SÍ", "SÍ", "SÍ", "SÍ"], None),
    ("Permite el plan de crecimiento a 2030", ["NO", "SÍ", "NO", "PARCIAL", "SÍ"],
     "E0 y E2 congelan la empresa en los 3.000 m² actuales."),
    ("Da seguridad económica a los fundadores", ["PARCIAL", "NO", "SÍ", "PARCIAL", "PARCIAL"],
     "E1 es la que menos margen deja para retribuir a los socios."),
    ("Reduce la dependencia de mano de obra", ["NO", "SÍ", "NO", "SÍ", "SÍ"],
     "Recurso «cada vez más difícil de gestionar y contratar» según el propio caso."),
    ("Riesgo de concentración de cliente", ["BAJO", "ALTO", "BAJO", "BAJO", "MEDIO"],
     "Persán supondría el 20% de la facturación con precio fijo a 5 años."),
]:
    ws.cell(f, 1, criterio).font = ETIQUETA
    for i, v in enumerate(vals):
        c = ws.cell(f, 2 + i, v)
        c.font = FORMULA
        c.alignment = Alignment(horizontal="center")
        if v == "NO":
            c.fill = F_ALERTA
        elif v == "SÍ":
            c.fill = F_OK
    if nota:
        ws.cell(f, 7, nota).font = NOTA
    f += 1

# ==========================================================================
# HOJA: SENSIBILIDAD
# ==========================================================================
ws = wb.create_sheet("Sensibilidad")
titulo_hoja(ws, "Análisis de sensibilidad del contrato Persán",
            "Impacto anual total (€) según el precio negociado y la evolución del precio de "
            "la madera. Incluye campa, leasing y coste del circulante.")
anchos(ws, {"A": 26})
for col in range(2, 10):
    ws.column_dimensions[get_column_letter(col)].width = 15

f = 4
ws.cell(f, 1, "Precio de venta negociado (€/ud) →").font = TOTAL
ws.cell(f, 1).alignment = Alignment(horizontal="right")
precios = [10.00, 10.50, 11.00, 11.50, 12.00, 12.50, 13.00, 13.50]
for i, p in enumerate(precios):
    c = ws.cell(f, 2 + i, p)
    c.font = CABECERA; c.fill = F_CABECERA; c.number_format = EUR2
    c.alignment = Alignment(horizontal="center")
fila_precios = f
f += 1

ws.cell(f, 1, "↓ Subida del precio de la madera").font = TOTAL
f += 1
fila_tabla = f
inflaciones = [-0.05, 0.00, 0.05, 0.10, 0.15, 0.20]
for infl in inflaciones:
    c = ws.cell(f, 1, infl)
    c.font = CABECERA; c.fill = F_CABECERA; c.number_format = PCT
    c.alignment = Alignment(horizontal="center")
    for i, _ in enumerate(precios):
        L = get_column_letter(2 + i)
        formula = (
            f"={R['vol_persan']}*({L}${fila_precios}"
            f"-({ESC['mp']['per']}*(1+$A{f})+{ESC['mod']['per']}+{ESC['tr']['per']}"
            f"+{ESC['otros']['per']}))"
            f"-{R['campa_anio']}-{R['leasing_anio']}"
            f"-({L}${fila_precios}*{R['vol_persan']}*{R['dias_cobro']}/365"
            f"+{R['reten']}*{ESC['coste']['per']})*{R['interes']}"
        )
        c = ws.cell(f, 2 + i, formula)
        c.font = FORMULA; c.number_format = EUR
        c.border = BORDE
    f += 1

f += 1
ws.cell(f, 1, "Lectura: la zona en negativo es la combinación en la que el contrato destruye "
              "valor. Con el precio actual de 10 € el contrato pierde dinero incluso si la "
              "madera BAJA un 5%.").font = NOTA
f += 1
ws.cell(f, 1, "El punto de equilibrio con madera estable está en torno a 12,17 €/ud: por "
              "debajo de ese precio, ningún volumen ni ninguna máquina arreglan el contrato.").font = NOTA
f += 3

f = seccion(ws, f, "SENSIBILIDAD AL PLAZO DE COBRO (precio fijo de 10 €/ud)", 6)
f = cabecera(ws, f, ["Plazo de cobro (días)", "Circulante inmovilizado (€)",
                     "Coste financiero anual (€)", "Impacto total (€)", "", "Lectura"])
for dias in (30, 60, 90, 120, 180):
    c = ws.cell(f, 1, dias); c.font = ENTRADA; c.number_format = NUM
    c.alignment = Alignment(horizontal="center")
    c = ws.cell(f, 2, f"={R['vol_persan']}*{R['precio_persan']}*A{f}/365"
                      f"+{R['reten']}*{ESC['coste']['per']}")
    c.font = FORMULA; c.number_format = EUR
    c = ws.cell(f, 3, f"=-B{f}*{R['interes']}"); c.font = FORMULA; c.number_format = EUR
    c = ws.cell(f, 4, f"={MC_PERSAN}-{R['campa_anio']}-{R['leasing_anio']}+C{f}")
    c.font = TOTAL; c.number_format = EUR
    f += 1
ws.cell(f, 1, "Negociar el plazo de cobro ayuda, pero no salva el contrato: el problema de "
              "fondo es el precio, no el aplazamiento.").font = NOTA

# ==========================================================================
# HOJA: INSTRUCCIONES (se coloca la primera)
# ==========================================================================
ws = wb.create_sheet("Guia", 0)
titulo_hoja(ws, "ALCOPALET (DTI-1399) — Modelo de análisis",
            "Trabajo final del Programa LYDES 2026 · Instituto Internacional San Telmo")
anchos(ws, {"A": 24, "B": 96})

f = 4
f = seccion(ws, f, "CÓMO USAR ESTE LIBRO", 2)
for etiqueta, texto in [
    ("Regla general", "Ninguna cifra de resultado está escrita a mano: todas son fórmulas. "
                      "Al cambiar una hipótesis, el libro entero se recalcula."),
    ("Dónde tocar", "Solo en la hoja «Datos». El bloque H reúne las hipótesis propias, "
                    "que son las que hay que declarar y defender ante el jurado."),
    ("Azul", "Dato tomado literalmente del enunciado o de sus anexos."),
    ("Negro", "Resultado calculado en la propia hoja."),
    ("Verde", "Enlace a otra hoja del libro."),
    ("Relleno amarillo", "Hipótesis del analista, NO dada por el caso. Hay que justificarla."),
]:
    ws.cell(f, 1, etiqueta).font = TOTAL
    ws.cell(f, 2, texto).font = ETIQUETA
    ws.cell(f, 2).alignment = Alignment(wrap_text=True, vertical="top")
    if etiqueta == "Relleno amarillo":
        ws.cell(f, 1).fill = F_HIPOTESIS
    ws.row_dimensions[f].height = 28
    f += 1
f += 1

f = seccion(ws, f, "CONTENIDO DE LAS HOJAS", 2)
for hoja, texto in [
    ("Datos", "Todos los datos del enunciado y los tres anexos, más las hipótesis propias."),
    ("PyG2025", "Anexo 1 recalculado, con contraste contra el valor publicado y conversión "
                "del EBITDA en caja."),
    ("Lineas", "Anexo 2 y comprobación de que las dos líneas suman el consolidado."),
    ("Escandallo", "Anexo 3 con el coste y el margen recalculados, y reconciliación completa "
                   "contra la cuenta de resultados."),
    ("Capacidad", "Capacidad anual de cada sistema, grado de utilización y encaje de Persán "
                  "en la fábrica. El caso no da estos datos hechos."),
    ("Persan", "Cuenta de resultados incremental del contrato, circulante, precio de "
               "equilibrio, riesgo del precio de la madera y escenario renegociado."),
    ("Robot", "La máquina analizada por separado del contrato: comparativa técnica, ahorro "
              "de mano de obra, efecto de la restricción «sin despidos» y coste de oportunidad."),
    ("Escenarios", "Las cinco opciones comparadas en resultado, caja y criterios cualitativos."),
    ("Sensibilidad", "Tabla de doble entrada precio × precio de la madera, y sensibilidad al "
                     "plazo de cobro."),
]:
    ws.cell(f, 1, hoja).font = TOTAL
    ws.cell(f, 2, texto).font = ETIQUETA
    ws.cell(f, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[f].height = 26
    f += 1
f += 1

f = seccion(ws, f, "LOS CINCO NÚMEROS QUE SOSTIENEN EL ANÁLISIS", 2)
for etiqueta, formula, fmt, texto in [
    ("Margen unitario Persán", f"={ESC['margen']['per']}", EUR2,
     "Negativo. Está en el propio Anexo 3 del caso: el contrato no cubre ni los costes variables."),
    ("Impacto anual del contrato", f"={IMPACTO_PERSAN}", EUR,
     "Con campa, leasing y coste del circulante incluidos."),
    ("Precio de equilibrio", f"={PRECIO_EQ}", EUR2,
     "Por debajo de este precio el contrato destruye valor, con o sin máquina."),
    ("Capacidad anual actual (palets)", f"={CAP_ACTUAL}", NUM,
     "Frente a 174.504 palets producidos en 2025: la fábrica está al 81%, no llena."),
    ("Caja operativa real de 2025", f"={CAJA_OP}", EUR,
     "Frente a 279.364 € de beneficio contable. El beneficio no es caja repartible."),
]:
    ws.cell(f, 1, etiqueta).font = TOTAL
    c = ws.cell(f, 2, formula)
    c.font = Font(name=FUENTE, size=12, bold=True, color="C00000")
    c.number_format = fmt
    ws.cell(f, 3, texto).font = NOTA
    f += 1
f += 1

f = seccion(ws, f, "ADVERTENCIAS DE HONESTIDAD ANALÍTICA", 2)
for texto in [
    "El coste del circulante es una aproximación de bolsillo (saldo medio × tipo de interés), "
    "no un estado de tesorería completo. Sirve como orden de magnitud.",
    "El caso no dice si existe demanda suficiente para llenar la capacidad del robot a precio "
    "de mercado. Hoy la fábrica está al 81%, así que la demanda es hoy la restricción real.",
    "El importe de la casa que quieren comprar los fundadores no aparece en el caso. La cifra "
    "usada es una hipótesis de trabajo, señalada en amarillo.",
    "La mano de obra directa del escandallo (646.000 €) no cuadra con la partida de personal "
    "de producción del Anexo 1 (481.660 €). El escandallo sí cuadra con la plantilla real "
    "(17 operarios × 38.000 €), por lo que se ha tomado como base.",
    "El anexo publica un coste manual de 12,36 €; la suma exacta de sus componentes da 12,37 €. "
    "Diferencia de redondeo, sin efecto en las conclusiones.",
]:
    ws.cell(f, 1, "•").font = ETIQUETA
    ws.cell(f, 2, texto).font = ETIQUETA
    ws.cell(f, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[f].height = 30
    f += 1

wb.save("/home/user/DANISANTELMO/Alcopalet_Modelo_Financiero.xlsx")
print("Modelo guardado.")
