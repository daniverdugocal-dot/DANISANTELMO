"""
Alcopalet (DTI-1399) — Análisis de la cuenta de resultados 2025.
Una sola hoja. Azul = dato del Anexo 1. Negro = fórmula.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

F = "Arial"
TIT = Font(name=F, size=15, bold=True, color="1F3864")
SUB = Font(name=F, size=9, italic=True, color="595959")
SEC = Font(name=F, size=11, bold=True, color="FFFFFF")
ETI = Font(name=F, size=10)
IN_ = Font(name=F, size=10, color="0000FF")
FOR = Font(name=F, size=10)
TOT = Font(name=F, size=10, bold=True)
BIG = Font(name=F, size=12, bold=True, color="C00000")
POR = Font(name=F, size=9, italic=True, color="595959")

FS = PatternFill("solid", fgColor="1F3864")
FT = PatternFill("solid", fgColor="D9E2F3")
FA = PatternFill("solid", fgColor="FCE4E4")
FV = PatternFill("solid", fgColor="E2EFDA")

fina = Side(style="thin", color="BFBFBF")
BOR = Border(left=fina, right=fina, top=fina, bottom=fina)

EUR = '#,##0 "€";(#,##0) "€";"-"'
EU2 = '#,##0.00 "€";(#,##0.00) "€";"-"'
PCT = '0.0%;(0.0%);"-"'
NUM = '#,##0;(#,##0);"-"'
MUL = '0.00"x"'

wb = Workbook()
ws = wb.active
ws.title = "PyG 2025"
ws.sheet_view.showGridLines = False
for col, w in {"A": 44, "B": 16, "C": 11, "D": 14, "E": 72}.items():
    ws.column_dimensions[col].width = w

ws["A1"] = "ALCOPALET — Cuenta de resultados 2025 y análisis"
ws["A1"].font = TIT
ws["A2"] = ("Caso DTI-1399 · Programa Lydes 2026 · Azul = dato del Anexo 1 · "
            "Negro = fórmula calculada · Columna E = para qué sirve cada cálculo")
ws["A2"].font = SUB

f = 4


def sec(texto):
    global f
    ws.cell(f, 1, texto).font = SEC
    for c in range(1, 6):
        ws.cell(f, c).fill = FS
    f += 1


def fila(etiqueta, valor, fmt=EUR, fuente=FOR, pct=None, porque=None,
         relleno=None, ffont=None):
    global f
    ws.cell(f, 1, etiqueta).font = ffont or (TOT if relleno else ETI)
    if valor != "":                      # las filas «¿PARA QUÉ?» no llevan importe
        c = ws.cell(f, 2, valor)
        c.font = ffont or fuente
        c.number_format = fmt
    if pct:
        p = ws.cell(f, 3, pct)
        p.font = FOR
        p.number_format = PCT
    if porque:
        n = ws.cell(f, 5, porque)
        n.font = POR
        n.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[f].height = max(15, 12 * (len(porque) // 78 + 1))
    if relleno:
        for c2 in range(1, 6):
            ws.cell(f, c2).fill = relleno
    f += 1
    return f - 1


# ---------------------------------------------------------------- Anexo 1
sec("LA CUENTA DE RESULTADOS DEL ANEXO 1")
for h, t in ((1, "Concepto"), (2, "Importe 2025"), (3, "% s/ventas"),
             (4, "Dif. anexo"), (5, "Nota")):
    c = ws.cell(f, h, t)
    c.font = Font(name=F, size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="4472C4")
    c.alignment = Alignment(horizontal="center")
    c.border = BOR
f += 1

r_ven = f
fila("VENTAS NETAS", 4585267, EUR, IN_, f"=B{f}/$B${f}", relleno=FT)
r_vex = fila("Variación de existencias", 376062, EUR, IN_, f"=B{f}/$B${r_ven}",
             porque="Aumento del almacén. Es ingreso contable, pero NO es caja: "
                    "se compró más madera de la que se consumió.")
r_apr = fila("Aprovisionamientos", -3068506, EUR, IN_, f"=B{f}/$B${r_ven}",
             porque="COMPRAS del ejercicio, no consumo. Ver cálculo 1.")
r_mb = f
fila("MARGEN BRUTO", f"=SUM(B{r_ven}:B{r_apr})", EUR, TOT, f"=B{f}/$B${r_ven}", relleno=FT)
ws.cell(r_mb, 4, f"=B{r_mb}-1892822").number_format = EUR
ws.cell(r_mb, 4).font = FOR
r_pp = fila("Gastos de personal — Producción", -481660, EUR, IN_, f"=B{f}/$B${r_ven}")
r_tr = fila("Gastos de transporte", -268218, EUR, IN_, f"=B{f}/$B${r_ven}",
            porque="Los 2 camiones actuales, YA al 100% de capacidad. Ver cálculo 6.")
r_pe = fila("Gastos de personal — Estructura", -311292, EUR, IN_, f"=B{f}/$B${r_ven}")
r_oc = fila("Otros costes de explotación", -378979, EUR, IN_, f"=B{f}/$B${r_ven}",
            porque="El 60% corresponde a palet nuevo. Varían con el volumen.")
r_ebda = f
fila("EBITDA", f"=B{r_mb}+SUM(B{r_pp}:B{r_oc})", EUR, TOT, f"=B{f}/$B${r_ven}", relleno=FT)
ws.cell(r_ebda, 4, f"=B{r_ebda}-452674").number_format = EUR
ws.cell(r_ebda, 4).font = FOR
r_am = fila("Amortización", -36143, EUR, IN_, f"=B{f}/$B${r_ven}")
r_ebit = f
fila("EBIT", f"=B{r_ebda}+B{r_am}", EUR, TOT, f"=B{f}/$B${r_ven}", relleno=FT)
r_gf = fila("Gastos financieros", -44046, EUR, IN_, f"=B{f}/$B${r_ven}",
            porque="Al 7,5% del banco, implican unos 587.000 € de deuda. Ver cálculo 5.")
r_bai = f
fila("RESULTADO ANTES DE IMPUESTOS", f"=B{r_ebit}+B{r_gf}", EUR, TOT, f"=B{f}/$B${r_ven}",
     relleno=FT)
r_imp = fila("Impuestos", -93121, EUR, IN_, f"=B{f}/$B${r_ven}")
r_rn = f
fila("RESULTADO NETO", f"=B{r_bai}+B{r_imp}", EUR, TOT, f"=B{f}/$B${r_ven}", relleno=FT)
ws.cell(r_rn, 4, f"=B{r_rn}-279364").number_format = EUR
ws.cell(r_rn, 4).font = FOR
f += 1

# ------------------------------------------------------------- Cálculo 1
sec("CÁLCULO 1 — CONSUMO REAL DE MATERIA PRIMA")
fila("¿PARA QUÉ? Conectar el Anexo 1 con el escandallo del Anexo 3.", "", EUR,
     porque="Los 3.068.506 € son lo que se COMPRÓ. Si usas esa cifra, el Anexo 3 parece "
            "descuadrado y acabas descartándolo. Restando el aumento de existencias, "
            "ambos anexos cuadran y puedes apoyarte en el escandallo con confianza.")
fila("Aprovisionamientos (compras)", f"=-B{r_apr}")
fila("(−) Aumento de existencias", f"=-B{r_vex}")
r_cons = f
fila("CONSUMO REAL DE MATERIA PRIMA", f"=-B{r_apr}-B{r_vex}", EUR, TOT,
     f"=B{f}/B{r_ven}", relleno=FT)
fila("Materia prima según el Anexo 3", 2692732, EUR, IN_,
     porque="8,27 €×174.504 palets nuevos + 4,80 €×260.330 palets usados.")
fila("DESVIACIÓN", f"=B{r_cons}-B{f-1}", EUR, TOT,
     porque="288 € sobre 2,7 millones. Los anexos del caso son coherentes.", relleno=FV)
f += 1

# ------------------------------------------------------------- Cálculo 2
sec("CÁLCULO 2 — SEPARAR COSTES FIJOS DE VARIABLES")
fila("¿PARA QUÉ? Es EL cálculo del caso.", "", EUR,
     porque="Persán es un pedido incremental, y un pedido incremental no se juzga con "
            "costes completos sino con margen de contribución. Sin esta separación no "
            "se puede decir nada del contrato.")
fila("Consumo de materia prima", f"=B{r_cons}", EUR, FOR, f"=B{f}/B{r_ven}")
fila("Transporte", f"=-B{r_tr}", EUR, FOR, f"=B{f}/B{r_ven}")
fila("Otros costes de explotación", f"=-B{r_oc}", EUR, FOR, f"=B{f}/B{r_ven}")
r_cv = f
fila("TOTAL COSTES VARIABLES", f"=SUM(B{f-3}:B{f-1})", EUR, TOT, f"=B{f}/B{r_ven}",
     relleno=FT)
r_mc = f
fila("MARGEN DE CONTRIBUCIÓN", f"=B{r_ven}-B{r_cv}", EUR, TOT, f"=B{f}/B{r_ven}",
     porque="Cada euro vendido deja 27 céntimos para pagar la estructura.", relleno=FT)
fila("Personal de producción", f"=-B{r_pp}", EUR, FOR, f"=B{f}/B{r_ven}",
     porque="Fijo en la práctica: Esteban afirma que no quiere despidos.")
fila("Personal de estructura", f"=-B{r_pe}", EUR, FOR, f"=B{f}/B{r_ven}")
r_cf = f
fila("TOTAL COSTES FIJOS", f"=B{f-2}+B{f-1}", EUR, TOT, f"=B{f}/B{r_ven}", relleno=FT)
fila("EBITDA (comprobación)", f"=B{r_mc}-B{r_cf}", EUR, TOT, f"=B{f}/B{r_ven}")
fila("Diferencia con el EBITDA del Anexo 1", f"=B{f-1}-B{r_ebda}", EUR, TOT,
     porque="Cuadra al euro. Esta estructura se puede defender ante el jurado.",
     relleno=FV)
f += 1

# ------------------------------------------------------------- Cálculo 3
sec("CÁLCULO 3 — PUNTO MUERTO Y APALANCAMIENTO OPERATIVO")
fila("¿PARA QUÉ? Medir el colchón y por qué Persán lo estrecha.", "", EUR,
     porque="Persán empeora las dos variables a la vez: sube los costes fijos en 108.000 € "
            "(campa + leasing) Y aporta margen de contribución negativo. Eleva el punto "
            "muerto mientras reduce la capacidad de alcanzarlo.")
r_ratio = f
fila("Ratio de margen de contribución", f"=B{r_mc}/B{r_ven}", PCT, TOT)
fila("Punto muerto (EBITDA = 0)", f"=B{r_cf}/B{r_ratio}", EUR, TOT)
r_pm = f
fila("PUNTO MUERTO (resultado = 0)", f"=(B{r_cf}-B{r_am}-B{r_gf})/B{r_ratio}", EUR, TOT,
     relleno=FT)
fila("Ventas reales 2025", f"=B{r_ven}", EUR, FOR)
fila("MARGEN DE SEGURIDAD", f"=(B{r_ven}-B{r_pm})/B{r_ven}", PCT, BIG,
     porque="La facturación puede caer un 30% antes de entrar en pérdidas. "
            "La empresa está sana HOY.", relleno=FV)
fila("Apalancamiento operativo", f"=B{r_mc}/B{r_ebda}", MUL, TOT,
     porque="Un 1% menos de margen de contribución mueve el EBITDA un 2,75%. "
            "Los errores de precio se amplifican.")
f += 1

# ------------------------------------------------------------- Cálculo 4
sec("CÁLCULO 4 — ¿CUÁNTA CAJA GENERÓ REALMENTE LA EMPRESA?")
fila("¿PARA QUÉ? Responder a los padres y al proyecto a la vez.", "", EUR,
     porque="Hay dos peticiones de dinero sobre la mesa: la casa de los fundadores y los "
            "643.000 € que exige Persán el primer año. La respuesta no está en el "
            "resultado neto, está aquí.")
fila("EBITDA", f"=B{r_ebda}")
fila("(−) Aumento de existencias", f"=-B{r_vex}",
     porque="Salida de caja: madera comprada que sigue en el almacén.")
fila("(−) Gastos financieros", f"=B{r_gf}")
fila("(−) Impuestos", f"=B{r_imp}")
r_caja = f
fila("CAJA OPERATIVA APROXIMADA", f"=SUM(B{f-4}:B{f-1})", EUR, BIG, relleno=FA)
fila("Resultado neto contable", f"=B{r_rn}", EUR, FOR)
fila("BRECHA entre beneficio y caja", f"=B{r_rn}-B{r_caja}", EUR, BIG,
     porque="El beneficio está en el almacén, no en el banco. Este es el argumento "
            "central frente a la petición de dividendo.", relleno=FA)
fila("Aviso: no hay balance en el caso.", "", EUR,
     porque="Se ignoran clientes y proveedores, así que es una aproximación y hay que "
            "declararlo así en el informe. El acopio de madera puede ser puntual. "
            "La dirección del argumento se mantiene.")
f += 1

# ------------------------------------------------------------- Cálculo 5
sec("CÁLCULO 5 — ENDEUDAMIENTO IMPLÍCITO")
fila("¿PARA QUÉ? El caso no da balance, pero se puede reconstruir.", "", EUR,
     porque="Los gastos financieros divididos por el tipo del banco dan la deuda "
            "aproximada. Sirve para saber si la empresa puede asumir más.")
r_tipo = f
fila("Tipo de interés del banco", 0.075, PCT, IN_)
r_deuda = f
fila("Deuda implícita", f"=-B{r_gf}/B{r_tipo}", EUR, TOT)
fila("Deuda / EBITDA — HOY", f"=B{r_deuda}/B{r_ebda}", MUL, TOT,
     porque="Por debajo de 2x. La empresa está cómoda y el banco le presta sin problema.",
     relleno=FV)
fila("Cobertura de intereses (EBITDA / gastos fin.)", f"=B{r_ebda}/-B{r_gf}", MUL, TOT)
fila("Deuda estimada TRAS firmar Persán", f"=B{r_deuda}+190000+454356", EUR, FOR,
     porque="Añade 190.000 € de leasing (robot + tráiler) y 454.356 € de circulante "
            "por el cobro a 180 días.")
fila("EBITDA estimado tras Persán", f"=B{r_ebda}-154800", EUR, FOR,
     porque="Margen de contribución negativo (−46.800) + campa (60.000) + leasing (48.000).")
fila("Deuda / EBITDA — TRAS PERSÁN", f"=B{f-2}/B{f-1}", MUL, BIG,
     porque="Por encima de 4x es la zona donde el banco empieza a poner condiciones. "
            "Y ocurre justo cuando los fundadores quieren retirar capital.", relleno=FA)
f += 1

# ------------------------------------------------------------- Cálculo 6
sec("CÁLCULO 6 — COSTE REAL DEL TRANSPORTE (inferencia propia)")
fila("¿PARA QUÉ? Comprobar si el Anexo 3 es optimista.", "", EUR,
     porque="El contrato exige DOS camiones de retén permanente en nave. Conviene "
            "comprobar si el coste de transporte que el escandallo asigna a Persán es "
            "coherente con lo que cuestan los camiones actuales.")
r_ct = f
fila("Coste total de transporte 2025", f"=-B{r_tr}")
fila("Número de camiones actuales", 2, NUM, IN_)
fila("COSTE POR CAMIÓN Y AÑO", f"=B{r_ct}/B{f-1}", EUR, TOT, relleno=FT)
fila("Transporte asignado a Persán en el Anexo 3", f"=0.53*90000", EUR, FOR,
     porque="0,53 €/palet × 90.000 unidades.")
fila("DIFERENCIA si el 3er camión se dedica a Persán", f"=B{f-2}-B{f-1}", EUR, BIG,
     porque="INFERENCIA PROPIA, no dato del caso: hay que presentarla como tal. Apunta a "
            "que el escandallo del Anexo 3 es optimista y que el contrato es aún peor "
            "de lo que ya parece.", relleno=FA)
f += 2

# ------------------------------------------------------------- Conclusión
sec("LO QUE DICE LA CUENTA DE RESULTADOS")
for txt in [
    "La empresa está sana pero es más frágil de lo que parece: 9,9% de EBITDA y un margen "
    "de seguridad del 30%, pero el 72,8% de la facturación se va en costes variables.",
    "La madera sola es el 58,7% de las ventas. En lo económico, Alcopalet es una "
    "transformadora de madera con margen estrecho sobre una materia prima que no controla.",
    "Con esa estructura el precio de venta no es una variable comercial: es LA variable. "
    "Vender a 10 € un palet cuyo coste variable es 10,52 € no es un descuento agresivo.",
    "El apalancamiento de 2,75x amplifica el error. Persán añade 108.000 € de costes fijos "
    "(un 13,6% más) sin aportar contribución alguna.",
    "Y la caja dice que no hay margen para equivocarse: el debate del comité está mal "
    "planteado. No es «invertir o repartir», es que ahora mismo no hay caja para ninguna "
    "de las dos cosas, porque el beneficio está inmovilizado en existencias.",
]:
    c = ws.cell(f, 1, "•")
    c.font = ETI
    c2 = ws.cell(f, 2, txt)
    c2.font = ETI
    c2.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=f, start_column=2, end_row=f, end_column=5)
    ws.row_dimensions[f].height = 13 * (len(txt) // 100 + 1) + 6
    f += 1

wb.save("/home/user/DANISANTELMO/Alcopalet_PyG_2025.xlsx")
print("Guardado.")
