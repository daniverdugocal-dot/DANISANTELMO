"""
ALCOPALET (DTI-1399) — Análisis completo del caso.

Un libro, siete hojas, todas con la misma estructura:
    TÍTULO DE BLOQUE  ->  para qué sirve (a todo el ancho)  ->  los números

Cada hoja es autónoma: repite los datos del caso que necesita en vez de
depender de referencias a otras hojas. Así se puede leer y auditar sola.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FUENTE = "Arial"

TIT = Font(name=FUENTE, size=16, bold=True, color="1F3864")
SUB = Font(name=FUENTE, size=9, italic=True, color="595959")
SEC = Font(name=FUENTE, size=11, bold=True, color="FFFFFF")
CAB = Font(name=FUENTE, size=10, bold=True, color="FFFFFF")
ETI = Font(name=FUENTE, size=10)
AZUL = Font(name=FUENTE, size=10, color="0000FF")
NEG = Font(name=FUENTE, size=10)
TOT = Font(name=FUENTE, size=10, bold=True)
BIG = Font(name=FUENTE, size=13, bold=True, color="C00000")
VERDE = Font(name=FUENTE, size=13, bold=True, color="006100")
EXP = Font(name=FUENTE, size=9, italic=True, color="404040")
MINI = Font(name=FUENTE, size=9, color="404040")

F_SEC = PatternFill("solid", fgColor="1F3864")
F_CAB = PatternFill("solid", fgColor="4472C4")
F_TOT = PatternFill("solid", fgColor="D9E2F3")
F_EXP = PatternFill("solid", fgColor="F2F2F2")
F_MAL = PatternFill("solid", fgColor="FCE4E4")
F_BIEN = PatternFill("solid", fgColor="E2EFDA")
F_CLAVE = PatternFill("solid", fgColor="FFF2CC")
F_HIP = PatternFill("solid", fgColor="FFFF00")

_f = Side(style="thin", color="BFBFBF")
BOR = Border(left=_f, right=_f, top=_f, bottom=_f)

EUR = '#,##0 "€";(#,##0) "€";"-"'
EU2 = '#,##0.00 "€";(#,##0.00) "€";"-"'
PCT = '0.0%;(0.0%);"-"'
NUM = '#,##0;(#,##0);"-"'
NU2 = '#,##0.00;(#,##0.00);"-"'
MUL = '0.00"x"'

ANCHOS = {"A": 50, "B": 17, "C": 13, "D": 42}
ANCHO_TXT = 122


class Hoja:
    """Envoltorio que lleva la cuenta de filas y aplica el estilo común."""

    def __init__(self, wb, nombre, titulo, subtitulo):
        self.ws = wb.create_sheet(nombre)
        self.nombre = nombre
        self.f = 1
        ws = self.ws
        ws.sheet_view.showGridLines = False
        for col, w in ANCHOS.items():
            ws.column_dimensions[col].width = w
        ws["A1"] = titulo
        ws["A1"].font = TIT
        ws["A2"] = subtitulo
        ws["A2"].font = SUB
        ws.merge_cells("A2:D2")
        ws.row_dimensions[2].height = 12.5 * (len(subtitulo) // ANCHO_TXT + 1) + 4
        self.f = 4
        ws.freeze_panes = "A4"

    # ---------------------------------------------------------------- bloques
    def barra(self, texto, etiqueta=None):
        ws, f = self.ws, self.f
        ws.cell(f, 1, texto).font = SEC
        for c in range(1, 5):
            ws.cell(f, c).fill = F_SEC
        if etiqueta:
            c = ws.cell(f, 4, etiqueta)
            c.font = Font(name=FUENTE, size=9, bold=True, color="FFFFFF")
            c.alignment = Alignment(horizontal="right")
        self.f += 1

    def explicar(self, texto):
        ws, f = self.ws, self.f
        c = ws.cell(f, 1, texto)
        c.font = EXP
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=4)
        for col in range(1, 5):
            ws.cell(f, col).fill = F_EXP
        ws.row_dimensions[f].height = 12.5 * (len(texto) // ANCHO_TXT + 1) + 5
        self.f += 1

    def cabecera(self, cols):
        ws, f = self.ws, self.f
        for i, t in enumerate(cols):
            c = ws.cell(f, 1 + i, t)
            c.font = CAB
            c.fill = F_CAB
            c.alignment = Alignment(horizontal="left" if i == 0 else "center",
                                    wrap_text=True)
            c.border = BOR
        self.f += 1

    def fila(self, etiqueta, valor="", fmt=EUR, fuente=NEG, col3=None, col3fmt=PCT,
             nota=None, relleno=None, grande=None):
        """Devuelve el número de fila escrita."""
        ws, f = self.ws, self.f
        ws.cell(f, 1, etiqueta).font = TOT if (relleno or grande) else ETI
        if valor != "":
            c = ws.cell(f, 2, valor)
            c.font = grande if grande else (TOT if relleno else fuente)
            c.number_format = fmt
        if col3 is not None:
            c = ws.cell(f, 3, col3)
            c.font = TOT if relleno else NEG
            c.number_format = col3fmt
        if nota:
            n = ws.cell(f, 4, nota)
            n.font = MINI
            n.alignment = Alignment(wrap_text=True, vertical="center")
        if relleno:
            for col in range(1, 5):
                ws.cell(f, col).fill = relleno
        self.f += 1
        return f

    def texto(self, txt, vinieta=True):
        ws, f = self.ws, self.f
        c = ws.cell(f, 1, ("• " if vinieta else "") + txt)
        c.font = ETI
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=4)
        ws.row_dimensions[f].height = 12.5 * (len(txt) // ANCHO_TXT + 1) + 6
        self.f += 1

    def hueco(self, n=1):
        self.f += n


wb = Workbook()
wb.remove(wb.active)

# ══════════════════════════════════════════════════════════════════════════
# HOJA 0 — GUÍA
# ══════════════════════════════════════════════════════════════════════════
g = Hoja(wb, "0. Guía", "ALCOPALET — Guía del análisis",
         "Caso DTI-1399 · Programa Lydes 2026 · Instituto Internacional San Telmo. "
         "Empieza por aquí: esta hoja dice qué hay en cada una de las demás y cómo leerlas.")

g.barra("CÓMO SE LEE ESTE LIBRO")
g.explicar("Todas las hojas tienen la misma estructura: una banda azul con el título del "
           "bloque, debajo una línea gris que explica PARA QUÉ sirve ese cálculo, y después "
           "los números. Se lee de arriba abajo, como un documento. No hace falta saltar "
           "entre hojas.")
g.cabecera(["Código de color", "", "", "Qué significa"])
for etiqueta, nota in [
    ("Texto AZUL", "Dato tomado literalmente del enunciado o de sus anexos. No lo he tocado."),
    ("Texto NEGRO", "Resultado calculado con fórmula. Si cambias un dato azul, esto se recalcula."),
    ("Fondo AMARILLO", "Hipótesis MÍA, no dada por el caso. Es lo que hay que declarar y defender."),
    ("Fondo VERDE", "Comprobación que cuadra, o resultado favorable."),
    ("Fondo ROJO", "Número que revela un problema."),
]:
    g.fila(etiqueta, "", EUR, NEG, None, PCT, nota)
g.hueco()

g.barra("LOS CINCO NÚMEROS QUE DECIDEN EL CASO")
g.explicar("Todo el análisis se reduce a esto. Cada cifra está calculada en la hoja que se "
           "indica, donde puedes ver de dónde sale.")
g.cabecera(["Número", "Valor", "Hoja", "Qué significa"])
g.fila("Margen de contribución del negocio", 0.2717, PCT, AZUL, "1", NUM,
       "De cada euro vendido quedan 27 céntimos tras pagar madera, transporte y variables.",
       F_CLAVE)
g.fila("Margen de contribución de Persán", -0.52, EU2, AZUL, "2", NUM,
       "NEGATIVO. Cada palet vendido a Persán pierde 52 céntimos antes de costes fijos.",
       F_MAL)
g.fila("Utilización actual de la fábrica", 0.810, PCT, AZUL, "3", NUM,
       "La fábrica NO está llena: sobran casi 41.000 palets/año de capacidad.", F_CLAVE)
g.fila("Precio de equilibrio de Persán", 12.17, EU2, AZUL, "4", NUM,
       "Por debajo de este precio el contrato destruye valor, con máquina o sin ella.",
       F_CLAVE)
g.fila("Caja operativa real de 2025", -60555, EUR, AZUL, "1", NUM,
       "Frente a 279.364 € de beneficio contable. No hay caja para repartir ni para invertir.",
       F_MAL)
g.hueco()

g.barra("QUÉ HAY EN CADA HOJA")
g.cabecera(["Hoja", "", "", "Contenido y para qué sirve"])
for hoja, desc in [
    ("1. Cuenta de resultados",
     "El Anexo 1 reordenado por comportamiento (fijo/variable), la conversión del beneficio "
     "en caja y el punto muerto. Responde: ¿cuánto deja cada venta y hay dinero disponible?"),
    ("2. Escandallo",
     "El Anexo 3 recalculado. Aquí aparece el margen negativo de Persán y el peso de la "
     "madera. Responde: ¿cuánto cuesta de verdad un palet?"),
    ("3. Capacidad",
     "Capacidad anual de cada sistema de fabricación. El caso da datos por hora y por "
     "jornada pero nunca el total. Responde: ¿cabe Persán en la fábrica?"),
    ("4. Contrato Persán",
     "Cuenta de resultados del contrato, circulante por el cobro a 180 días, precio de "
     "equilibrio y sensibilidad. Responde: ¿cuánto cuesta firmar?"),
    ("5. El robot",
     "La máquina analizada POR SEPARADO del contrato. Responde: ¿la compro aunque no firme?"),
    ("6. Las opciones",
     "Las cinco alternativas comparadas en resultado, caja y criterios cualitativos. "
     "Responde: ¿qué hago?"),
]:
    g.fila(hoja, "", EUR, NEG, None, PCT, desc)
g.hueco()

g.barra("MIS HIPÓTESIS (esto NO lo dice el caso)")
g.explicar("Son las cuatro cosas que he tenido que suponer. Van marcadas en amarillo allí "
           "donde se usan. Hay que declararlas en el informe: el jurado valora más una "
           "hipótesis explícita que un número sin origen.")
g.cabecera(["Hipótesis", "Valor usado", "", "Por qué y qué pasa si falla"])
g.fila("Precio renegociado objetivo con Persán", 12.50, EU2, NEG, None, PCT,
       "Por encima del equilibrio de 12,17 €. Si Persán no acepta, el contrato no interesa.",
       F_HIP)
g.fila("Plazo de cobro renegociado", 90, NUM, NEG, None, PCT,
       "La mitad de los 180 días pactados. Libera la mitad del circulante.", F_HIP)
g.fila("% de mano de obra liberada que se monetiza", 0.00, PCT, NEG, None, PCT,
       "Esteban NO quiere despidos. Si nadie se reasigna, el ahorro del robot es cero.",
       F_HIP)
g.fila("Importe del dividendo planteado", 200000, EUR, NEG, None, PCT,
       "El caso no dice cuánto cuesta la casa. Cifra de trabajo.", F_HIP)
g.hueco()

g.barra("ADVERTENCIAS DE HONESTIDAD")
for t in [
    "El coste del circulante es una aproximación (saldo medio × tipo de interés), no un "
    "estado de tesorería. El caso no da balance, así que se ignoran clientes y proveedores.",
    "El caso no dice si hay demanda para llenar la capacidad del robot a precio de mercado. "
    "Con la fábrica al 81%, hoy la restricción es la demanda, no la capacidad.",
    "El Anexo 3 publica un coste manual de 12,36 €; la suma exacta de sus componentes da "
    "12,37 €. Diferencia de redondeo, sin efecto en las conclusiones.",
    "La mano de obra del escandallo (646.000 €) no coincide con la partida de personal de "
    "producción del Anexo 1 (481.660 €). El escandallo sí cuadra con la plantilla real "
    "(17 operarios × 38.000 €), por eso se ha tomado como base.",
]:
    g.texto(t)

# ══════════════════════════════════════════════════════════════════════════
# HOJA 1 — CUENTA DE RESULTADOS
# ══════════════════════════════════════════════════════════════════════════
h = Hoja(wb, "1. Cuenta resultados", "1. La cuenta de resultados de 2025",
         "Responde a dos preguntas: ¿cuánto deja realmente cada venta? y ¿hay dinero "
         "disponible para invertir y para repartir?")

h.barra("LA CUENTA TAL COMO LA DA EL CASO (Anexo 1)")
h.explicar("Los subtotales están calculados con fórmula en vez de copiados, para comprobar "
           "que cuadran con lo publicado. Es control de calidad, no análisis.")
h.cabecera(["Concepto", "Importe 2025", "% s/ventas", "Comentario"])

r_ven = h.fila("VENTAS NETAS", 4585267, EUR, AZUL, None, PCT, None, F_TOT)
h.ws.cell(r_ven, 3, f"=B{r_ven}/$B${r_ven}").number_format = PCT
h.ws.cell(r_ven, 3).font = TOT


def pyg(et, val, nota=None):
    return h.fila(et, val, EUR, AZUL, f"=B{h.f}/$B${r_ven}", PCT, nota)


r_vex = pyg("Variación de existencias", 376062, "Ingreso contable, pero NO es caja")
r_apr = pyg("Aprovisionamientos", -3068506, "Son COMPRAS, no consumo")
r_mb = h.fila("MARGEN BRUTO", f"=SUM(B{r_ven}:B{r_apr})", EUR, TOT,
              f"=B{h.f}/$B${r_ven}", PCT, None, F_TOT)
r_pp = pyg("Gastos de personal — Producción", -481660)
r_tr = pyg("Gastos de transporte", -268218, "2 camiones, ya al 100% de capacidad")
r_pe = pyg("Gastos de personal — Estructura", -311292)
r_oc = pyg("Otros costes de explotación", -378979, "60% es palet nuevo; varían con el volumen")
r_ebda = h.fila("EBITDA", f"=B{r_mb}+SUM(B{r_pp}:B{r_oc})", EUR, TOT,
                f"=B{h.f}/$B${r_ven}", PCT, None, F_TOT)
r_am = pyg("Amortización", -36143, "Línea de 2019: 360.000 € a 10 años")
r_ebit = h.fila("EBIT", f"=B{r_ebda}+B{r_am}", EUR, TOT, f"=B{h.f}/$B${r_ven}",
                PCT, None, F_TOT)
r_gf = pyg("Gastos financieros", -44046, "Al 7,5% implican ~587.000 € de deuda")
r_bai = h.fila("RESULTADO ANTES DE IMPUESTOS", f"=B{r_ebit}+B{r_gf}", EUR, TOT,
               f"=B{h.f}/$B${r_ven}", PCT, None, F_TOT)
r_imp = pyg("Impuestos", -93121)
r_rn = h.fila("RESULTADO NETO", f"=B{r_bai}+B{r_imp}", EUR, TOT,
              f"=B{h.f}/$B${r_ven}", PCT, None, F_TOT)
h.hueco()

h.barra("REORDENAR LA CUENTA: COSTES FIJOS FRENTE A VARIABLES", "→ VA AL INFORME")
h.explicar("El caso ordena los gastos por naturaleza (personal, transporte, otros). Para "
           "decidir sobre un pedido NUEVO eso no sirve: hay que reordenarlos por "
           "comportamiento, es decir, cuáles crecen si fabricas más y cuáles se pagan igual. "
           "Lo que queda tras pagar los variables es el margen de contribución, y ese es el "
           "único criterio válido para juzgar el contrato de Persán. Fíjate en que al final "
           "reconstruye el EBITDA exacto del Anexo 1: la reordenación es indiscutible.")
h.cabecera(["Concepto", "Importe", "% s/ventas", "Comentario"])

r_cons = h.fila("Consumo real de materia prima", f"=-B{r_apr}-B{r_vex}", EUR, NEG,
                f"=B{h.f}/B{r_ven}", PCT, "Compras menos lo que quedó en el almacén")
h.fila("Transporte", f"=-B{r_tr}", EUR, NEG, f"=B{h.f}/B{r_ven}", PCT)
h.fila("Otros costes de explotación", f"=-B{r_oc}", EUR, NEG, f"=B{h.f}/B{r_ven}", PCT)
r_cv = h.fila("TOTAL COSTES VARIABLES", f"=SUM(B{h.f-3}:B{h.f-1})", EUR, TOT,
              f"=B{h.f}/B{r_ven}", PCT, "Crecen si se fabrica más", F_TOT)
r_mc = h.fila("MARGEN DE CONTRIBUCIÓN", f"=B{r_ven}-B{r_cv}", EUR, TOT,
              f"=B{h.f}/B{r_ven}", PCT, "De cada euro vendido quedan 27 céntimos", F_BIEN)
h.fila("Personal de producción", f"=-B{r_pp}", EUR, NEG, f"=B{h.f}/B{r_ven}", PCT,
       "Se trata como FIJO: Esteban no quiere despidos")
h.fila("Personal de estructura", f"=-B{r_pe}", EUR, NEG, f"=B{h.f}/B{r_ven}", PCT)
r_cf = h.fila("TOTAL COSTES FIJOS", f"=B{h.f-2}+B{h.f-1}", EUR, TOT,
              f"=B{h.f}/B{r_ven}", PCT, "Se pagan se fabrique o no", F_TOT)
h.fila("EBITDA reconstruido", f"=B{r_mc}-B{r_cf}", EUR, TOT, f"=B{h.f}/B{r_ven}", PCT)
h.fila("Diferencia con el EBITDA del Anexo 1", f"=B{h.f-1}-B{r_ebda}", EUR, TOT, None,
       PCT, "Cuadra al euro", F_BIEN)
h.hueco()

h.barra("¿CUÁNTA CAJA GENERÓ REALMENTE LA EMPRESA?", "→ VA AL INFORME")
h.explicar("Hay dos peticiones de dinero sobre la mesa: la casa de los fundadores y los "
           "643.000 € que exige Persán el primer año. La respuesta no está en el beneficio, "
           "está aquí. Alcopalet compró 376.062 € más de madera de la que consumió: eso es "
           "beneficio contable que salió del banco y sigue en el almacén.")
h.cabecera(["Concepto", "Importe", "", "Comentario"])
h.fila("EBITDA", f"=B{r_ebda}")
h.fila("(−) Aumento de existencias", f"=-B{r_vex}", EUR, NEG, None, PCT,
       "Madera comprada que aún no se ha vendido")
h.fila("(−) Gastos financieros", f"=B{r_gf}")
h.fila("(−) Impuestos", f"=B{r_imp}")
r_caja = h.fila("CAJA OPERATIVA APROXIMADA", f"=SUM(B{h.f-4}:B{h.f-1})", EUR, TOT, None,
                PCT, "NEGATIVA, con 279.364 € de beneficio", F_MAL, grande=BIG)
h.fila("Resultado neto contable", f"=B{r_rn}")
h.fila("BRECHA entre beneficio y caja", f"=B{r_rn}-B{r_caja}", EUR, TOT, None, PCT,
       "El beneficio está en el almacén, no en el banco", F_MAL)
h.hueco()

h.barra("PUNTO MUERTO Y APALANCAMIENTO", "→ VA AL INFORME")
h.explicar("Mide el colchón: cuánto puede caer la facturación antes de entrar en pérdidas. "
           "Sirve para dos cosas opuestas y ambas útiles. Demuestra que la empresa está sana "
           "HOY, lo que te da credibilidad ante el jurado. Y demuestra que Persán empeora las "
           "dos variables a la vez: sube los costes fijos en 108.000 € (campa y leasing) y "
           "encima aporta margen negativo.")
h.cabecera(["Concepto", "Importe", "", "Comentario"])
r_ratio = h.fila("Ratio de margen de contribución", f"=B{r_mc}/B{r_ven}", PCT, TOT)
h.fila("Punto muerto (EBITDA = 0)", f"=B{r_cf}/B{r_ratio}", EUR, NEG, None, PCT,
       "Ventas necesarias para no perder caja")
r_pm = h.fila("PUNTO MUERTO (resultado = 0)", f"=(B{r_cf}-B{r_am}-B{r_gf})/B{r_ratio}",
              EUR, TOT, None, PCT, "Ventas necesarias para no entrar en pérdidas", F_TOT)
h.fila("Ventas reales 2025", f"=B{r_ven}")
h.fila("MARGEN DE SEGURIDAD", f"=(B{r_ven}-B{r_pm})/B{r_ven}", PCT, TOT, None, PCT,
       "Puede caer un 30% antes de perder dinero", F_BIEN, grande=VERDE)
h.fila("Apalancamiento operativo", f"=B{r_mc}/B{r_ebda}", MUL, TOT, None, PCT,
       "1% menos de margen mueve el EBITDA un 2,75%")
h.hueco()

h.barra("COMPROBACIÓN: ¿CUADRA EL ANEXO 1 CON EL ANEXO 3?", "no va al informe")
h.explicar("Esto no se publica. Es lo que te autoriza a usar el escandallo del Anexo 3 sin "
           "que nadie te lo discuta. Si alguien dice que los anexos del caso no cuadran, "
           "esta es la respuesta.")
h.cabecera(["Concepto", "Importe", "", "Comentario"])
h.fila("Consumo real calculado desde el Anexo 1", f"=B{r_cons}")
h.fila("Materia prima según el Anexo 3", 2692732, EUR, AZUL, None, PCT,
       "8,27 € × 174.504 nuevos + 4,80 € × 260.330 usados")
h.fila("DESVIACIÓN", f"=B{h.f-2}-B{h.f-1}", EUR, TOT, None, PCT,
       "288 € sobre 2,7 millones: los anexos son coherentes", F_BIEN)
h.hueco()

h.barra("ENDEUDAMIENTO IMPLÍCITO", "turno de preguntas")
h.explicar("El caso no da balance, pero los gastos financieros divididos por el tipo del "
           "banco reconstruyen la deuda aproximada. Es tu respuesta si te preguntan «¿y no "
           "puede financiarlo el banco?».")
h.cabecera(["Concepto", "Importe", "", "Comentario"])
r_tipo = h.fila("Tipo de interés del banco", 0.075, PCT, AZUL)
r_deuda = h.fila("Deuda implícita", f"=-B{r_gf}/B{r_tipo}", EUR, TOT)
h.fila("Deuda / EBITDA — HOY", f"=B{r_deuda}/B{r_ebda}", MUL, TOT, None, PCT,
       "Por debajo de 2x: el banco presta sin problema", F_BIEN)
h.fila("Deuda estimada TRAS firmar Persán", f"=B{r_deuda}+190000+454356", EUR, NEG, None,
       PCT, "+190.000 leasing +454.356 circulante a 180 días")
h.fila("EBITDA estimado TRAS Persán", f"=B{r_ebda}-154800", EUR, NEG, None, PCT,
       "−46.800 margen −60.000 campa −48.000 leasing")
h.fila("Deuda / EBITDA — TRAS PERSÁN", f"=B{h.f-2}/B{h.f-1}", MUL, TOT, None, PCT,
       "Por encima de 4x el banco empieza a poner condiciones", F_MAL, grande=BIG)

# ══════════════════════════════════════════════════════════════════════════
# HOJA 2 — ESCANDALLO
# ══════════════════════════════════════════════════════════════════════════
e = Hoja(wb, "2. Escandallo", "2. El escandallo: cuánto cuesta un palet",
         "El Anexo 3 recalculado componente a componente. Aquí aparece el hallazgo central "
         "del caso: el contrato de Persán tiene margen NEGATIVO.")

e.barra("EL ESCANDALLO DEL ANEXO 3, RECALCULADO")
e.explicar("El caso publica el coste total de cada palet. Aquí se vuelve a sumar componente "
           "a componente en vez de copiarlo, para verificar que la suma da lo que dice. Lo "
           "importante es la última fila: el margen de contribución, que es lo que deja cada "
           "palet tras pagar lo que ese palet consume.")
e.cabecera(["Concepto de coste (€/palet)", "Manual actual", "Palet usado", "Proyecto Persán"])

fila_ing = e.f
for etiqueta, v1, v2, v3, nota in [
    ("Ingreso medio", 15.16, 7.45, 10.00, None),
    ("Materia prima (madera y clavos)", 8.27, 4.80, 8.27, None),
    ("Mano de obra directa", 2.18, 1.02, 0.42, None),
    ("Transporte", 0.62, 0.62, 0.53, None),
    ("Otros costes variables", 1.30, 0.58, 1.30, None),
]:
    ws = e.ws
    f = e.f
    ws.cell(f, 1, etiqueta).font = ETI
    for col, val in ((2, v1), (3, v2), (4, v3)):
        c = ws.cell(f, col, val)
        c.font = AZUL
        c.number_format = EU2
        c.alignment = Alignment(horizontal="center")
    e.f += 1

fila_coste = e.f
ws = e.ws
ws.cell(fila_coste, 1, "COSTE TOTAL POR PALET").font = TOT
for col in (2, 3, 4):
    L = chr(64 + col)
    c = ws.cell(fila_coste, col, f"=SUM({L}{fila_ing+1}:{L}{fila_coste-1})")
    c.font = TOT
    c.number_format = EU2
    c.alignment = Alignment(horizontal="center")
    c.fill = F_TOT
ws.cell(fila_coste, 1).fill = F_TOT
e.f += 1

fila_margen = e.f
ws.cell(fila_margen, 1, "MARGEN DE CONTRIBUCIÓN POR PALET").font = TOT
for col in (2, 3, 4):
    L = chr(64 + col)
    c = ws.cell(fila_margen, col, f"={L}{fila_ing}-{L}{fila_coste}")
    c.font = BIG if col == 4 else TOT
    c.number_format = EU2
    c.alignment = Alignment(horizontal="center")
    c.fill = F_MAL if col == 4 else F_BIEN
ws.cell(fila_margen, 1).fill = F_TOT
e.f += 1

e.fila("Margen sobre el precio de venta",
       f"=B{fila_margen}/B{fila_ing}", PCT, NEG, None, PCT)
ws = e.ws
for col in (3, 4):
    L = chr(64 + col)
    c = ws.cell(e.f - 1, col, f"={L}{fila_margen}/{L}{fila_ing}")
    c.font = NEG
    c.number_format = PCT
    c.alignment = Alignment(horizontal="center")
e.hueco()

e.barra("LO QUE ESTO SIGNIFICA")
e.explicar("Un margen de contribución negativo quiere decir que el palet no cubre ni "
           "siquiera lo que consume al fabricarse: madera, mano de obra, transporte y "
           "variables. No es un contrato de margen bajo, es un contrato que pierde dinero "
           "en cada unidad. Y por eso NINGÚN volumen lo arregla: cuantos más palets se "
           "fabriquen, más se pierde. Tampoco lo arregla ninguna máquina, porque lo que "
           "domina el coste es la madera, no la mano de obra.")
e.cabecera(["Concepto", "Valor", "", "Comentario"])
e.fila("Peso de la madera sobre el precio Persán", f"=D{fila_ing+1}/D{fila_ing}", PCT, TOT,
       None, PCT, "La madera se come el 83% del precio de venta", F_MAL, grande=BIG)
e.fila("Peso de la madera sobre el precio actual", f"=B{fila_ing+1}/B{fila_ing}", PCT, NEG,
       None, PCT, "En el negocio normal la madera es el 55%")
e.fila("Descuento de Persán sobre el precio medio",
       f"=D{fila_ing}/B{fila_ing}-1", PCT, TOT, None, PCT,
       "Persán paga un 34% menos que el cliente medio de palet nuevo", F_MAL)
e.fila("Ahorro de mano de obra del robot (€/palet)",
       f"=B{fila_ing+2}-D{fila_ing+2}", EU2, NEG, None, PCT,
       "El robot ahorra 1,76 €/palet de mano de obra...")
e.fila("Sobrecoste de vender a Persán (€/palet)",
       f"=B{fila_ing}-D{fila_ing}", EU2, TOT, None, PCT,
       "...pero el precio cede 5,16 €. El ahorro no compensa ni de lejos", F_MAL)
e.hueco()

e.barra("EL RIESGO QUE NADIE HA PUESTO EN EL CONTRATO")
e.explicar("El contrato fija el precio de venta durante CINCO AÑOS pero no fija el precio de "
           "la madera, que es el 83% de ese precio. Todo el riesgo de inflación lo asume "
           "Alcopalet. Esta tabla muestra qué pasa con el margen unitario si la madera sube.")
e.cabecera(["Si la madera sube...", "Coste por palet", "Margen por palet", "Pérdida anual (90.000 ud)"])
for infl in (0.00, 0.05, 0.10, 0.15, 0.20):
    ws = e.ws
    f = e.f
    c = ws.cell(f, 1, infl)
    c.font = AZUL
    c.number_format = PCT
    c = ws.cell(f, 2, f"=$D${fila_ing+1}*(1+A{f})+$D${fila_ing+2}+$D${fila_ing+3}+$D${fila_ing+4}")
    c.font = NEG
    c.number_format = EU2
    c.alignment = Alignment(horizontal="center")
    c = ws.cell(f, 3, f"=$D${fila_ing}-B{f}")
    c.font = NEG
    c.number_format = EU2
    c.alignment = Alignment(horizontal="center")
    c = ws.cell(f, 4, f"=C{f}*90000")
    c.font = TOT
    c.number_format = EUR
    c.alignment = Alignment(horizontal="center")
    e.f += 1
e.explicar("Conclusión: un contrato a precio fijo sobre una materia prima volátil traslada "
           "todo el riesgo al fabricante. Si se firma, debe llevar cláusula de revisión "
           "anual del precio de la madera.")
e.hueco()

e.barra("COMPROBACIÓN CONTRA LA CUENTA DE RESULTADOS", "no va al informe")
e.explicar("Si el escandallo multiplicado por las unidades vendidas reproduce las partidas "
           "del Anexo 1, es que el escandallo es fiable. Estas cuatro líneas son tu defensa "
           "si alguien cuestiona los datos.")
e.cabecera(["Partida", "Según escandallo", "Según Anexo 1", "Diferencia"])
UN, UU = 174504, 260330
for etiqueta, comp, anexo1 in [
    ("Materia prima", f"=B{fila_ing+1}*{UN}+C{fila_ing+1}*{UU}", 2692444),
    ("Transporte", f"=B{fila_ing+3}*{UN}+C{fila_ing+3}*{UU}", 268218),
    ("Otros costes variables", f"=B{fila_ing+4}*{UN}+C{fila_ing+4}*{UU}", 378979),
]:
    ws = e.ws
    f = e.f
    ws.cell(f, 1, etiqueta).font = ETI
    c = ws.cell(f, 2, comp); c.font = NEG; c.number_format = EUR
    c = ws.cell(f, 3, anexo1); c.font = AZUL; c.number_format = EUR
    c = ws.cell(f, 4, f"=B{f}-C{f}"); c.font = TOT; c.number_format = EUR
    e.f += 1
e.explicar("Las tres cuadran con desviaciones de menos del 0,5%. La cuarta partida, la mano "
           "de obra, NO cuadra con el Anexo 1 (646.000 € frente a 481.660 €), pero sí cuadra "
           "exactamente con la plantilla real: 10 operarios en palet nuevo y 7 en usado, a "
           "38.000 € cada uno. Es una diferencia de clasificación contable, no un error.")

# ══════════════════════════════════════════════════════════════════════════
# HOJA 3 — CAPACIDAD
# ══════════════════════════════════════════════════════════════════════════
c = Hoja(wb, "3. Capacidad", "3. ¿Cabe Persán en la fábrica?",
         "El caso da rendimientos por hora y por jornada, pero nunca la capacidad anual. "
         "Sin ella no se puede saber si el contrato es viable ni si la máquina hace falta.")

c.barra("DATOS DE PARTIDA (del enunciado)")
c.cabecera(["Dato", "Valor", "", "De dónde sale"])
r_meses = c.fila("Meses productivos al año", 11, NUM, AZUL, None, PCT, "Nota final del Anexo 3")
r_dmes = c.fila("Días productivos al mes", 22, NUM, AZUL, None, PCT, "Nota final del Anexo 3")
r_dias = c.fila("DÍAS PRODUCTIVOS AL AÑO", f"=B{r_meses}*B{r_dmes}", NUM, TOT, None, PCT,
                "11 × 22 = 242 días", F_TOT)
r_horas = c.fila("Horas por jornada", 8, NUM, AZUL, None, PCT,
                 "Turno único 7:00-15:00. Esteban descarta tarde y noche")
c.hueco()

c.barra("CAPACIDAD ANUAL DE CADA SISTEMA")
c.explicar("Aquí se convierte el rendimiento que da el caso en palets al año. Es el cálculo "
           "que permite responder si el contrato cabe o no.")
c.cabecera(["Sistema de fabricación", "Palets/día", "Palets/año", "Cómo se calcula"])

ws = c.ws
f = c.f
ws.cell(f, 1, "Línea automática (2019)").font = ETI
cc = ws.cell(f, 2, f"=65*0.75*B{r_horas}"); cc.font = NEG; cc.number_format = NUM
cc = ws.cell(f, 3, f"=B{f}*B{r_dias}"); cc.font = TOT; cc.number_format = NUM
ws.cell(f, 4, "65 palets/h × 75% rendimiento × 8 h").font = MINI
r_linea = f
c.f += 1

f = c.f
ws.cell(f, 1, "Fabricación manual").font = ETI
cc = ws.cell(f, 2, 500); cc.font = AZUL; cc.number_format = NUM
cc = ws.cell(f, 3, f"=B{f}*B{r_dias}"); cc.font = TOT; cc.number_format = NUM
ws.cell(f, 4, "500 palets/día entre 8 operarios = 62,5 cada uno").font = MINI
r_manual = f
c.f += 1

f = c.f
ws.cell(f, 1, "CAPACIDAD ACTUAL TOTAL").font = TOT
cc = ws.cell(f, 3, f"=C{r_linea}+C{r_manual}"); cc.font = TOT; cc.number_format = NUM
for col in range(1, 5):
    ws.cell(f, col).fill = F_TOT
r_capact = f
c.f += 1

f = c.f
ws.cell(f, 1, "Robot nuevo (si se compra)").font = ETI
cc = ws.cell(f, 2, "=400*0.9"); cc.font = NEG; cc.number_format = NUM
cc = ws.cell(f, 3, f"=B{f}*B{r_dias}"); cc.font = TOT; cc.number_format = NUM
ws.cell(f, 4, "400 palets/jornada × 90% rendimiento, 1 operario").font = MINI
r_robot = f
c.f += 1
c.hueco()

c.barra("EL DATO QUE CAMBIA LA CONVERSACIÓN", "→ VA AL INFORME")
c.explicar("Esteban dice «no llegamos» y por eso urge la máquina. Pero sus propios números "
           "dicen otra cosa: la fábrica está al 81%. Hay casi 41.000 palets al año de "
           "capacidad ociosa que se pueden fabricar hoy sin invertir un euro. El argumento "
           "de que la máquina es imprescindible por capacidad no se sostiene.")
c.cabecera(["Concepto", "Palets/año", "", "Comentario"])
r_prod = c.fila("Producción real de palet nuevo en 2025", 174504, NUM, AZUL, None, PCT,
                "Anexo 2")
c.fila("Capacidad teórica actual", f"=C{r_capact}", NUM, NEG)
c.fila("GRADO DE UTILIZACIÓN", f"=B{r_prod}/B{c.f-1}", PCT, TOT, None, PCT,
       "La fábrica NO está llena", F_CLAVE, grande=BIG)
c.fila("Capacidad ociosa disponible hoy", f"=B{c.f-2}-B{r_prod}", NUM, TOT, None, PCT,
       "Se puede fabricar esto sin invertir nada", F_CLAVE)
c.hueco()

c.barra("¿LLEGA EL ROBOT A CUBRIR PERSÁN?")
c.explicar("El Anexo 3 calcula la mano de obra de Persán dividiendo 38.000 € entre 90.000 "
           "unidades. Es decir, da por supuesto que el robot fabrica las 90.000. Conviene "
           "comprobarlo, porque no es así.")
c.cabecera(["Concepto", "Palets/año", "", "Comentario"])
r_rcap = c.fila("Capacidad anual del robot", f"=C{r_robot}", NUM, NEG)
r_pers = c.fila("Compromiso con Persán", 90000, NUM, AZUL, None, PCT, "90 mil unidades/año")
c.fila("DIFERENCIA", f"=B{r_rcap}-B{r_pers}", NUM, TOT, None, PCT,
       "El robot se queda CORTO: no llega al compromiso", F_MAL, grande=BIG)
c.fila("Déficit en % del compromiso", f"=B{c.f-1}/B{r_pers}", PCT, TOT, None, PCT,
       "Habría que completarlo con las otras líneas o con días extra", F_MAL)
c.hueco()

c.barra("¿CABE PERSÁN EN LA FÁBRICA?")
c.cabecera(["Escenario", "Capacidad", "Demanda total", "¿Encaja?"])
for etiqueta, cap, dem, nota in [
    ("Sin robot, sin Persán", f"=C{r_capact}", f"=B{r_prod}", None),
    ("Sin robot, CON Persán", f"=C{r_capact}", f"=B{r_prod}+B{r_pers}", None),
    ("Con robot, CON Persán", f"=C{r_capact}+C{r_robot}", f"=B{r_prod}+B{r_pers}", None),
]:
    ws = c.ws
    f = c.f
    ws.cell(f, 1, etiqueta).font = ETI
    cc = ws.cell(f, 2, cap); cc.font = NEG; cc.number_format = NUM
    cc = ws.cell(f, 3, dem); cc.font = NEG; cc.number_format = NUM
    cc = ws.cell(f, 4, f'=IF(B{f}>=C{f},"SÍ, cabe","NO cabe")')
    cc.font = TOT
    cc.alignment = Alignment(horizontal="center")
    c.f += 1
c.explicar("Conclusión: si se firma Persán hace falta capacidad nueva sí o sí, porque la "
           "actual se queda corta en unas 49.000 unidades. Pero si NO se firma, la máquina "
           "no hace falta por capacidad. Son dos conclusiones distintas y conviene no "
           "mezclarlas.")
c.hueco()

c.barra("LA OPORTUNIDAD QUE EL CASO NO SEÑALA", "→ VA AL INFORME")
c.explicar("La línea de palet usado está limitada a 30.000 unidades al mes por FALTA DE "
           "ESPACIO, no por falta de demanda. Y la campa que habría que alquilar para el "
           "retén de Persán cuesta 5.000 €/mes. Es decir: el mismo alquiler que se plantea "
           "para un contrato que pierde dinero desbloquearía una línea que sí gana dinero.")
c.cabecera(["Concepto", "Palets/año", "", "Comentario"])
r_usa = c.fila("Producción actual de palet usado", 260330, NUM, AZUL, None, PCT, "Anexo 2")
r_techo = c.fila("Techo logístico", "=30000*12", NUM, NEG, None, PCT, "30.000/mes por falta de espacio")
r_rec = c.fila("RECORRIDO DISPONIBLE", f"=B{r_techo}-B{r_usa}", NUM, TOT, None, PCT,
               "Se puede crecer esto si se resuelve el espacio", F_BIEN)
c.fila("Margen de contribución por palet usado", 0.43, EU2, AZUL, None, PCT, "Anexo 3")
c.fila("Contribución adicional posible", f"=B{r_rec}*B{c.f-1}", EUR, TOT, None, PCT,
       "No cubre por sí sola los 60.000 €/año de la campa, pero reparte su coste", F_BIEN)
c.fila("Operarios necesarios para llegar al techo",
       f"=B{r_techo}/(B{r_dias}*155)", NU2, NEG, None, PCT,
       "Hoy hay 7. Harían falta unos 2,6 más")
c.explicar("Y aquí encaja la pieza que falta: el robot libera personal de palet nuevo, y esta "
           "línea necesita personal. Esteban no quiere despedir a nadie; no haría falta.")

# ══════════════════════════════════════════════════════════════════════════
# HOJA 4 — CONTRATO PERSÁN
# ══════════════════════════════════════════════════════════════════════════
p = Hoja(wb, "4. Contrato Persán", "4. Cuánto cuesta firmar el contrato",
         "90.000 unidades/año · 10 €/unidad FIJO durante 5 años · cobro a 180 días · "
         "retén permanente de 1.000 palets en nave.")

p.barra("DATOS DEL CONTRATO Y COSTES ASOCIADOS")
p.cabecera(["Dato", "Valor", "", "De dónde sale"])
r_vol = p.fila("Volumen anual comprometido", 90000, NUM, AZUL, None, PCT, "Borrador del contrato")
r_pre = p.fila("Precio de venta unitario", 10.00, EU2, AZUL, None, PCT, "FIJO durante 5 años")
r_cu = p.fila("Coste variable unitario", 10.52, EU2, AZUL, None, PCT, "Anexo 3, hoja 2")
r_cobro = p.fila("Plazo de cobro", 180, NUM, AZUL, None, PCT, "Pago a 180 días")
r_int = p.fila("Tipo de interés del banco", 0.075, PCT, AZUL, None, PCT, "Apartado (c)")
r_campa = p.fila("Alquiler de la campa (€/año)", "=5000*12", EUR, NEG, None, PCT,
                 "5.000 €/mes por el terreno colindante")
r_leas = p.fila("Leasing máquina + camión (€/año)", "=4000*12", EUR, NEG, None, PCT,
                "4.000 €/mes por los dos activos")
r_ret = p.fila("Retén permanente exigido (palets)", 1000, NUM, AZUL, None, PCT,
               "Dos camiones en nave para servir en menos de 24 h")
p.hueco()

p.barra("LA CUENTA DE RESULTADOS DEL CONTRATO", "→ VA AL INFORME")
p.explicar("Esto es lo que aporta el contrato al resultado de la empresa, año a año. Fíjate "
           "en que el margen de contribución ya es negativo ANTES de sumar la campa y el "
           "leasing: el contrato no cubre ni lo que consume fabricarlo.")
p.cabecera(["Concepto", "Importe anual", "", "Comentario"])
p.fila("Ingresos", f"=B{r_vol}*B{r_pre}", EUR, NEG, None, PCT, "90.000 × 10 €")
p.fila("(−) Costes variables", f"=-B{r_vol}*B{r_cu}", EUR, NEG, None, PCT, "90.000 × 10,52 €")
r_mcp = p.fila("MARGEN DE CONTRIBUCIÓN", f"=B{p.f-2}+B{p.f-1}", EUR, TOT, None, PCT,
               "Negativo antes de cualquier coste fijo", F_MAL)
p.fila("(−) Alquiler de la campa", f"=-B{r_campa}")
p.fila("(−) Leasing", f"=-B{r_leas}")
r_sub = p.fila("SUBTOTAL", f"=B{r_mcp}+B{p.f-2}+B{p.f-1}", EUR, TOT, None, PCT, None, F_TOT)
p.hueco()

p.barra("EL COSTE OCULTO: COBRAR A 180 DÍAS")
p.explicar("Persán paga a seis meses. Eso significa que Alcopalet tiene permanentemente "
           "medio año de facturación de ese cliente sin cobrar, y ese dinero hay que "
           "financiarlo al 7,5%. El escandallo no lo recoge, pero es un coste real.")
p.cabecera(["Concepto", "Importe", "", "Comentario"])
r_cxc = p.fila("Dinero pendiente de cobro (media)", f"=B{r_vol}*B{r_pre}*B{r_cobro}/365",
               EUR, NEG, None, PCT, "Ingresos × 180/365")
r_stock = p.fila("Retén inmovilizado, a coste", f"=B{r_ret}*B{r_cu}", EUR, NEG, None, PCT,
                 "1.000 palets que hay que tener siempre en nave")
r_circ = p.fila("CIRCULANTE TOTAL INMOVILIZADO", f"=B{r_cxc}+B{r_stock}", EUR, TOT, None,
                PCT, "Dinero de Alcopalet atrapado en el contrato", F_TOT)
r_cfin = p.fila("Coste financiero anual", f"=-B{r_circ}*B{r_int}", EUR, TOT, None, PCT,
                "Al 7,5%", F_MAL)
p.hueco()

p.barra("EL RESULTADO FINAL", "→ VA AL INFORME")
p.cabecera(["Concepto", "Importe", "", "Comentario"])
r_imp_tot = p.fila("IMPACTO ANUAL DEL CONTRATO", f"=B{r_sub}+B{r_cfin}", EUR, TOT, None,
                   PCT, "Lo que resta al resultado de la empresa cada año", F_MAL, grande=BIG)
p.fila("Resultado neto actual de la empresa", 279364, EUR, AZUL, None, PCT, "Anexo 1")
p.fila("En % del resultado actual", f"=B{r_imp_tot}/B{p.f-1}", PCT, TOT, None, PCT,
       "Se lleva por delante dos tercios del beneficio de TODA la empresa", F_MAL)
p.fila("Impacto acumulado a 5 años", f"=B{r_imp_tot}*5", EUR, TOT, None, PCT,
       "El contrato dura 5 años sin revisión de precio", F_MAL)
p.fila("Resultado de la empresa tras firmar", f"=279364+B{r_imp_tot}", EUR, TOT, None, PCT,
       "La empresa quedaría al borde de las pérdidas")
p.hueco()

p.barra("¿A CUÁNTO HABRÍA QUE VENDER? EL PRECIO DE EQUILIBRIO", "→ VA AL INFORME")
p.explicar("Si el problema es el precio, la pregunta útil no es «¿firmo o no?» sino «¿a "
           "partir de qué precio interesa?». Estos cuatro niveles responden eso. El tercero "
           "es el importante: es el precio que hace que el contrato no reste ni sume.")
p.cabecera(["Nivel de cobertura", "Precio mínimo", "Subida necesaria", "Qué cubre"])
r_be1 = p.fila("1. Solo los costes variables", f"=B{r_cu}", EU2, NEG,
               f"=B{p.f}/B{r_pre}-1", PCT, "Madera, mano de obra, transporte y variables")
r_be2 = p.fila("2. + campa y leasing", f"=B{r_cu}+(B{r_campa}+B{r_leas})/B{r_vol}", EU2,
               NEG, f"=B{p.f}/B{r_pre}-1", PCT, "Añade los costes fijos que obliga a asumir")
r_be3 = p.fila("3. + coste del circulante",
               f"=(B{r_vol}*B{r_cu}+B{r_campa}+B{r_leas})/(B{r_vol}*(1-B{r_cobro}/365*B{r_int}))",
               EU2, TOT, f"=B{p.f}/B{r_pre}-1", PCT,
               "PRECIO DE EQUILIBRIO REAL. Por debajo, el contrato destruye valor",
               F_CLAVE, grande=BIG)
p.fila("4. + margen igual al del negocio actual", f"=B{r_be2}+2.79", EU2, NEG,
       f"=B{p.f}/B{r_pre}-1", PCT, "Para que Persán sea tan rentable como el resto")
p.hueco()

p.barra("SENSIBILIDAD: PRECIO NEGOCIADO FRENTE A PRECIO DE LA MADERA")
p.explicar("Las dos variables que de verdad mueven el resultado. Cada celda es el impacto "
           "anual del contrato. Los números en rojo son pérdidas. Sirve para saber hasta "
           "dónde hay que negociar y qué pasa si la madera se encarece.")

ws = p.ws
f = p.f
ws.cell(f, 1, "Madera ↓ / Precio →").font = CAB
ws.cell(f, 1).fill = F_CAB
precios = [10.00, 11.00, 12.00, 12.50, 13.00]
for i, pr in enumerate(precios):
    cc = ws.cell(f, 2 + i, pr)
    cc.font = CAB
    cc.fill = F_CAB
    cc.number_format = EU2
    cc.alignment = Alignment(horizontal="center")
    cc.border = BOR
fila_prec = f
p.f += 1

MADERA_BASE = 8.27
NO_MADERA = 0.42 + 0.53 + 1.30
for infl in (0.00, 0.05, 0.10, 0.15):
    f = p.f
    cc = ws.cell(f, 1, infl)
    cc.font = CAB
    cc.fill = F_CAB
    cc.number_format = PCT
    cc.alignment = Alignment(horizontal="center")
    for i, _ in enumerate(precios):
        L = chr(66 + i)
        formula = (f"=$B${r_vol}*({L}${fila_prec}-({MADERA_BASE}*(1+$A{f})+{NO_MADERA}))"
                   f"-$B${r_campa}-$B${r_leas}"
                   f"-({L}${fila_prec}*$B${r_vol}*$B${r_cobro}/365+$B${r_ret}*$B${r_cu})*$B${r_int}")
        cc = ws.cell(f, 2 + i, formula)
        cc.font = NEG
        cc.number_format = EUR
        cc.border = BOR
        cc.alignment = Alignment(horizontal="center")
    p.f += 1
p.explicar("Lectura: con el precio actual de 10 € el contrato pierde dinero incluso si la "
           "madera no sube nada. Hace falta llegar a unos 12,17 € solo para no perder. Y si "
           "la madera sube un 10%, ni siquiera 13 € bastan.")

# ══════════════════════════════════════════════════════════════════════════
# HOJA 5 — EL ROBOT
# ══════════════════════════════════════════════════════════════════════════
r = Hoja(wb, "5. El robot", "5. La máquina, como decisión independiente",
         "El caso presenta máquina y contrato como una sola decisión. Separarlos es la clave "
         "del análisis: la tecnología puede ser buena aunque el contrato no lo sea.")

r.barra("COMPARATIVA DE LOS TRES SISTEMAS")
r.explicar("Antes de hablar de dinero conviene ver qué compra realmente la máquina. Y lo que "
           "compra no es sobre todo capacidad: es FLEXIBILIDAD.")
r.cabecera(["Indicador", "Manual", "Línea 2019", "Robot nuevo"])
ws = r.ws
for etiqueta, v1, v2, v3, fmt in [
    ("Palets por jornada", 500, "=65*0.75*8", "=400*0.9", NUM),
    ("Operarios necesarios", 8, 2, 1, NUM),
    ("Cambio de formato (horas)", "—", 6, "=25/60", NU2),
    ("Mantenimiento", "—", "Complejo y constante", "Solo limpieza diaria", None),
]:
    f = r.f
    ws.cell(f, 1, etiqueta).font = ETI
    for col, val in ((2, v1), (3, v2), (4, v3)):
        cc = ws.cell(f, col, val)
        cc.font = AZUL if not (isinstance(val, str) and val.startswith("=")) else NEG
        if fmt and not isinstance(val, str):
            cc.number_format = fmt
        elif fmt and isinstance(val, str) and val.startswith("="):
            cc.number_format = fmt
        cc.alignment = Alignment(horizontal="center")
    r.f += 1
f = r.f
ws.cell(f, 1, "Productividad (palets/operario/día)").font = TOT
for col in (2, 3, 4):
    L = chr(64 + col)
    cc = ws.cell(f, col, f"={L}{r.f-4}/{L}{r.f-3}")
    cc.font = TOT
    cc.number_format = NU2
    cc.alignment = Alignment(horizontal="center")
    cc.fill = F_TOT
ws.cell(f, 1).fill = F_TOT
r.f += 1
r.explicar("El dato decisivo es el cambio de formato: SEIS HORAS de parada en la línea de "
           "2019 frente a menos de media hora en el robot. Alcopalet gana dinero en series "
           "especiales (el 114×114 de Don Simón, las exigencias de Tetra Pak). Con seis horas "
           "de parada una serie corta es inviable; con media hora, es negocio.")
r.hueco()

r.barra("EL AHORRO TEÓRICO DE MANO DE OBRA")
r.explicar("Si el robot fabrica lo mismo que varios operarios manuales, el ahorro parece "
           "evidente. Este bloque lo cuantifica. Ojo: es TEÓRICO, y el bloque siguiente "
           "explica por qué.")
r.cabecera(["Concepto", "Valor", "", "Comentario"])
r_rcap2 = r.fila("Capacidad anual del robot (palets)", "=400*0.9*242", NUM, NEG, None, PCT,
                 "Ver hoja 3")
r_prodop = r.fila("Productividad manual (palets/operario/día)", "=500/8", NU2, NEG)
r_equiv = r.fila("Operarios manuales equivalentes", f"=B{r_rcap2}/(B{r_prodop}*242)", NU2,
                 TOT, None, PCT, "Los que harían falta para fabricar lo mismo a mano")
r.fila("Operarios que exige el robot", 1, NUM, AZUL)
r_libre = r.fila("Operarios liberados", f"=B{r_equiv}-B{r.f-1}", NU2, TOT)
r_ahorro = r.fila("AHORRO TEÓRICO (€/año)", f"=B{r_libre}*38000", EUR, TOT, None, PCT,
                  "A 38.000 € por operario", F_BIEN, grande=VERDE)
r.fila("Coste anual del leasing solo de la máquina",
       "=-40000*(0.075/12)/(1-(1+0.075/12)^(-60))*12", EUR, NEG, None, PCT,
       "40.000 € al 7,5% a 5 años. SIN el tráiler")
r.fila("Plazo de recuperación (meses)", f"=40000/B{r_ahorro}*12", NU2, TOT, None, PCT,
       "Menos de tres meses", F_BIEN)
r.hueco()

r.barra("PERO: «NO QUIERO DESPIDOS»", "→ VA AL INFORME")
r.explicar("Esteban afirma expresamente que no quiere despedir a nadie, y que el turno único "
           "es un valor diferencial para retener a la plantilla. Un ahorro de mano de obra "
           "que no se traduce ni en menos nóminas ni en más producción vendida NO es un "
           "ahorro real. Esta es la hipótesis más atacable de todo el trabajo, y por eso hay "
           "que ponerla encima de la mesa antes de que la pongan ellos.")
r.cabecera(["Concepto", "Valor", "", "Comentario"])
r_mon = r.fila("% de mano de obra liberada que se monetiza", 0.00, PCT, NEG, None, PCT,
               "HIPÓTESIS MÍA. Cámbiala y mira qué pasa abajo", F_HIP)
r.fila("Ahorro REAL de mano de obra", f"=B{r_ahorro}*B{r_mon}", EUR, TOT)
r.fila("(−) Leasing de la máquina", "=-40000*(0.075/12)/(1-(1+0.075/12)^(-60))*12", EUR, NEG)
r.fila("RESULTADO REAL DEL ROBOT AISLADO", f"=B{r.f-2}+B{r.f-1}", EUR, TOT, None, PCT,
       "Con 0% de monetización, el robot solo aporta coste", F_MAL, grande=BIG)
r.explicar("Tres vías para que ese ahorro sea real, y ninguna implica despedir: reasignar "
           "personal a la línea de palet usado, que necesita 2,6 operarios más y solo le "
           "falta espacio; crecer en palet nuevo a precio de mercado; o no reponer las bajas "
           "naturales de una plantilla con más de una década de antigüedad.")
r.hueco()

r.barra("LA PREGUNTA CORRECTA: ¿A QUÉ DEDICAR LA MÁQUINA?", "→ VA AL INFORME")
r.explicar("La capacidad del robot es la misma se use para lo que se use. Lo que cambia es a "
           "quién se le vende. Esta comparación es, probablemente, el argumento más potente "
           "del informe.")
r.cabecera(["Destino de la producción", "Precio", "Margen/palet", "Margen anual"])
# El transporte difiere: 0,53 € solo aplica a las rutas de Persán con el tráiler
# propio a plena carga; para el mercado ordinario rige el 0,62 € del Anexo 3.
for etiqueta, precio, transporte in [
    ("Vender a Persán a 10 €", 10.00, 0.53),
    ("Vender al mercado actual a 15,16 €", 15.16, 0.62),
]:
    coste = 8.27 + 0.42 + transporte + 1.30
    f = r.f
    ws.cell(f, 1, etiqueta).font = ETI
    cc = ws.cell(f, 2, precio); cc.font = AZUL; cc.number_format = EU2
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 3, f"=B{f}-{round(coste, 2)}"); cc.font = TOT; cc.number_format = EU2
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 4, f"=C{f}*{int(400*0.9*242)}"); cc.font = TOT; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    r.f += 1
r.fila("COSTE DE OPORTUNIDAD", f"=D{r.f-1}-D{r.f-2}", EUR, TOT, None, PCT,
       "Lo que cuesta dedicar la máquina a Persán en vez de al mercado", F_MAL, grande=BIG)
r.explicar("Matiz honesto y necesario: esto supone que hay demanda para colocar esos palets a "
           "15,16 €. Con la fábrica al 81%, hoy la restricción es la demanda, no la "
           "capacidad. Es la principal debilidad de este argumento y hay que decirlo antes de "
           "que lo diga el jurado.")

# ══════════════════════════════════════════════════════════════════════════
# HOJA 6 — LAS OPCIONES
# ══════════════════════════════════════════════════════════════════════════
o = Hoja(wb, "6. Las opciones", "6. Las cinco opciones, comparadas",
         "Todo lo anterior confluye aquí. Cada columna es una de las alternativas que el "
         "caso pone sobre la mesa.")

o.barra("QUÉ ES CADA OPCIÓN")
o.cabecera(["Opción", "", "", "En qué consiste"])
for et, desc in [
    ("A — No hacer nada", "Ni máquina ni contrato ni dividendo. La empresa sigue igual."),
    ("B — Persán + robot", "Lo que propone Esteban: firmar el contrato y comprar la máquina "
                           "con el tráiler."),
    ("C — Dividendo sin invertir", "Lo que piden los padres: repartir y no comprometer nada."),
    ("D — Robot sin Persán", "Comprar la máquina SOLA, sin tráiler, y no firmar el contrato."),
    ("E — Robot + Persán renegociado", "Comprar la máquina y firmar solo si Persán sube el "
                                       "precio y acorta el plazo de pago."),
]:
    o.fila(et, "", EUR, NEG, None, PCT, desc)
o.hueco()

o.barra("IMPACTO ECONÓMICO ANUAL", "→ VA AL INFORME")
o.explicar("Todas las cifras son el efecto sobre el resultado anual de la empresa, partiendo "
           "del beneficio actual de 279.364 €.")
o.cabecera(["Concepto", "B — Persán + robot", "D — Robot solo", "E — Renegociado"])

PRECIO_REN, DIAS_REN = 12.50, 90
ws = o.ws
filas_op = {}
for clave, etiqueta, fB, fD, fE, nota in [
    ("mc", "Margen de contribución del contrato",
     "=90000*(10-10.52)", "0", f"=90000*({PRECIO_REN}-10.52)",
     "En D no hay contrato, luego no hay margen"),
    ("ah", "Ahorro real de mano de obra", "0", "0", "0",
     "Cero mientras no se reasigne personal (hipótesis conservadora)"),
    ("ca", "Alquiler de la campa", "=-5000*12", "0", "=-5000*12", None),
    ("le", "Leasing", "=-4000*12",
     "=-40000*(0.075/12)/(1-(1+0.075/12)^(-60))*12", "=-4000*12",
     "En D solo se financia la máquina: no hace falta tráiler"),
    ("ci", "Coste financiero del circulante",
     "=-(90000*10*180/365+1000*10.52)*0.075", "0",
     f"=-(90000*{PRECIO_REN}*{DIAS_REN}/365+1000*10.52)*0.075",
     "En E se supone además cobrar a 90 días en vez de 180"),
]:
    f = o.f
    ws.cell(f, 1, etiqueta).font = ETI
    for col, formula in ((2, fB), (3, fD), (4, fE)):
        cc = ws.cell(f, col, formula)
        cc.font = NEG
        cc.number_format = EUR
        cc.alignment = Alignment(horizontal="center")
    if nota:
        ws.cell(f, 1).comment = None
    filas_op[clave] = f
    o.f += 1

f = o.f
ws.cell(f, 1, "IMPACTO ANUAL SOBRE EL RESULTADO").font = TOT
for col in (2, 3, 4):
    L = chr(64 + col)
    cc = ws.cell(f, col, f"=SUM({L}{filas_op['mc']}:{L}{filas_op['ci']})")
    cc.font = TOT
    cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    cc.fill = F_TOT
ws.cell(f, 1).fill = F_TOT
fila_impacto = f
o.f += 1

f = o.f
ws.cell(f, 1, "RESULTADO NETO PROYECTADO").font = TOT
for col in (2, 3, 4):
    L = chr(64 + col)
    cc = ws.cell(f, col, f"=279364+{L}{fila_impacto}")
    cc.font = TOT
    cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    cc.fill = F_TOT
ws.cell(f, 1).fill = F_TOT
o.f += 1
o.fila("(Opciones A y C dejan el resultado en 279.364 €, sin cambio)", "", EUR, NEG,
       None, PCT, "A no hace nada; C solo saca caja, no altera el resultado")
o.hueco()

o.barra("IMPACTO EN LA CAJA DEL PRIMER AÑO", "→ VA AL INFORME")
o.explicar("Tan importante como el resultado. La empresa generó caja NEGATIVA en 2025, así "
           "que lo que decida tiene que caber en la tesorería.")
o.cabecera(["Concepto", "B — Persán + robot", "D — Robot solo", "E — Renegociado"])
f = o.f
ws.cell(f, 1, "Impacto en el resultado").font = ETI
for col in (2, 3, 4):
    L = chr(64 + col)
    cc = ws.cell(f, col, f"={L}{fila_impacto}")
    cc.font = NEG
    cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
o.f += 1
f = o.f
ws.cell(f, 1, "(−) Circulante inmovilizado el primer año").font = ETI
for col, formula in ((2, "=-(90000*10*180/365+1000*10.52)"), (3, "0"),
                     (4, f"=-(90000*{PRECIO_REN}*{DIAS_REN}/365+1000*10.52)")):
    cc = ws.cell(f, col, formula)
    cc.font = NEG
    cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
o.f += 1
f = o.f
ws.cell(f, 1, "NECESIDAD DE CAJA DEL PRIMER AÑO").font = TOT
for col in (2, 3, 4):
    L = chr(64 + col)
    cc = ws.cell(f, col, f"={L}{f-2}+{L}{f-1}")
    cc.font = TOT
    cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    cc.fill = F_MAL
ws.cell(f, 1).fill = F_MAL
o.f += 1
o.fila("Caja operativa generada en 2025 (referencia)", -60555, EUR, AZUL, None, PCT,
       "Ver hoja 1. La empresa NO generó caja el año pasado", F_MAL)
o.hueco()

o.barra("COMPARACIÓN CUALITATIVA")
o.explicar("No todo se decide con euros. Estos son los criterios que el propio caso pone "
           "sobre la mesa, y que hay que respetar para que la propuesta sea realista.")
o.cabecera(["Criterio", "B — Persán + robot", "D — Robot solo", "E — Renegociado"])
for criterio, vB, vD, vE in [
    ("¿El contrato cubre sus costes?", "NO", "n/a", "SÍ"),
    ("¿Cabe en la tesorería?", "NO", "SÍ", "AJUSTADO"),
    ("¿Respeta «sin despidos»?", "SÍ", "SÍ", "SÍ"),
    ("¿Respeta el turno único?", "SÍ", "SÍ", "SÍ"),
    ("¿Permite el plan de crecimiento a 2030?", "SÍ", "PARCIAL", "SÍ"),
    ("¿Da seguridad a los fundadores?", "NO", "PARCIAL", "PARCIAL"),
    ("¿Reduce la dependencia de mano de obra?", "SÍ", "SÍ", "SÍ"),
    ("Riesgo de concentración de cliente", "ALTO", "BAJO", "MEDIO"),
]:
    f = o.f
    ws.cell(f, 1, criterio).font = ETI
    for col, v in ((2, vB), (3, vD), (4, vE)):
        cc = ws.cell(f, col, v)
        cc.font = TOT
        cc.alignment = Alignment(horizontal="center")
        if v in ("NO", "ALTO"):
            cc.fill = F_MAL
        elif v in ("SÍ", "BAJO"):
            cc.fill = F_BIEN
        else:
            cc.fill = F_CLAVE
    o.f += 1
o.hueco()

o.barra("LA LECTURA")
for t in [
    "El caso presenta la máquina y el contrato como una sola decisión, pero tienen plazos "
    "distintos: la máquina hay que decidirla HOY, mientras que el contrato de Persán todavía "
    "es un borrador. No hay ninguna razón para resolver las dos cosas en la misma reunión.",
    "La máquina se justifica por flexibilidad (media hora de cambio de formato frente a seis "
    "horas) y por reducir la dependencia de una mano de obra difícil de contratar. No por "
    "capacidad: la fábrica está al 81%.",
    "Si se compra la máquina sin el tráiler, el leasing baja de 48.000 € a unos 9.600 € al "
    "año. Los 150.000 € del camión solo hacen falta para el retén de Persán.",
    "El contrato a 10 € pierde dinero en cada palet. El precio de equilibrio está en 12,17 €. "
    "Perder a Persán cuesta una oportunidad; firmarlo cuesta casi un millón de euros en "
    "cinco años.",
    "Hay tres palancas de negociación, no una: el precio, el plazo de cobro y —la que nadie "
    "mira— la especificación de la madera. Los 8,27 € son el estándar Tetra Pak, nacido de un "
    "incidente en una aduana china. Persán fabrica detergentes en Sevilla: probablemente no "
    "necesita esa calidad.",
    "Y a los fundadores hay que decirles que rechazar Persán a 10 € es precisamente lo que "
    "protege su jubilación: si se firma, el beneficio cae a 90.000 € y no habrá dividendo "
    "en cinco años.",
]:
    o.texto(t)

wb.save("/home/user/DANISANTELMO/Alcopalet_Analisis_Completo.xlsx")
print("Guardado:", [s.title for s in wb.worksheets])
