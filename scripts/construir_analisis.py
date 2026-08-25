"""
ALCOPALET (DTI-1399) — Análisis completo del caso.

Un libro, diez hojas, todas con la misma estructura:
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

_s = Side(style="thin", color="BFBFBF")
BOR = Border(left=_s, right=_s, top=_s, bottom=_s)

EUR = '#,##0 "€";(#,##0) "€";"-"'
EU2 = '#,##0.00 "€";(#,##0.00) "€";"-"'
EU3 = '#,##0.000 "€";(#,##0.000) "€";"-"'
PCT = '0.0%;(0.0%);"-"'
NUM = '#,##0;(#,##0);"-"'
NU2 = '#,##0.00;(#,##0.00);"-"'
MUL = '0.00"x"'

ANCHOS = {"A": 50, "B": 17, "C": 15, "D": 15, "E": 40}
ANCHO_TXT = 130

# Constantes del caso usadas para precalcular literales
DIAS = 242
CMO = 38000
MOD_MANUAL = CMO * 8 / (500 * DIAS)          # 2,5124 €/palet
MOD_LINEA = CMO * 2 / (65 * 0.75 * 8 * DIAS)  # 0,8053 €/palet


class Hoja:
    def __init__(self, wb, nombre, titulo, subtitulo):
        self.ws = wb.create_sheet(nombre)
        ws = self.ws
        ws.sheet_view.showGridLines = False
        for col, w in ANCHOS.items():
            ws.column_dimensions[col].width = w
        ws["A1"] = titulo
        ws["A1"].font = TIT
        ws["A2"] = subtitulo
        ws["A2"].font = SUB
        ws.merge_cells("A2:E2")
        ws.row_dimensions[2].height = 12.5 * (len(subtitulo) // ANCHO_TXT + 1) + 4
        self.f = 4
        ws.freeze_panes = "A4"

    def barra(self, texto, etiqueta=None):
        ws, f = self.ws, self.f
        ws.cell(f, 1, texto).font = SEC
        for c in range(1, 6):
            ws.cell(f, c).fill = F_SEC
        if etiqueta:
            c = ws.cell(f, 5, etiqueta)
            c.font = Font(name=FUENTE, size=9, bold=True, color="FFFFFF")
            c.alignment = Alignment(horizontal="right")
        self.f += 1

    def explicar(self, texto, relleno=F_EXP):
        ws, f = self.ws, self.f
        c = ws.cell(f, 1, texto)
        c.font = EXP
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=5)
        for col in range(1, 6):
            ws.cell(f, col).fill = relleno
        ws.row_dimensions[f].height = 12.5 * (len(texto) // ANCHO_TXT + 1) + 5
        self.f += 1

    def cabecera(self, cols):
        ws, f = self.ws, self.f
        for i, t in enumerate(cols):
            c = ws.cell(f, 1 + i, t)
            c.font = CAB
            c.fill = F_CAB
            c.alignment = Alignment(horizontal="left" if i == 0 else "center",
                                    wrap_text=True, vertical="center")
            c.border = BOR
        ws.row_dimensions[f].height = 26
        self.f += 1

    def fila(self, etiqueta, *valores, fmt=EUR, fuente=NEG, nota=None,
             relleno=None, grande=None, fmts=None):
        """valores van a las columnas B, C, D...  nota a la columna E."""
        ws, f = self.ws, self.f
        ws.cell(f, 1, etiqueta).font = TOT if (relleno or grande) else ETI
        for i, v in enumerate(valores):
            if v is None or v == "":
                continue
            c = ws.cell(f, 2 + i, v)
            c.font = grande if grande else (TOT if relleno else fuente)
            c.number_format = (fmts[i] if fmts else fmt)
            if i > 0:
                c.alignment = Alignment(horizontal="center")
        if nota:
            n = ws.cell(f, 5, nota)
            n.font = MINI
            n.alignment = Alignment(wrap_text=True, vertical="center")
        if relleno:
            for col in range(1, 6):
                ws.cell(f, col).fill = relleno
        self.f += 1
        return f

    def texto(self, txt, vinieta=True):
        ws, f = self.ws, self.f
        c = ws.cell(f, 1, ("• " if vinieta else "") + txt)
        c.font = ETI
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=5)
        ws.row_dimensions[f].height = 12.5 * (len(txt) // ANCHO_TXT + 1) + 6
        self.f += 1

    def hueco(self, n=1):
        self.f += n


wb = Workbook()
wb.remove(wb.active)

# ══════════════════════════════════════════════════════════════════════════
# 0. GUÍA Y RECOMENDACIÓN
# ══════════════════════════════════════════════════════════════════════════
g = Hoja(wb, "0. Guía", "ALCOPALET — Guía y recomendación",
         "Caso DTI-1399 · Programa Lydes 2026 · Instituto Internacional San Telmo. "
         "Empieza por aquí: la recomendación, por qué, y qué hay en cada hoja.")

g.barra("LA RECOMENDACIÓN EN UNA FRASE")
g.explicar("COMPRAR LA MÁQUINA HOY, SIN EL TRÁILER. NO FIRMAR PERSÁN A 10 €. VOLVER A LA MESA "
           "CON TRES PALANCAS: precio, especificación de la madera y plazo de cobro. "
           "Son dos decisiones distintas que el caso presenta como una sola, y tienen plazos "
           "distintos: el fabricante de la máquina pide respuesta HOY, mientras que el "
           "contrato de Persán todavía es un borrador sin firmar.", F_CLAVE)
g.hueco()

g.barra("POR QUÉ COMPRAR LA MÁQUINA")
g.cabecera(["Razón", "Dato", "", "", "Explicación"])
g.fila("NO es por capacidad", 0.810, fmt=PCT, fuente=TOT,
       nota="La fábrica está al 81%. Sobran 40.876 palets/año sin invertir un euro. "
            "El argumento de Esteban («no llegamos») no se sostiene con sus propios números.")
g.fila("Es por FLEXIBILIDAD", 6.0, fmt=NU2, fuente=TOT,
       nota="Horas de parada por cambio de formato en la línea de 2019, frente a menos de "
            "media hora en el robot. Alcopalet gana dinero en series especiales (114×114 de "
            "Don Simón, exigencias de Tetra Pak). Con 6 horas de parada, una serie corta es "
            "inviable; con media hora, es negocio.")
g.fila("Es por riesgo de mano de obra", 5.76, fmt=NU2, fuente=TOT,
       nota="Operarios manuales equivalentes que sustituye. El caso dice que la mano de obra "
            "es «cada vez más difícil de gestionar y contratar», y menciona absentismo los "
            "lunes. La máquina reduce esa dependencia estructural.")
g.fila("Y porque es barata", 40000, fmt=EUR, fuente=TOT,
       nota="Sobre un beneficio de 279.364 €. Apuesta pequeña con mucho valor de opción. "
            "Comprada SOLA (sin tráiler), el leasing baja de 48.000 € a 9.618 €/año.")
g.hueco()

g.barra("POR QUÉ NO FIRMAR A 10 €")
g.cabecera(["Razón", "Dato", "", "", "Explicación"])
g.fila("Pierde dinero en cada palet", -0.55, fmt=EU2, fuente=BIG,
       nota="Margen de contribución real por unidad, con el coste de mano de obra ponderado. "
            "Ningún volumen lo arregla: cuantos más palets, más se pierde.")
g.fila("Pierde incluso SIN mano de obra", -0.10, fmt=EU2, fuente=BIG,
       nota="Aunque el operario fuese gratis, el precio no cubre madera + transporte + "
            "variables (8,27 + 0,53 + 1,30 = 10,10 €). Este es el argumento definitivo: "
            "el problema no es la productividad, es el precio.")
g.fila("El precio de equilibrio está muy por encima", 12.17, fmt=EU2, fuente=BIG,
       nota="Frente a los 10 € ofrecidos. Faltan 2,17 € por unidad solo para no perder.")
g.fila("Y el riesgo de la madera es todo tuyo", 0.827, fmt=PCT, fuente=BIG,
       nota="Peso de la madera sobre el precio de venta. El contrato fija el precio 5 años "
            "pero no lleva cláusula de revisión de materia prima.")
g.fila("Coste de equivocarse", -944383, fmt=EUR, fuente=BIG,
       nota="Impacto acumulado a 5 años. Perder a Persán cuesta una oportunidad; firmarlo "
            "cuesta casi un millón de euros.", relleno=F_MAL)
g.hueco()

g.barra("LOS TÉRMINOS DEL CONTRATO IDEAL", "→ hoja 6")
g.explicar("La pregunta útil no es «¿firmo o no?» sino «¿en qué condiciones sí?». Hay tres "
           "palancas, no una. Pedir 12,17 € a secas es difícil; combinar las tres es una "
           "negociación ganable.")
g.cabecera(["Variable", "Hoy", "Mínimo", "Objetivo", "Por qué"])
g.fila("Precio por palet", 10.00, 12.17, 12.50, fmt=EU2, fuente=TOT,
       nota="El mínimo es el equilibrio exacto. Por debajo, el contrato destruye valor.")
g.fila("Coste de la madera", 8.27, 8.00, 7.50, fmt=EU2, fuente=TOT,
       nota="LA PALANCA QUE NADIE MIRA. Los 8,27 € son el estándar Tetra Pak (madera seca, "
            "blanca, lijada, sin astillas), nacido de un incidente de moho en una aduana "
            "china. Persán fabrica detergentes en Sevilla: casi con seguridad no necesita "
            "esa calidad. Re-especificar el palet es más fácil que subir el precio.")
g.fila("Plazo de cobro (días)", 180, 120, 90, fmt=NUM, fuente=TOT,
       nota="Cada día de aplazamiento inmoviliza caja al 7,5%. Pasar de 180 a 90 libera la "
            "mitad del circulante.")
g.fila("Cláusula de revisión de la madera", "NO", "Anual", "Anual", fmt=NUM, fuente=TOT,
       fmts=[NUM, NUM, NUM],
       nota="Innegociable. Sin ella, cinco años de riesgo de materia prima a cargo de Alcopalet.")
g.fila("Retén exigido (palets)", 1000, 1000, 500, fmt=NUM, fuente=TOT,
       nota="O que el coste de mantenerlo se repercuta en el precio.")
g.fila("Duración (años)", 5, 3, 3, fmt=NUM, fuente=TOT,
       nota="Menos años = menos exposición a un precio que se puede quedar obsoleto.")
g.hueco()
g.cabecera(["Resultado según lo que se consiga", "Impacto anual", "", "", "Lectura"])
g.fila("Contrato actual (10 €, madera 8,27, 180 días)", -191398, fmt=EUR, fuente=TOT,
       nota="Destruye valor", relleno=F_MAL)
g.fila("Mínimo aceptable (12,17 €, madera 8,27, 180 días)", -3321, fmt=EUR, fuente=TOT,
       nota="Ni gana ni pierde. No merece la pena por sí solo", relleno=F_CLAVE)
g.fila("Objetivo realista (12,50 €, madera 8,27, 90 días)", 46085, fmt=EUR, fuente=TOT,
       nota="Solo con precio y plazo", relleno=F_BIEN)
g.fila("CONTRATO IDEAL (12,50 €, madera 7,50, 90 días)", 115443, fmt=EUR, fuente=VERDE,
       nota="Las tres palancas juntas. Este es el objetivo de la negociación", relleno=F_BIEN)
g.hueco()

g.barra("CÓMO SE LEE ESTE LIBRO")
g.explicar("Todas las hojas tienen la misma estructura: banda azul con el título del bloque, "
           "debajo una línea gris que explica PARA QUÉ sirve ese cálculo, y después los "
           "números. Se lee de arriba abajo, como un documento.")
g.cabecera(["Código de color", "", "", "", "Qué significa"])
for et, nota in [
    ("Texto AZUL", "Dato tomado literalmente del enunciado o de sus anexos."),
    ("Texto NEGRO", "Resultado calculado con fórmula. Si cambias un dato azul, se recalcula."),
    ("Fondo AMARILLO", "Hipótesis MÍA, no dada por el caso. Hay que declararla y defenderla."),
    ("Fondo VERDE", "Comprobación que cuadra, o resultado favorable."),
    ("Fondo ROJO", "Número que revela un problema."),
]:
    g.fila(et, nota=nota)
g.hueco()

g.barra("QUÉ HAY EN CADA HOJA")
g.cabecera(["Hoja", "", "", "", "Contenido y pregunta que responde"])
for hoja, desc in [
    ("1. Datos del caso", "Todo el enunciado y los tres anexos en un solo sitio, con su origen."),
    ("2. Costes y equilibrio", "Fijos frente a variables, punto muerto, apalancamiento y caja. "
                               "¿Cuánto deja cada venta y hay dinero disponible?"),
    ("3. Margen por líneas", "Margen de contribución de palet nuevo, usado y Persán, con las "
                             "DOS definiciones. ¿Qué línea gana dinero de verdad?"),
    ("4. Capacidad", "Capacidad oficial (90%) y realista (85/80/75%), y coste medio ponderado "
                     "de fabricar las 90.000. ¿Cabe Persán y a qué coste real?"),
    ("5. Contrato Persán", "Qué hace falta para servirlo, costes fijos adicionales e impacto "
                           "anual. ¿Cuánto cuesta firmar?"),
    ("6. Contrato ideal", "Los términos que habría que negociar y qué se gana con cada uno."),
    ("7. Inversión máquina", "VAN, TIR y plazo de recuperación según a qué se dedique."),
    ("8. Escenarios", "Las cinco alternativas comparadas en resultado, caja y criterios."),
    ("9. Anexos informe", "Los cuatro anexos que llevaría al documento final, y qué va en cada uno."),
]:
    g.fila(hoja, nota=desc)
g.hueco()

g.barra("MIS HIPÓTESIS (esto NO lo dice el caso)")
g.explicar("Cinco cosas que he tenido que suponer. Van marcadas en amarillo donde se usan. "
           "Hay que declararlas en el informe: el jurado valora más una hipótesis explícita "
           "que un número sin origen.")
g.cabecera(["Hipótesis", "Valor", "", "", "Por qué y qué pasa si falla"])
g.fila("Rendimiento realista de la máquina", 0.80, fmt=PCT, fuente=TOT, relleno=F_HIP,
       nota="El caso dice 90%, pero es la estimación de quien quiere comprarla. La línea de "
            "2019 se compró y rinde al 75%. Si el robot rinde al 80%, no llega a las 90.000 "
            "y hay que completarlas con otra línea, más cara.")
g.fila("Coste de la madera re-especificada", 7.50, fmt=EU2, fuente=TOT, relleno=F_HIP,
       nota="Supone que Persán acepta un palet estándar en vez del estándar Tetra Pak. Es "
            "la hipótesis más rentable y también la más incierta: hay que contrastarla con "
            "el cliente antes de prometerla.")
g.fila("Precio renegociado objetivo", 12.50, fmt=EU2, fuente=TOT, relleno=F_HIP,
       nota="Por encima del equilibrio de 12,17 €. Si Persán no pasa de 11 €, no interesa.")
g.fila("Plazo de cobro renegociado (días)", 90, fmt=NUM, fuente=TOT, relleno=F_HIP,
       nota="La mitad de lo pactado. Es lo más fácil de conseguir de las tres palancas.")
g.fila("% de mano de obra liberada que se monetiza", 0.00, fmt=PCT, fuente=TOT, relleno=F_HIP,
       nota="Esteban NO quiere despidos. Si nadie se reasigna, el ahorro del robot es cero. "
            "Es la hipótesis más conservadora posible, a propósito.")
g.hueco()

g.barra("ADVERTENCIAS DE HONESTIDAD")
for t in [
    "El coste del circulante es una aproximación (saldo medio × tipo de interés), no un "
    "estado de tesorería. El caso no da balance, así que se ignoran clientes y proveedores.",
    "El caso no dice si hay demanda para llenar la capacidad del robot a precio de mercado. "
    "Con la fábrica al 81%, hoy la restricción es la demanda, no la capacidad. Es la mayor "
    "debilidad del argumento del coste de oportunidad.",
    "Que Persán acepte un palet de menor especificación es una hipótesis razonable pero NO "
    "verificada. Si el cliente exige la misma calidad, esa palanca desaparece.",
    "El Anexo 3 publica un coste manual de 12,36 €; la suma exacta de sus componentes da "
    "12,37 €. Diferencia de redondeo, sin efecto en las conclusiones.",
    "La mano de obra del escandallo (646.000 €) no coincide con la partida de personal de "
    "producción del Anexo 1 (481.660 €). El escandallo sí cuadra con la plantilla real "
    "(17 operarios × 38.000 €), por eso se ha tomado como base.",
]:
    g.texto(t)

# ══════════════════════════════════════════════════════════════════════════
# 1. DATOS DEL CASO
# ══════════════════════════════════════════════════════════════════════════
d = Hoja(wb, "1. Datos del caso", "1. Todos los datos del caso, en un solo sitio",
         "Cada cifra con su origen. Todo lo que aparece en azul en las demás hojas sale de aquí.")

d.barra("A. PARÁMETROS DE CÁLCULO")
d.cabecera(["Dato", "Valor", "", "", "De dónde sale"])
r_mes = d.fila("Meses productivos al año", 11, fmt=NUM, fuente=AZUL,
               nota="Nota final del Anexo 3")
r_dm = d.fila("Días productivos al mes", 22, fmt=NUM, fuente=AZUL,
              nota="Nota final del Anexo 3")
d.fila("DÍAS PRODUCTIVOS AL AÑO", f"=B{r_mes}*B{r_dm}", fmt=NUM, fuente=TOT,
       relleno=F_TOT, nota="11 × 22 = 242")
d.fila("Horas por jornada", 8, fmt=NUM, fuente=AZUL,
       nota="Turno único 7:00-15:00. Esteban descarta turnos de tarde y noche")
d.fila("Coste medio de mano de obra (€/año)", 38000, fmt=EUR, fuente=AZUL,
       nota="Sueldos y cotizaciones. Apartado (d)")
d.fila("Coste medio de despido (€)", 30000, fmt=EUR, fuente=AZUL,
       nota="Elevado por la antigüedad. Apartado (d). No se usa: no hacen falta despidos")
d.fila("Tipo de interés de la financiación", 0.075, fmt=PCT, fuente=AZUL, nota="Apartado (c)")
d.hueco()

d.barra("B. EL CONTRATO DE PERSÁN")
d.cabecera(["Dato", "Valor", "", "", "De dónde sale"])
d.fila("Volumen anual comprometido (palets)", 90000, fmt=NUM, fuente=AZUL, nota="Borrador")
d.fila("Precio de venta unitario", 10.00, fmt=EU2, fuente=AZUL,
       nota="FIJO durante los 5 años, sin cláusula de revisión")
d.fila("Duración (años)", 5, fmt=NUM, fuente=AZUL)
d.fila("Plazo de cobro (días)", 180, fmt=NUM, fuente=AZUL, nota="Pago a 180 días")
d.fila("Retén permanente exigido (palets)", 1000, fmt=NUM, fuente=AZUL,
       nota="Dos camiones en nave para servir en menos de 24 h")
d.hueco()

d.barra("C. LOS TRES SISTEMAS DE FABRICACIÓN")
d.cabecera(["Dato", "Línea 2019", "Manual", "Robot nuevo", "Comentario"])
d.fila("Capacidad nominal", 65, 500, 400, fmt=NUM, fuente=AZUL,
       nota="Línea: palets/hora. Manual y robot: palets/jornada")
d.fila("Rendimiento considerado", 0.75, 1.00, 0.90, fmt=PCT, fuente=AZUL,
       nota="La línea de 2019 al 75% por mantenimiento complejo; el robot al 90% según Esteban")
d.fila("Operarios necesarios", 2, 8, 1, fmt=NUM, fuente=AZUL)
d.fila("Cambio de formato (horas)", 6, "", 0.42, fmt=NU2, fuente=AZUL,
       nota="6 HORAS frente a 25 minutos. La diferencia decisiva")
d.fila("Inversión", 360000, "", 40000, fmt=EUR, fuente=AZUL,
       nota="La línea de 2019 costó 360.000 € y se amortiza a 10 años")
d.hueco()

d.barra("D. LÍNEA DE PALET USADO (Palet Ojeda)")
d.cabecera(["Dato", "Valor", "", "", "Comentario"])
d.fila("Operarios", 7, fmt=NUM, fuente=AZUL)
d.fila("Productividad (palets/persona/día)", 155, fmt=NUM, fuente=AZUL)
d.fila("Techo logístico (palets/mes)", 30000, fmt=NUM, fuente=AZUL,
       nota="Límite por FALTA DE ESPACIO, no por falta de demanda. Clave del caso")
d.hueco()

d.barra("E. INVERSIONES ADICIONALES SI SE FIRMA PERSÁN")
d.cabecera(["Concepto", "Importe", "", "", "De dónde sale"])
d.fila("Alquiler campa (€/mes)", 5000, fmt=EUR, fuente=AZUL, nota="Terreno colindante. Apartado (a)")
d.fila("Nuevo tráiler y remolque", 150000, fmt=EUR, fuente=AZUL,
       nota="La logística ya está saturada. Apartado (b). SOLO hace falta para Persán")
d.fila("Leasing máquina + camión (€/mes)", 4000, fmt=EUR, fuente=AZUL,
       nota="Cuota conjunta de los dos activos. Apartado (c)")
d.hueco()

d.barra("F. CUENTA DE RESULTADOS 2025 (Anexo 1)")
d.cabecera(["Concepto", "Importe", "", "", "Comentario"])
for et, v, nota in [
    ("Ventas netas", 4585267, None),
    ("Variación de existencias", 376062, "Ingreso contable, pero NO es caja"),
    ("Aprovisionamientos", -3068506, "COMPRAS del ejercicio, no consumo"),
    ("Gastos de personal — Producción", -481660, None),
    ("Gastos de transporte", -268218, "2 camiones, ya al 100% de su capacidad"),
    ("Gastos de personal — Estructura", -311292, None),
    ("Otros costes de explotación", -378979, "60% es palet nuevo; varían con el volumen"),
    ("Amortización", -36143, None),
    ("Gastos financieros", -44046, "Al 7,5% implican unos 587.000 € de deuda"),
    ("Impuestos", -93121, None),
    ("RESULTADO NETO", 279364, None),
]:
    d.fila(et, v, fmt=EUR, fuente=AZUL, nota=nota,
           relleno=F_TOT if et.isupper() else None)
d.hueco()

d.barra("G. LÍNEAS DE NEGOCIO (Anexo 2) Y ESCANDALLO (Anexo 3)")
d.cabecera(["Concepto", "Palet nuevo", "Palet usado", "Persán", "Comentario"])
d.fila("Ventas netas", 2645224, 1940042, "", fmt=EUR, fuente=AZUL)
d.fila("Resultado neto", 158882, 120482, "", fmt=EUR, fuente=AZUL)
d.fila("Unidades vendidas", 174504, 260330, 90000, fmt=NUM, fuente=AZUL)
d.fila("Ingreso medio (€/palet)", 15.16, 7.45, 10.00, fmt=EU2, fuente=AZUL)
d.fila("Materia prima (€/palet)", 8.27, 4.80, 8.27, fmt=EU2, fuente=AZUL,
       nota="IDÉNTICA en manual y Persán: el robot no abarata la madera")
d.fila("Mano de obra directa (€/palet)", 2.18, 1.02, 0.42, fmt=EU2, fuente=AZUL,
       nota="La de Persán supone 90.000 ud con 1 operario. Ver hoja 4: no es realista")
d.fila("Transporte (€/palet)", 0.62, 0.62, 0.53, fmt=EU2, fuente=AZUL)
d.fila("Otros costes variables (€/palet)", 1.30, 0.58, 1.30, fmt=EU2, fuente=AZUL)

# ══════════════════════════════════════════════════════════════════════════
# 2. COSTES Y EQUILIBRIO
# ══════════════════════════════════════════════════════════════════════════
h = Hoja(wb, "2. Costes y equilibrio", "2. Costes fijos, variables y punto de equilibrio",
         "Reordena la cuenta de resultados por comportamiento en vez de por naturaleza. "
         "Es el cálculo que permite juzgar un pedido nuevo.")

h.barra("LA CUENTA TAL COMO LA DA EL CASO")
h.explicar("Los subtotales se calculan con fórmula en vez de copiarlos, para comprobar que "
           "cuadran con lo publicado. Es control de calidad, no análisis.")
h.cabecera(["Concepto", "Importe 2025", "% s/ventas", "", "Comentario"])
_rv = h.f
r_ven = h.fila("VENTAS NETAS", 4585267, f"=B{_rv}/$B${_rv}", fmt=EUR, fuente=AZUL,
               relleno=F_TOT, fmts=[EUR, PCT])


def pyg(et, val, nota=None):
    return h.fila(et, val, f"=B{h.f}/$B${r_ven}", fmt=EUR, fuente=AZUL, nota=nota,
                  fmts=[EUR, PCT])


r_vex = pyg("Variación de existencias", 376062, "Ingreso contable, pero NO es caja")
r_apr = pyg("Aprovisionamientos", -3068506, "Son COMPRAS, no consumo")
r_mb = h.fila("MARGEN BRUTO", f"=SUM(B{r_ven}:B{r_apr})", f"=B{h.f}/$B${r_ven}",
              fmt=EUR, fuente=TOT, relleno=F_TOT, fmts=[EUR, PCT])
r_pp = pyg("Gastos de personal — Producción", -481660)
r_tr = pyg("Gastos de transporte", -268218)
r_pe = pyg("Gastos de personal — Estructura", -311292)
r_oc = pyg("Otros costes de explotación", -378979)
r_ebda = h.fila("EBITDA", f"=B{r_mb}+SUM(B{r_pp}:B{r_oc})", f"=B{h.f}/$B${r_ven}",
                fmt=EUR, fuente=TOT, relleno=F_TOT, fmts=[EUR, PCT])
r_am = pyg("Amortización", -36143)
r_ebit = h.fila("EBIT", f"=B{r_ebda}+B{r_am}", f"=B{h.f}/$B${r_ven}", fmt=EUR,
                fuente=TOT, relleno=F_TOT, fmts=[EUR, PCT])
r_gf = pyg("Gastos financieros", -44046)
r_bai = h.fila("RESULTADO ANTES DE IMPUESTOS", f"=B{r_ebit}+B{r_gf}", f"=B{h.f}/$B${r_ven}",
               fmt=EUR, fuente=TOT, relleno=F_TOT, fmts=[EUR, PCT])
r_imp = pyg("Impuestos", -93121)
r_rn = h.fila("RESULTADO NETO", f"=B{r_bai}+B{r_imp}", f"=B{h.f}/$B${r_ven}", fmt=EUR,
              fuente=TOT, relleno=F_TOT, fmts=[EUR, PCT])
h.hueco()

h.barra("SEPARAR COSTES FIJOS DE VARIABLES", "→ VA AL INFORME")
h.explicar("El caso ordena los gastos por naturaleza (personal, transporte, otros). Para "
           "decidir sobre un pedido NUEVO eso no sirve: hay que reordenarlos por "
           "comportamiento, es decir, cuáles crecen si fabricas más y cuáles se pagan igual. "
           "Lo que queda tras pagar los variables es el margen de contribución, y ese es el "
           "único criterio válido para juzgar el contrato. Al final reconstruye el EBITDA "
           "exacto del Anexo 1: la reordenación es indiscutible.")
h.cabecera(["Concepto", "Importe", "% s/ventas", "", "Comentario"])
r_cons = h.fila("Consumo real de materia prima", f"=-B{r_apr}-B{r_vex}", f"=B{h.f}/B{r_ven}",
                fmt=EUR, fmts=[EUR, PCT], nota="Compras menos lo que quedó en el almacén")
h.fila("Transporte", f"=-B{r_tr}", f"=B{h.f}/B{r_ven}", fmt=EUR, fmts=[EUR, PCT])
h.fila("Otros costes de explotación", f"=-B{r_oc}", f"=B{h.f}/B{r_ven}", fmt=EUR,
       fmts=[EUR, PCT])
r_cv = h.fila("TOTAL COSTES VARIABLES", f"=SUM(B{h.f-3}:B{h.f-1})", f"=B{h.f}/B{r_ven}",
              fmt=EUR, fuente=TOT, relleno=F_TOT, fmts=[EUR, PCT],
              nota="Crecen si se fabrica más")
r_mc = h.fila("MARGEN DE CONTRIBUCIÓN", f"=B{r_ven}-B{r_cv}", f"=B{h.f}/B{r_ven}", fmt=EUR,
              fuente=TOT, relleno=F_BIEN, fmts=[EUR, PCT],
              nota="De cada euro vendido quedan 27 céntimos")
h.fila("Personal de producción", f"=-B{r_pp}", f"=B{h.f}/B{r_ven}", fmt=EUR, fmts=[EUR, PCT],
       nota="Se trata como FIJO: Esteban no quiere despidos")
h.fila("Personal de estructura", f"=-B{r_pe}", f"=B{h.f}/B{r_ven}", fmt=EUR, fmts=[EUR, PCT])
r_cf = h.fila("TOTAL COSTES FIJOS", f"=B{h.f-2}+B{h.f-1}", f"=B{h.f}/B{r_ven}", fmt=EUR,
              fuente=TOT, relleno=F_TOT, fmts=[EUR, PCT], nota="Se pagan se fabrique o no")
h.fila("EBITDA reconstruido", f"=B{r_mc}-B{r_cf}", f"=B{h.f}/B{r_ven}", fmt=EUR, fuente=TOT,
       fmts=[EUR, PCT])
h.fila("Diferencia con el EBITDA del Anexo 1", f"=B{h.f-1}-B{r_ebda}", fmt=EUR, fuente=TOT,
       relleno=F_BIEN, nota="Cuadra al euro")
h.hueco()

h.barra("PUNTO DE EQUILIBRIO", "→ VA AL INFORME")
h.explicar("Cuánto puede caer la facturación antes de entrar en pérdidas. Sirve para dos "
           "cosas opuestas y ambas útiles: demuestra que la empresa está sana HOY, lo que da "
           "credibilidad, y demuestra que Persán empeora las dos variables a la vez, porque "
           "sube los costes fijos en 108.000 € y encima aporta margen negativo.")
h.cabecera(["Concepto", "Importe", "", "", "Comentario"])
r_ratio = h.fila("Ratio de margen de contribución", f"=B{r_mc}/B{r_ven}", fmt=PCT, fuente=TOT)
h.fila("Punto de equilibrio (EBITDA = 0)", f"=B{r_cf}/B{r_ratio}", fmt=EUR,
       nota="Ventas necesarias para no perder caja")
r_pm = h.fila("PUNTO DE EQUILIBRIO (resultado = 0)",
              f"=(B{r_cf}-B{r_am}-B{r_gf})/B{r_ratio}", fmt=EUR, fuente=TOT, relleno=F_TOT,
              nota="Ventas necesarias para no entrar en pérdidas")
h.fila("Ventas reales 2025", f"=B{r_ven}", fmt=EUR)
h.fila("MARGEN DE SEGURIDAD", f"=(B{r_ven}-B{r_pm})/B{r_ven}", fmt=PCT, fuente=TOT,
       relleno=F_BIEN, grande=VERDE, nota="Puede caer un 30% antes de perder dinero")
h.fila("Apalancamiento operativo", f"=B{r_mc}/B{r_ebda}", fmt=MUL, fuente=TOT,
       nota="1% menos de margen mueve el EBITDA un 2,75%. Los errores de precio se amplifican")
h.hueco()
h.cabecera(["Efecto de Persán sobre el equilibrio", "Antes", "Después", "", "Comentario"])
h.fila("Costes fijos", f"=B{r_cf}", f"=B{r_cf}+108000", fmt=EUR, fuente=TOT,
       nota="Campa 60.000 € + leasing 48.000 €")
h.fila("Punto de equilibrio", f"=B{r_pm}", f"=(B{r_cf}+108000-B{r_am}-B{r_gf})/B{r_ratio}",
       fmt=EUR, fuente=TOT, relleno=F_MAL,
       nota="Sube casi 400.000 € de facturación necesaria, y el contrato no aporta margen")
h.hueco()

h.barra("¿CUÁNTA CAJA GENERÓ REALMENTE LA EMPRESA?", "→ VA AL INFORME")
h.explicar("Hay dos peticiones de dinero sobre la mesa: la casa de los fundadores y los "
           "643.000 € que exige Persán el primer año. La respuesta no está en el beneficio, "
           "está aquí. Alcopalet compró 376.062 € más de madera de la que consumió: es "
           "beneficio contable que salió del banco y sigue en el almacén.")
h.cabecera(["Concepto", "Importe", "", "", "Comentario"])
h.fila("EBITDA", f"=B{r_ebda}", fmt=EUR)
h.fila("(−) Aumento de existencias", f"=-B{r_vex}", fmt=EUR, nota="Madera comprada sin vender")
h.fila("(−) Gastos financieros", f"=B{r_gf}", fmt=EUR)
h.fila("(−) Impuestos", f"=B{r_imp}", fmt=EUR)
r_caja = h.fila("CAJA OPERATIVA APROXIMADA", f"=SUM(B{h.f-4}:B{h.f-1})", fmt=EUR, fuente=TOT,
                relleno=F_MAL, grande=BIG, nota="NEGATIVA, con 279.364 € de beneficio")
h.fila("BRECHA entre beneficio y caja", f"=B{r_rn}-B{r_caja}", fmt=EUR, fuente=TOT,
       relleno=F_MAL, nota="El beneficio está en el almacén, no en el banco")

# ══════════════════════════════════════════════════════════════════════════
# 3. MARGEN POR LÍNEAS
# ══════════════════════════════════════════════════════════════════════════
m = Hoja(wb, "3. Margen por líneas", "3. Margen de contribución por líneas de negocio",
         "Cuánto deja realmente cada tipo de palet. Con DOS definiciones, porque la respuesta "
         "cambia según se considere la mano de obra fija o variable.")

m.barra("POR QUÉ HAY DOS DEFINICIONES")
m.explicar("El escandallo del Anexo 3 trata la mano de obra como un coste variable, y por eso "
           "la resta al calcular el margen. Pero Esteban dice expresamente que NO quiere "
           "despidos: un coste que no se puede eliminar no es variable, es fijo. Las dos "
           "lecturas son legítimas y conviene enseñar ambas. La segunda —sin mano de obra— "
           "es la que da el argumento más demoledor contra Persán.")
m.hueco()

m.barra("DEFINICIÓN 1: MARGEN CON MANO DE OBRA (como el Anexo 3)")
m.cabecera(["Concepto (€/palet)", "Palet nuevo", "Palet usado", "Persán", "Comentario"])
fi = m.f
m.fila("Ingreso medio", 15.16, 7.45, 10.00, fmt=EU2, fuente=AZUL)
m.fila("(−) Materia prima", 8.27, 4.80, 8.27, fmt=EU2, fuente=AZUL)
m.fila("(−) Mano de obra directa", 2.18, 1.02, 0.42, fmt=EU2, fuente=AZUL)
m.fila("(−) Transporte", 0.62, 0.62, 0.53, fmt=EU2, fuente=AZUL)
m.fila("(−) Otros variables", 1.30, 0.58, 1.30, fmt=EU2, fuente=AZUL)
r_c1 = m.fila("COSTE TOTAL", f"=SUM(B{fi+1}:B{fi+4})", f"=SUM(C{fi+1}:C{fi+4})",
              f"=SUM(D{fi+1}:D{fi+4})", fmt=EU2, fuente=TOT, relleno=F_TOT)
r_m1 = m.fila("MARGEN CON MANO DE OBRA", f"=B{fi}-B{r_c1}", f"=C{fi}-C{r_c1}",
              f"=D{fi}-D{r_c1}", fmt=EU2, fuente=TOT, relleno=F_TOT)
m.fila("Unidades vendidas", 174504, 260330, 90000, fmt=NUM, fuente=AZUL)
r_t1 = m.fila("CONTRIBUCIÓN TOTAL (€/año)", f"=B{r_m1}*B{m.f-1}", f"=C{r_m1}*C{m.f-1}",
              f"=D{r_m1}*D{m.f-1}", fmt=EUR, fuente=TOT, relleno=F_TOT)
m.hueco()

m.barra("DEFINICIÓN 2: MARGEN SIN MANO DE OBRA (mano de obra como coste fijo)", "→ VA AL INFORME")
m.explicar("Si la plantilla se paga igual se fabrique o no —que es lo que implica «no quiero "
           "despidos»—, el margen relevante para aceptar un pedido es este. Fíjate en la "
           "columna de Persán.")
m.cabecera(["Concepto (€/palet)", "Palet nuevo", "Palet usado", "Persán", "Comentario"])
fi2 = m.f
m.fila("Ingreso medio", f"=B{fi}", f"=C{fi}", f"=D{fi}", fmt=EU2)
m.fila("(−) Materia prima", f"=B{fi+1}", f"=C{fi+1}", f"=D{fi+1}", fmt=EU2)
m.fila("(−) Transporte", f"=B{fi+3}", f"=C{fi+3}", f"=D{fi+3}", fmt=EU2)
m.fila("(−) Otros variables", f"=B{fi+4}", f"=C{fi+4}", f"=D{fi+4}", fmt=EU2)
r_c2 = m.fila("COSTE SIN MANO DE OBRA", f"=SUM(B{fi2+1}:B{fi2+3})",
              f"=SUM(C{fi2+1}:C{fi2+3})", f"=SUM(D{fi2+1}:D{fi2+3})", fmt=EU2,
              fuente=TOT, relleno=F_TOT)
r_m2 = m.fila("MARGEN SIN MANO DE OBRA", f"=B{fi2}-B{r_c2}", f"=C{fi2}-C{r_c2}",
              f"=D{fi2}-D{r_c2}", fmt=EU2, fuente=TOT, relleno=F_TOT)
m.explicar("ESTE ES EL ARGUMENTO DEFINITIVO CONTRA EL CONTRATO: aunque el operario del robot "
           "fuese GRATIS, Persán seguiría perdiendo 10 céntimos por palet. El precio de 10 € "
           "no cubre ni la madera (8,27) más el transporte (0,53) más los variables (1,30), "
           "que suman 10,10 €. El problema no es la productividad ni la tecnología: es el "
           "precio.", F_MAL)
m.hueco()

m.barra("COMPROBACIÓN CONTRA LA CUENTA DE RESULTADOS", "no va al informe")
m.explicar("El margen sin mano de obra, multiplicado por las unidades, debe reproducir el "
           "margen de contribución que sale de la cuenta de resultados (hoja 2).")
m.cabecera(["Concepto", "Palet nuevo", "Palet usado", "TOTAL", "Comentario"])
r_tot2 = m.fila("Contribución sin mano de obra (€/año)", f"=B{r_m2}*B{r_t1-1}",
                f"=C{r_m2}*C{r_t1-1}", f"=B{m.f}+C{m.f}", fmt=EUR, fuente=TOT, relleno=F_TOT)
m.fila("Margen de contribución de la hoja 2", "", "", 1245626, fmt=EUR, fuente=AZUL)
m.fila("DESVIACIÓN", "", "", f"=D{r_tot2}-D{m.f-1}", fmt=EUR, fuente=TOT, relleno=F_BIEN,
       nota="Menos de 900 € sobre 1,2 millones: los dos caminos llevan al mismo sitio")
m.hueco()

m.barra("LA LECTURA")
for t in [
    "El palet usado deja 1,45 € por unidad sin contar mano de obra, frente a 4,97 € del "
    "nuevo. Pero el usado necesita 7 operarios para 260.330 palets y el nuevo necesita 10 "
    "para 174.504: por persona, el usado rinde más del doble en unidades.",
    "La línea de palet usado está bloqueada por ESPACIO, no por demanda. Y el espacio "
    "cuesta 5.000 €/mes, el mismo alquiler que se plantea para el retén de Persán.",
    "Persán es la única de las tres columnas con margen negativo, en las dos definiciones. "
    "No es un pedido de margen bajo: es un pedido que resta.",
]:
    m.texto(t)

# ══════════════════════════════════════════════════════════════════════════
# 4. CAPACIDAD
# ══════════════════════════════════════════════════════════════════════════
c = Hoja(wb, "4. Capacidad", "4. Capacidad real y coste de fabricar las 90.000",
         "El caso da rendimientos por hora y por jornada, pero nunca la capacidad anual. "
         "Y da por hecho que el robot fabrica las 90.000 unidades. Ninguna de las dos cosas "
         "es evidente.")

c.barra("CAPACIDAD ANUAL DE CADA SISTEMA")
c.explicar("Convierte el rendimiento que da el caso en palets al año. Es lo que permite "
           "responder si el contrato cabe en la fábrica.")
c.cabecera(["Sistema", "Palets/día", "Días/año", "Palets/año", "Cómo se calcula"])
_rl = c.f
r_lin = c.fila("Línea automática (2019)", "=65*0.75*8", 242, f"=B{_rl}*C{_rl}", fmt=NUM,
               fmts=[NUM, NUM, NUM], nota="65 palets/h × 75% × 8 h de jornada")
_rm = c.f
r_man = c.fila("Fabricación manual", 500, 242, f"=B{_rm}*C{_rm}", fmt=NUM, fuente=AZUL,
               fmts=[NUM, NUM, NUM], nota="500/día entre 8 operarios = 62,5 cada uno")
r_cap = c.fila("CAPACIDAD ACTUAL TOTAL", "", "", f"=D{r_lin}+D{r_man}", fmt=NUM, fuente=TOT,
               relleno=F_TOT)
c.hueco()

c.barra("EL DATO QUE CAMBIA LA CONVERSACIÓN", "→ VA AL INFORME")
c.explicar("Esteban dice «no llegamos» y por eso urge la máquina. Pero sus propios números "
           "dicen otra cosa: la fábrica está al 81%. Hay casi 41.000 palets al año que se "
           "pueden fabricar hoy sin invertir un euro. El argumento de que la máquina es "
           "imprescindible POR CAPACIDAD no se sostiene. Hay que comprarla, pero por otras "
           "razones.")
c.cabecera(["Concepto", "Palets/año", "", "", "Comentario"])
r_prod = c.fila("Producción real de palet nuevo 2025", 174504, fmt=NUM, fuente=AZUL,
                nota="Anexo 2")
c.fila("Capacidad teórica actual", f"=D{r_cap}", fmt=NUM)
c.fila("GRADO DE UTILIZACIÓN", f"=B{r_prod}/B{c.f-1}", fmt=PCT, fuente=TOT, relleno=F_CLAVE,
       grande=BIG, nota="La fábrica NO está llena")
c.fila("Capacidad ociosa disponible hoy", f"=B{c.f-2}-B{r_prod}", fmt=NUM, fuente=TOT,
       relleno=F_CLAVE, nota="Se puede fabricar esto sin invertir nada")
c.hueco()

c.barra("CAPACIDAD DE LA MÁQUINA: OFICIAL FRENTE A REALISTA", "→ VA AL INFORME")
c.explicar("El 90% de rendimiento es la estimación de Esteban, que es quien quiere comprar la "
           "máquina. Conviene contrastarlo: la línea de 2019 también se compró con "
           "expectativas y hoy se considera al 75%. Además el robot lo lleva UN SOLO "
           "operario, y el caso menciona absentismo los lunes: si falta esa persona, la "
           "máquina para. Esta tabla muestra qué pasa con cada supuesto.")
c.cabecera(["Rendimiento", "Palets/día", "Palets/año", "Déficit vs 90.000", "Lectura"])
fila_rend = c.f
for rend, nota in [
    (0.90, "OFICIAL. Lo que dice el caso. Aun así NO llega a las 90.000"),
    (0.85, "Prudente. Faltarían 7.720 unidades"),
    (0.80, "REALISTA en el primer año, con curva de aprendizaje"),
    (0.75, "Pesimista: el mismo rendimiento que la línea de 2019"),
]:
    f = c.f
    ws = c.ws
    cc = ws.cell(f, 1, rend); cc.font = AZUL; cc.number_format = PCT
    cc = ws.cell(f, 2, f"=400*A{f}"); cc.font = NEG; cc.number_format = NUM
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 3, f"=B{f}*242"); cc.font = TOT; cc.number_format = NUM
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 4, f"=C{f}-90000"); cc.font = TOT; cc.number_format = NUM
    cc.alignment = Alignment(horizontal="center")
    cc.fill = F_MAL
    ws.cell(f, 5, nota).font = MINI
    ws.cell(f, 5).alignment = Alignment(wrap_text=True, vertical="center")
    c.f += 1
c.explicar("Conclusión: EN NINGÚN ESCENARIO el robot cubre por sí solo las 90.000 unidades. "
           "Ni siquiera con el 90% oficial. Y el Anexo 3 calcula la mano de obra dividiendo "
           "38.000 € entre 90.000 unidades, es decir, DA POR HECHO que sí las cubre. Esa "
           "hipótesis está mal y encarece el contrato.", F_MAL)
c.hueco()

c.barra("COSTE MEDIO PONDERADO DE FABRICAR LAS 90.000", "→ VA AL INFORME")
c.explicar("Si el robot no llega, el resto hay que fabricarlo con otra línea, y cada línea "
           "tiene un coste de mano de obra distinto por palet. El coste real de servir a "
           "Persán es la media ponderada de las dos fuentes. Se completa con la línea de "
           "2019, que es la más barata de las dos disponibles.")
c.cabecera(["Coste de mano de obra por sistema", "€/palet", "", "", "Cómo se calcula"])
r_mman = c.fila("Fabricación manual", round(MOD_MANUAL, 4), fmt=EU3, fuente=TOT,
                nota="38.000 € × 8 operarios ÷ (500 × 242 días)")
r_mlin = c.fila("Línea automática 2019", round(MOD_LINEA, 4), fmt=EU3, fuente=TOT,
                nota="38.000 € × 2 operarios ÷ 94.380 palets")
c.hueco()
c.cabecera(["Rendimiento del robot", "Robot (ud)", "Resto (ud)", "MOD ponderada", "Comentario"])
fila_pond = c.f
for i, rend in enumerate((0.90, 0.85, 0.80, 0.75)):
    f = c.f
    ws = c.ws
    fr = fila_rend + i
    cc = ws.cell(f, 1, f"=A{fr}"); cc.font = NEG; cc.number_format = PCT
    cc = ws.cell(f, 2, f"=MIN(C{fr},90000)"); cc.font = NEG; cc.number_format = NUM
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 3, f"=90000-B{f}"); cc.font = NEG; cc.number_format = NUM
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 4, f"=(38000+C{f}*$B${r_mlin})/90000")
    cc.font = TOT; cc.number_format = EU3
    cc.alignment = Alignment(horizontal="center")
    cc.fill = F_CLAVE
    ws.cell(f, 5, "El robot cuesta 38.000 € (1 operario) haga las unidades que haga; "
                  "el resto se paga a 0,805 €/palet").font = MINI
    ws.cell(f, 5).alignment = Alignment(wrap_text=True, vertical="center")
    c.f += 1
c.hueco()

c.cabecera(["Coste unitario real de Persán", "MOD ponderada", "Coste total", "Margen a 10 €",
            "Impacto anual"])
for i, rend in enumerate((0.90, 0.85, 0.80, 0.75)):
    f = c.f
    ws = c.ws
    fp = fila_pond + i
    cc = ws.cell(f, 1, f"=A{fp}"); cc.font = NEG; cc.number_format = PCT
    cc = ws.cell(f, 2, f"=D{fp}"); cc.font = NEG; cc.number_format = EU3
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 3, f"=8.27+B{f}+0.53+1.3"); cc.font = TOT; cc.number_format = EU3
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 4, f"=10-C{f}"); cc.font = TOT; cc.number_format = EU3
    cc.alignment = Alignment(horizontal="center")
    cc.fill = F_MAL
    cc = ws.cell(f, 5, f"=D{f}*90000"); cc.font = TOT; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    cc.fill = F_MAL
    c.f += 1
c.explicar("El Anexo 3 dice que Persán cuesta 10,52 € y pierde 0,52 € por palet. Con el coste "
           "de mano de obra correctamente ponderado, incluso en el escenario OFICIAL del 90% "
           "el coste es 10,548 € y la pérdida 0,548 €. En el escenario realista del 80%, la "
           "pérdida sube a 0,635 € por palet. El contrato es peor de lo que el propio caso "
           "presenta.", F_MAL)
c.hueco()

c.barra("¿CABE PERSÁN EN LA FÁBRICA?")
c.cabecera(["Escenario", "Capacidad", "Demanda total", "Holgura", "¿Encaja?"])
for et, cap, dem in [
    ("Sin robot, sin Persán", f"=D{r_cap}", f"=B{r_prod}"),
    ("Sin robot, CON Persán", f"=D{r_cap}", f"=B{r_prod}+90000"),
    ("Con robot al 90%, CON Persán", f"=D{r_cap}+C{fila_rend}", f"=B{r_prod}+90000"),
    ("Con robot al 80%, CON Persán", f"=D{r_cap}+C{fila_rend+2}", f"=B{r_prod}+90000"),
]:
    f = c.f
    ws = c.ws
    ws.cell(f, 1, et).font = ETI
    cc = ws.cell(f, 2, cap); cc.font = NEG; cc.number_format = NUM
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 3, dem); cc.font = NEG; cc.number_format = NUM
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 4, f"=B{f}-C{f}"); cc.font = TOT; cc.number_format = NUM
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 5, f'=IF(D{f}>=0,"SÍ","NO")'); cc.font = TOT
    cc.alignment = Alignment(horizontal="center")
    c.f += 1
c.hueco()

c.barra("LA OPORTUNIDAD QUE EL CASO NO SEÑALA", "→ VA AL INFORME")
c.explicar("La línea de palet usado está limitada a 30.000 unidades al mes por FALTA DE "
           "ESPACIO, no por falta de demanda. Y la campa que habría que alquilar para el "
           "retén de Persán cuesta 5.000 €/mes. Es decir: el mismo alquiler que se plantea "
           "para un contrato que pierde dinero desbloquearía una línea que gana dinero.")
c.cabecera(["Concepto", "Valor", "", "", "Comentario"])
r_usa = c.fila("Producción actual de palet usado", 260330, fmt=NUM, fuente=AZUL, nota="Anexo 2")
r_tec = c.fila("Techo logístico", "=30000*12", fmt=NUM, nota="30.000/mes por falta de espacio")
r_rec = c.fila("RECORRIDO DISPONIBLE", f"=B{r_tec}-B{r_usa}", fmt=NUM, fuente=TOT,
               relleno=F_BIEN)
c.fila("Contribución adicional (sin mano de obra)", f"=B{r_rec}*1.45", fmt=EUR, fuente=TOT,
       relleno=F_BIEN, nota="A 1,45 €/palet. Cubre con holgura los 60.000 € de la campa")
c.fila("Operarios necesarios para llegar al techo", f"=B{r_tec}/(242*155)", fmt=NU2,
       nota="Hoy hay 7. Harían falta unos 2,6 más, que el robot puede liberar")

# ══════════════════════════════════════════════════════════════════════════
# 5. CONTRATO PERSÁN
# ══════════════════════════════════════════════════════════════════════════
p = Hoja(wb, "5. Contrato Persán", "5. Qué hace falta para servir a Persán y cuánto cuesta",
         "90.000 unidades/año · 10 €/unidad FIJO durante 5 años · cobro a 180 días · "
         "retén permanente de 1.000 palets en nave.")

p.barra("QUÉ SE NECESITARÍA SI SE FIRMA", "→ VA AL INFORME")
p.explicar("Antes de mirar el resultado, conviene listar todo lo que el contrato obliga a "
           "poner encima de la mesa. No es solo la máquina.")
p.cabecera(["Necesidad", "Importe / cantidad", "", "", "Detalle"])
p.fila("1. Capacidad productiva", 90000, fmt=NUM, fuente=TOT,
       nota="El robot cubre 87.120 al 90% (77.440 al 80%). El resto hay que completarlo con "
            "la línea de 2019, que tiene holgura.")
p.fila("2. Personal", 1, fmt=NUM, fuente=TOT,
       nota="Un operario nuevo para la máquina (38.000 €/año). NO hacen falta despidos: la "
            "capacidad extra se necesita igualmente.")
p.fila("3. Espacio: campa", 60000, fmt=EUR, fuente=TOT,
       nota="5.000 €/mes para almacenar el retén de 1.000 palets que exige el cliente.")
p.fila("4. Logística: tercer camión", 150000, fmt=EUR, fuente=TOT,
       nota="Los 2 actuales están al 100%. Solo hace falta por el retén de Persán.")
p.fila("5. Leasing (máquina + camión)", 48000, fmt=EUR, fuente=TOT,
       nota="4.000 €/mes durante 5 años. Si se compra SOLO la máquina, baja a 9.618 €/año.")
p.fila("6. Circulante a financiar", 454356, fmt=EUR, fuente=TOT, relleno=F_MAL,
       nota="443.836 € de cobro a 180 días + 10.520 € del retén inmovilizado.")
p.fila("7. Compra adicional de madera", 744300, fmt=EUR, fuente=TOT, relleno=F_MAL,
       nota="90.000 × 8,27 €. Supone un 52% MÁS de compra de madera de palet nuevo, a un "
            "único proveedor gallego. Riesgo de suministro que el caso no menciona.")
p.fila("8. Caja necesaria el primer año", -643232, fmt=EUR, fuente=BIG, relleno=F_MAL,
       nota="Frente a una caja operativa de −60.555 € en 2025.")
p.hueco()

p.barra("COSTES FIJOS ADICIONALES QUE GENERA EL CONTRATO")
p.cabecera(["Concepto", "€/año", "", "", "Comentario"])
r_campa = p.fila("Alquiler de la campa", "=5000*12", fmt=EUR, nota="5.000 €/mes")
r_leas = p.fila("Leasing de máquina y camión", "=4000*12", fmt=EUR, nota="4.000 €/mes")
r_cfad = p.fila("TOTAL COSTES FIJOS ADICIONALES", f"=B{r_campa}+B{r_leas}", fmt=EUR,
                fuente=TOT, relleno=F_TOT,
                nota="Sube los costes fijos de la empresa un 13,6%")
p.hueco()

p.barra("LA CUENTA DE RESULTADOS DEL CONTRATO", "→ VA AL INFORME")
p.explicar("Lo que aporta el contrato al resultado de la empresa, año a año. Se usa el coste "
           "unitario REAL de 10,548 € (mano de obra ponderada, hoja 4), no los 10,52 € del "
           "Anexo 3.")
p.cabecera(["Concepto", "Importe anual", "", "", "Comentario"])
r_vol = p.fila("Volumen (palets)", 90000, fmt=NUM, fuente=AZUL)
r_pre = p.fila("Precio unitario", 10.00, fmt=EU2, fuente=AZUL)
r_cu = p.fila("Coste unitario real", 10.548, fmt=EU3, fuente=NEG,
              nota="Con mano de obra ponderada al 90% de rendimiento. Ver hoja 4")
p.fila("Ingresos", f"=B{r_vol}*B{r_pre}", fmt=EUR)
p.fila("(−) Costes variables", f"=-B{r_vol}*B{r_cu}", fmt=EUR)
r_mcp = p.fila("MARGEN DE CONTRIBUCIÓN", f"=B{p.f-2}+B{p.f-1}", fmt=EUR, fuente=TOT,
               relleno=F_MAL, nota="Negativo antes de cualquier coste fijo")
p.fila("(−) Costes fijos adicionales", f"=-B{r_cfad}", fmt=EUR, nota="Campa + leasing")
r_sub = p.fila("SUBTOTAL", f"=B{r_mcp}+B{p.f-1}", fmt=EUR, fuente=TOT, relleno=F_TOT)
p.hueco()

p.barra("EL COSTE OCULTO: COBRAR A 180 DÍAS")
p.explicar("Persán paga a seis meses. Alcopalet tendría permanentemente medio año de "
           "facturación de ese cliente sin cobrar, y ese dinero hay que financiarlo al 7,5%. "
           "El escandallo no lo recoge, pero es un coste real.")
p.cabecera(["Concepto", "Importe", "", "", "Comentario"])
r_cxc = p.fila("Pendiente de cobro (media)", f"=B{r_vol}*B{r_pre}*180/365", fmt=EUR,
               nota="Ingresos × 180/365")
r_stk = p.fila("Retén inmovilizado, a coste", f"=1000*B{r_cu}", fmt=EUR,
               nota="1.000 palets siempre en nave")
r_cir = p.fila("CIRCULANTE TOTAL", f"=B{r_cxc}+B{r_stk}", fmt=EUR, fuente=TOT, relleno=F_TOT)
r_cfin = p.fila("Coste financiero anual", f"=-B{r_cir}*0.075", fmt=EUR, fuente=TOT,
                relleno=F_MAL, nota="Al 7,5%")
p.hueco()

p.barra("IMPACTO ANUAL", "→ VA AL INFORME")
p.cabecera(["Concepto", "Importe", "", "", "Comentario"])
r_imp = p.fila("IMPACTO ANUAL DEL CONTRATO", f"=B{r_sub}+B{r_cfin}", fmt=EUR, fuente=TOT,
               relleno=F_MAL, grande=BIG, nota="Lo que resta al resultado cada año")
p.fila("Resultado neto actual de la empresa", 279364, fmt=EUR, fuente=AZUL, nota="Anexo 1")
p.fila("En % del resultado actual", f"=B{r_imp}/B{p.f-1}", fmt=PCT, fuente=TOT, relleno=F_MAL,
       nota="Se lleva por delante dos tercios del beneficio de TODA la empresa")
p.fila("Impacto acumulado a 5 años", f"=B{r_imp}*5", fmt=EUR, fuente=TOT, relleno=F_MAL,
       nota="El contrato dura 5 años sin revisión de precio")
p.fila("Resultado de la empresa tras firmar", f"=279364+B{r_imp}", fmt=EUR, fuente=TOT,
       nota="La empresa quedaría al borde de las pérdidas")
p.hueco()

p.barra("PRECIO DE EQUILIBRIO", "→ VA AL INFORME")
p.explicar("Si el problema es el precio, la pregunta útil es «¿a partir de qué precio "
           "interesa?». El tercer nivel es el importante: el precio que hace que el contrato "
           "no reste ni sume.")
p.cabecera(["Nivel de cobertura", "Precio mínimo", "Subida necesaria", "", "Qué cubre"])
p.fila("1. Solo costes variables", f"=B{r_cu}", f"=B{p.f}/B{r_pre}-1", fmt=EU2,
       fmts=[EU2, PCT], nota="Madera, mano de obra, transporte y variables")
r_be2 = p.fila("2. + campa y leasing", f"=B{r_cu}+B{r_cfad}/B{r_vol}",
               f"=B{p.f}/B{r_pre}-1", fmt=EU2, fmts=[EU2, PCT],
               nota="Añade los costes fijos que obliga a asumir")
p.fila("3. + coste del circulante",
       f"=(B{r_vol}*B{r_cu}+B{r_cfad})/(B{r_vol}*(1-180/365*0.075))",
       f"=B{p.f}/B{r_pre}-1", fmt=EU2, fmts=[EU2, PCT], fuente=TOT, relleno=F_CLAVE,
       grande=BIG, nota="PRECIO DE EQUILIBRIO REAL")
p.fila("4. + margen igual al negocio actual", f"=B{r_be2}+2.79", f"=B{p.f}/B{r_pre}-1",
       fmt=EU2, fmts=[EU2, PCT], nota="Para que Persán sea tan rentable como el resto")
p.hueco()

p.barra("SENSIBILIDAD: PRECIO FRENTE A COSTE DE LA MADERA")
p.explicar("Las dos variables que de verdad mueven el resultado. Cada celda es el impacto "
           "anual del contrato en euros.")
ws = p.ws
f = p.f
ws.cell(f, 1, "Madera ↓ / Precio →").font = CAB
ws.cell(f, 1).fill = F_CAB
precios = [10.00, 11.00, 12.00, 12.50]
for i, pr in enumerate(precios):
    cc = ws.cell(f, 2 + i, pr)
    cc.font = CAB; cc.fill = F_CAB; cc.number_format = EU2
    cc.alignment = Alignment(horizontal="center"); cc.border = BOR
fila_prec = f
p.f += 1
NO_MADERA = 0.448 + 0.53 + 1.30
for mad in (8.27, 7.75, 7.50, 7.00):
    f = p.f
    cc = ws.cell(f, 1, mad)
    cc.font = CAB; cc.fill = F_CAB; cc.number_format = EU2
    cc.alignment = Alignment(horizontal="center")
    for i in range(len(precios)):
        L = chr(66 + i)
        formula = (f"=90000*({L}${fila_prec}-($A{f}+{NO_MADERA}))-$B${r_cfad}"
                   f"-({L}${fila_prec}*90000*180/365+1000*($A{f}+{NO_MADERA}))*0.075")
        cc = ws.cell(f, 2 + i, formula)
        cc.font = NEG; cc.number_format = EUR; cc.border = BOR
        cc.alignment = Alignment(horizontal="center")
    p.f += 1
p.explicar("Lectura: con el precio actual de 10 € el contrato pierde dinero aunque la madera "
           "baje a 7 €. Y con la madera actual de 8,27 € hace falta llegar a unos 12,17 € "
           "solo para no perder. La combinación de las dos palancas es lo que hace viable el "
           "contrato.")

# ══════════════════════════════════════════════════════════════════════════
# 6. CONTRATO IDEAL
# ══════════════════════════════════════════════════════════════════════════
i = Hoja(wb, "6. Contrato ideal", "6. Qué habría que negociar con Persán",
         "La pregunta correcta no es «¿firmo o no?» sino «¿en qué condiciones sí?». "
         "Hay tres palancas, y ninguna por separado basta.")

i.barra("LAS TRES PALANCAS")
i.explicar("Casi todo el mundo va a discutir el precio. El precio es solo una de tres, y "
           "probablemente la más difícil. La segunda —la especificación de la madera— es la "
           "que el caso deja servida sin decirlo.")
i.cabecera(["Palanca", "Hoy", "Objetivo", "Ganancia anual", "Por qué es negociable"])
i.fila("1. Precio por palet", 10.00, 12.50, 213172, fmt=EU2, fmts=[EU2, EU2, EUR],
       fuente=TOT, nota="90.000 × 2,50 € menos el mayor coste de circulante. Es la palanca "
                        "más difícil: Persán ha negociado durante meses.")
i.fila("2. Coste de la madera", 8.27, 7.50, 69300, fmt=EU2, fmts=[EU2, EU2, EUR],
       fuente=TOT, nota="LA QUE NADIE MIRA. Los 8,27 € son el estándar Tetra Pak: madera "
                        "gallega premium, seca, blanca, lijada y sin astillas, nacido de un "
                        "incidente de moho en una aduana china. Persán fabrica detergentes "
                        "en Sevilla y sus palets no cruzan aduanas asiáticas.")
i.fila("3. Plazo de cobro (días)", 180, 90, 16638, fmt=NUM, fmts=[NUM, NUM, EUR],
       fuente=TOT, nota="La más fácil de conseguir. Reduce a la mitad el circulante "
                        "inmovilizado y su coste financiero.")
i.hueco()

i.barra("CLÁUSULAS QUE NO SON NEGOCIABLES")
i.cabecera(["Cláusula", "", "", "", "Por qué"])
i.fila("Revisión anual del precio ligada al coste de la madera", relleno=F_CLAVE,
       nota="INNEGOCIABLE. La madera es el 83% del precio de venta y el contrato dura 5 años. "
            "Sin esta cláusula, Alcopalet asume todo el riesgo de materia prima. Una subida "
            "del 10% cuesta 74.430 €/año.")
i.fila("Volumen mínimo garantizado con penalización", relleno=F_CLAVE,
       nota="Si Alcopalet invierte 190.000 € y alquila una campa por el contrato, necesita "
            "certeza de que las 90.000 unidades se van a pedir de verdad.")
i.fila("Reducir el retén de 1.000 a 500 palets, o repercutir su coste", relleno=F_CLAVE,
       nota="El retén es stock inmovilizado que financia Alcopalet para dar un servicio "
            "premium. O se reduce, o se cobra.")
i.fila("Duración de 3 años en vez de 5, con revisión", relleno=F_CLAVE,
       nota="Menos exposición a un precio que puede quedar obsoleto. Y da una segunda "
            "oportunidad de negociación.")
i.hueco()

i.barra("QUÉ SE CONSIGUE CON CADA COMBINACIÓN", "→ VA AL INFORME")
i.explicar("El impacto anual del contrato según lo que se logre en la negociación. Es la "
           "tabla que hay que llevar a la reunión con Persán.")
i.cabecera(["Escenario de negociación", "Precio", "Madera", "Cobro (días)", "Impacto anual"])
COSTE_NM = 0.448 + 0.53 + 1.30  # MOD ponderada + transporte + otros variables
for et, pr, mad, dias, relleno in [
    ("Contrato actual, tal como está", 10.00, 8.27, 180, F_MAL),
    ("Solo se consigue el plazo de cobro", 10.00, 8.27, 90, F_MAL),
    ("Solo se consigue re-especificar la madera", 10.00, 7.50, 180, F_MAL),
    ("MÍNIMO ACEPTABLE (equilibrio)", 12.17, 8.27, 180, F_CLAVE),
    ("Objetivo realista: precio y plazo", 12.50, 8.27, 90, F_BIEN),
    ("CONTRATO IDEAL: las tres palancas", 12.50, 7.50, 90, F_BIEN),
    ("Ambicioso: las tres al máximo", 13.00, 7.00, 90, F_BIEN),
]:
    f = i.f
    ws = i.ws
    ws.cell(f, 1, et).font = TOT if et.isupper() or "IDEAL" in et else ETI
    cc = ws.cell(f, 2, pr); cc.font = NEG; cc.number_format = EU2
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 3, mad); cc.font = NEG; cc.number_format = EU2
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 4, dias); cc.font = NEG; cc.number_format = NUM
    cc.alignment = Alignment(horizontal="center")
    formula = (f"=90000*(B{f}-(C{f}+{COSTE_NM}))-108000"
               f"-(B{f}*90000*D{f}/365+1000*(C{f}+{COSTE_NM}))*0.075")
    cc = ws.cell(f, 5, formula); cc.font = TOT; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    for col in range(1, 6):
        ws.cell(f, col).fill = relleno
    i.f += 1
i.explicar("Conclusión para la negociación: ninguna palanca por separado salva el contrato. "
           "Con solo el plazo de cobro sigue perdiendo. Con solo la madera, también. Hace "
           "falta subir el precio SÍ O SÍ, pero combinarlo con las otras dos reduce lo que "
           "hay que pedir: en vez de exigir 12,17 € a secas, se puede cerrar en 12,50 € con "
           "un palet más barato de fabricar y cobrando antes.")
i.hueco()

i.barra("Y SI PERSÁN DICE QUE NO")
for t in [
    "No pasa nada. Perder un contrato que pierde dinero no es perder: la capacidad ociosa de "
    "40.876 palets/año sigue ahí, y el robot se justifica igualmente por flexibilidad.",
    "El coste de equivocarse es asimétrico: rechazar cuesta una oportunidad; firmar a 10 € "
    "cuesta 944.383 € en cinco años.",
    "Hay que hablar ANTES con el proveedor gallego de madera. Un aumento del 52% en el "
    "volumen de compra da poder de negociación, y si la madera baja, todas las cuentas "
    "mejoran, se firme o no con Persán.",
]:
    i.texto(t)

# ══════════════════════════════════════════════════════════════════════════
# 7. INVERSIÓN MÁQUINA
# ══════════════════════════════════════════════════════════════════════════
v = Hoja(wb, "7. Inversión máquina", "7. Análisis de la inversión de la máquina",
         "40.000 € a 5 años, con un coste de financiación del 7,5%. La rentabilidad no "
         "depende de la máquina: depende de a qué se dedique.")

v.barra("DATOS DE LA INVERSIÓN")
v.cabecera(["Concepto", "Valor", "", "", "Comentario"])
r_inv = v.fila("Inversión", 40000, fmt=EUR, fuente=AZUL, nota="Precio del robot")
r_k = v.fila("Coste del dinero", 0.075, fmt=PCT, fuente=AZUL, nota="Tipo ofrecido por el banco")
r_n = v.fila("Horizonte (años)", 5, fmt=NUM, fuente=AZUL, nota="Igual que la duración del leasing")
r_cuota = v.fila("Cuota anual del leasing, solo la máquina",
                 f"=B{r_inv}*(B{r_k}/12)/(1-(1+B{r_k}/12)^(-B{r_n}*12))*12", fmt=EUR,
                 fuente=TOT, nota="Frente a 48.000 €/año si se financia junto con el tráiler. "
                                  "Separar las dos compras ahorra 38.382 €/año")
v.hueco()

v.barra("LA RENTABILIDAD DEPENDE DEL USO", "→ VA AL INFORME")
v.explicar("La misma máquina, con el mismo coste, da resultados radicalmente distintos según "
           "a qué se destine. Esta tabla es el argumento para separar la decisión de comprar "
           "de la decisión de firmar.")
v.cabecera(["Uso que se le dé", "Flujo anual", "VAN a 5 años", "Plazo de recuperación", "Comentario"])
for et, flujo, nota, relleno in [
    ("A. Sustituye mano de obra y esta se reasigna al 100%", 180880,
     "Los 4,76 operarios liberados pasan a producir algo vendible (palet usado, crecimiento).",
     F_BIEN),
    ("B. Se reasigna la mitad", 90440,
     "Escenario intermedio y probablemente el más realista.", F_BIEN),
    ("C. No se reasigna a nadie", 0,
     "Si nadie cambia de puesto y no se vende más, el ahorro teórico no existe.", F_MAL),
    ("D. Se dedica a Persán a 10 €", -45302,
     "Fabricar a pérdida con una máquina eficiente sigue siendo fabricar a pérdida.", F_MAL),
    ("E. Se dedica al mercado actual a 15,16 €", 396396,
     "El mejor uso posible, PERO requiere demanda: hoy la fábrica está al 81%.", F_BIEN),
]:
    f = v.f
    ws = v.ws
    ws.cell(f, 1, et).font = ETI
    cc = ws.cell(f, 2, flujo); cc.font = NEG; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 3, f"=-$B${r_inv}+B{f}*(1-(1+$B${r_k})^-$B${r_n})/$B${r_k}")
    cc.font = TOT; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 4, f'=IF(B{f}<=0,"nunca",$B${r_inv}/B{f}*12)')
    cc.font = TOT; cc.number_format = '0.0" meses"'
    cc.alignment = Alignment(horizontal="center")
    ws.cell(f, 5, nota).font = MINI
    ws.cell(f, 5).alignment = Alignment(wrap_text=True, vertical="center")
    for col in range(1, 6):
        ws.cell(f, col).fill = relleno
    v.f += 1
v.explicar("Lectura: el caso C es el escenario más conservador y el único en que la máquina "
           "no se paga sola, y aun así solo cuesta 40.000 €. Los casos A, B y E la convierten "
           "en la mejor inversión disponible. El caso D —dedicarla a Persán tal como está "
           "negociado— es el ÚNICO que destruye valor de forma estructural.")
v.hueco()

v.barra("POR QUÉ COMPRARLA IGUALMENTE, AUNQUE NO SE FIRME PERSÁN")
v.cabecera(["Razón", "Dato", "", "", "Explicación"])
v.fila("Flexibilidad de formato", 0.42, fmt=NU2, fuente=TOT,
       nota="Horas de cambio de formato, frente a 6 horas de la línea de 2019. Alcopalet gana "
            "dinero en series especiales: el 114×114 de Don Simón, las exigencias de Tetra "
            "Pak. Con 6 horas de parada una serie corta es inviable.")
v.fila("Productividad por operario", 360, fmt=NUM, fuente=TOT,
       nota="Palets por operario y día, frente a 62,5 en manual y 195 en la línea de 2019.")
v.fila("Reduce el riesgo laboral y de absentismo", 5.76, fmt=NU2, fuente=TOT,
       nota="Operarios manuales equivalentes. El caso dice que la mano de obra es «cada vez "
            "más difícil de gestionar y contratar» y menciona absentismo los lunes.")
v.fila("Mantenimiento mínimo", "", nota="Sin electrónica compleja ni rodamientos: funciona con "
                                       "poleas y solo requiere limpieza diaria. La línea de "
                                       "2019 tiene mantenimiento «complejo y constante».")
v.fila("Riesgo económico limitado", 40000, fmt=EUR, fuente=TOT, relleno=F_BIEN,
       nota="Sobre un beneficio de 279.364 €. En el peor escenario se pierden 40.000 €; en el "
            "mejor se ganan 180.000 € al año.")

# ══════════════════════════════════════════════════════════════════════════
# 8. ESCENARIOS
# ══════════════════════════════════════════════════════════════════════════
s = Hoja(wb, "8. Escenarios", "8. Las cinco opciones, comparadas",
         "Todo lo anterior confluye aquí. Cada columna es una alternativa real que el caso "
         "pone sobre la mesa.")

s.barra("QUÉ ES CADA OPCIÓN")
s.cabecera(["Opción", "", "", "", "En qué consiste"])
for et, desc in [
    ("A — No hacer nada", "Ni máquina, ni contrato, ni dividendo. La empresa sigue igual y se "
                          "queda estancada en los 3.000 m² actuales."),
    ("B — Persán + robot (lo que propone Esteban)",
     "Firmar el contrato a 10 € y comprar máquina y tráiler con leasing conjunto."),
    ("C — Dividendo sin invertir (lo que piden los padres)",
     "Repartir beneficio para la casa y no comprometer nada."),
    ("D — Robot solo, sin Persán",
     "Comprar la máquina SIN el tráiler y no firmar. Se financia solo 40.000 €."),
    ("E — Robot + Persán renegociado (RECOMENDADA)",
     "Comprar la máquina hoy y firmar solo si se consiguen las tres palancas: 12,50 €, "
     "madera a 7,50 € y cobro a 90 días."),
]:
    s.fila(et, nota=desc)
s.hueco()

s.barra("IMPACTO ECONÓMICO ANUAL", "→ VA AL INFORME")
s.explicar("Efecto sobre el resultado anual, partiendo del beneficio actual de 279.364 €. "
           "Las opciones A y C no alteran el resultado: A no hace nada y C solo saca caja.")
s.cabecera(["Concepto", "B — Persán + robot", "D — Robot solo", "E — Renegociado", "Comentario"])
ws = s.ws
filas = {}
for clave, et, fB, fD, fE, nota in [
    ("mc", "Margen de contribución del contrato",
     "=90000*(10-10.548)", "0", f"=90000*(12.5-(7.5+{COSTE_NM}))",
     "En D no hay contrato. En E, con madera re-especificada"),
    ("ah", "Ahorro real de mano de obra", "0", "0", "0",
     "Cero en todas: hipótesis conservadora de no reasignar a nadie"),
    ("ca", "Alquiler de la campa", "=-5000*12", "0", "=-5000*12", None),
    ("le", "Leasing", "=-4000*12",
     "=-40000*(0.075/12)/(1-(1+0.075/12)^(-60))*12", "=-4000*12",
     "En D solo se financia la máquina: no hace falta tráiler"),
    ("ci", "Coste financiero del circulante",
     "=-(90000*10*180/365+1000*10.548)*0.075", "0",
     f"=-(90000*12.5*90/365+1000*(7.5+{COSTE_NM}))*0.075",
     "En E se cobra a 90 días en vez de 180"),
]:
    f = s.f
    ws.cell(f, 1, et).font = ETI
    for col, formula in ((2, fB), (3, fD), (4, fE)):
        cc = ws.cell(f, col, formula)
        cc.font = NEG; cc.number_format = EUR
        cc.alignment = Alignment(horizontal="center")
    if nota:
        ws.cell(f, 5, nota).font = MINI
        ws.cell(f, 5).alignment = Alignment(wrap_text=True, vertical="center")
    filas[clave] = f
    s.f += 1

f = s.f
ws.cell(f, 1, "IMPACTO ANUAL SOBRE EL RESULTADO").font = TOT
for col in (2, 3, 4):
    L = chr(64 + col)
    cc = ws.cell(f, col, f"=SUM({L}{filas['mc']}:{L}{filas['ci']})")
    cc.font = TOT; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
for col in range(1, 6):
    ws.cell(f, col).fill = F_TOT
fila_imp = f
s.f += 1

f = s.f
ws.cell(f, 1, "RESULTADO NETO PROYECTADO").font = TOT
for col in (2, 3, 4):
    L = chr(64 + col)
    cc = ws.cell(f, col, f"=279364+{L}{fila_imp}")
    cc.font = TOT; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
for col in range(1, 6):
    ws.cell(f, col).fill = F_TOT
ws.cell(f, 5, "Opciones A y C: 279.364 € (sin cambio)").font = MINI
s.f += 1
s.hueco()

s.barra("IMPACTO EN LA CAJA DEL PRIMER AÑO", "→ VA AL INFORME")
s.explicar("Tan importante como el resultado. La empresa generó caja NEGATIVA en 2025, así "
           "que cualquier decisión tiene que caber en la tesorería.")
s.cabecera(["Concepto", "B — Persán + robot", "D — Robot solo", "E — Renegociado", "Comentario"])
f = s.f
ws.cell(f, 1, "Impacto en el resultado").font = ETI
for col in (2, 3, 4):
    L = chr(64 + col)
    cc = ws.cell(f, col, f"={L}{fila_imp}")
    cc.font = NEG; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
s.f += 1
f = s.f
ws.cell(f, 1, "(−) Circulante inmovilizado el primer año").font = ETI
for col, formula in ((2, "=-(90000*10*180/365+1000*10.548)"), (3, "0"),
                     (4, f"=-(90000*12.5*90/365+1000*(7.5+{COSTE_NM}))")):
    cc = ws.cell(f, col, formula)
    cc.font = NEG; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
s.f += 1
f = s.f
ws.cell(f, 1, "NECESIDAD DE CAJA DEL PRIMER AÑO").font = TOT
for col in (2, 3, 4):
    L = chr(64 + col)
    cc = ws.cell(f, col, f"={L}{f-2}+{L}{f-1}")
    cc.font = TOT; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
for col in range(1, 6):
    ws.cell(f, col).fill = F_MAL
s.f += 1
s.fila("Caja operativa generada en 2025 (referencia)", -60555, fmt=EUR, fuente=AZUL,
       relleno=F_MAL, nota="Ver hoja 2. La empresa NO generó caja el año pasado")
s.hueco()

s.barra("COMPARACIÓN CUALITATIVA")
s.explicar("No todo se decide con euros. Estos son los criterios que el propio caso pone "
           "sobre la mesa y que hay que respetar para que la propuesta sea realista.")
s.cabecera(["Criterio", "B — Persán + robot", "D — Robot solo", "E — Renegociado", "Observación"])
for crit, vB, vD, vE, nota in [
    ("¿El contrato cubre sus costes?", "NO", "n/a", "SÍ",
     "Solo E supera el precio de equilibrio de 12,17 €"),
    ("¿Cabe en la tesorería?", "NO", "SÍ", "AJUSTADO",
     "B exige 643.000 € el primer año"),
    ("¿Respeta «sin despidos»?", "SÍ", "SÍ", "SÍ",
     "Ninguna obliga a despedir: la capacidad extra se necesita igual"),
    ("¿Respeta el turno único?", "SÍ", "SÍ", "SÍ", None),
    ("¿Permite crecer hasta 2030?", "SÍ", "PARCIAL", "SÍ",
     "A y C congelan la empresa en los 3.000 m² actuales"),
    ("¿Da seguridad a los fundadores?", "NO", "PARCIAL", "PARCIAL",
     "B deja el beneficio en 90.000 €: sin dividendo en 5 años"),
    ("¿Reduce la dependencia de mano de obra?", "SÍ", "SÍ", "SÍ",
     "Recurso «cada vez más difícil de contratar» según el caso"),
    ("Riesgo de concentración de cliente", "ALTO", "BAJO", "MEDIO",
     "Persán sería el 20% de la facturación a precio fijo"),
]:
    f = s.f
    ws.cell(f, 1, crit).font = ETI
    for col, val in ((2, vB), (3, vD), (4, vE)):
        cc = ws.cell(f, col, val)
        cc.font = TOT
        cc.alignment = Alignment(horizontal="center")
        cc.fill = F_MAL if val in ("NO", "ALTO") else (
            F_BIEN if val in ("SÍ", "BAJO") else F_CLAVE)
    if nota:
        ws.cell(f, 5, nota).font = MINI
        ws.cell(f, 5).alignment = Alignment(wrap_text=True, vertical="center")
    s.f += 1
s.hueco()

s.barra("LA DECISIÓN")
for t in [
    "La máquina y el contrato tienen plazos distintos: el fabricante pide respuesta HOY, "
    "mientras que el contrato de Persán todavía es un borrador. No hay ninguna razón para "
    "resolver las dos cosas en la misma reunión, y confundirlas es exactamente lo que le "
    "está pasando a Esteban.",
    "COMPRAR LA MÁQUINA, sin el tráiler: 40.000 € de riesgo, leasing de 9.618 €/año en vez de "
    "48.000 €, y se justifica por flexibilidad y por reducir la dependencia de mano de obra.",
    "NO FIRMAR a 10 €: pierde 0,55 € por palet, y hasta 0,10 € aunque el operario fuese "
    "gratis. Volver a Persán con las tres palancas y un objetivo claro de 12,50 €.",
    "A LOS PADRES: rechazar Persán a 10 € es precisamente lo que protege su jubilación. Si se "
    "firma, el beneficio cae a 90.000 € y no habrá dividendo en cinco años. Proponer un "
    "reparto moderado ahora y una política de dividendo creciente.",
]:
    s.texto(t)

# ══════════════════════════════════════════════════════════════════════════
# 9. ANEXOS PARA EL INFORME
# ══════════════════════════════════════════════════════════════════════════
a = Hoja(wb, "9. Anexos informe", "9. Los anexos que llevaría al informe final",
         "El Anexo al CPNNI-24 limita el informe a 10 carillas INCLUYENDO anexos, portada e "
         "índice. Con ~5 carillas de informe y 1 de portada e índice, quedan unas 4 para "
         "anexos. Estos son los cuatro que elegiría.")

a.barra("EL PRESUPUESTO DE ESPACIO")
a.cabecera(["Elemento", "Carillas", "", "", "Comentario"])
a.fila("Portada e índice", 1.0, fmt=NU2, fuente=TOT, nota="Título, autor, fecha, programa")
a.fila("Cuerpo del informe", 5.0, fmt=NU2, fuente=TOT,
       nota="Introducción, problema, opciones, criterios, elección y conclusión")
a.fila("Anexos", 3.5, fmt=NU2, fuente=TOT, nota="Los cuatro de abajo")
a.fila("TOTAL", 9.5, fmt=NU2, fuente=TOT, relleno=F_BIEN,
       nota="Por debajo del máximo de 10. El ideal recomendado son 7")
a.hueco()

a.barra("ANEXO 1 — ESTRUCTURA DE COSTES Y PUNTO DE EQUILIBRIO", "½ carilla")
a.explicar("Sale de la hoja 2. Es el anexo que sostiene TODO el análisis: sin la separación "
           "entre fijos y variables no se puede juzgar el contrato.")
a.cabecera(["Qué incluir", "", "", "", "Por qué"])
a.fila("Tabla de costes fijos y variables sobre ventas", nota="Con el % sobre ventas de cada "
       "partida y el margen de contribución del 27,2%")
a.fila("La línea de comprobación del EBITDA", nota="Diferencia cero con el Anexo 1. Es lo que "
       "hace la reordenación indiscutible")
a.fila("Punto de equilibrio y margen de seguridad", nota="3.214.115 € y 29,9%")
a.fila("Efecto de Persán sobre el punto de equilibrio", nota="Sube casi 400.000 € por los "
       "108.000 € de costes fijos adicionales")
a.hueco()

a.barra("ANEXO 2 — ANÁLISIS ECONÓMICO DEL CONTRATO PERSÁN", "1 carilla")
a.explicar("Sale de las hojas 4 y 5. Es el anexo central: contiene la prueba de que el "
           "contrato destruye valor y a partir de qué precio dejaría de hacerlo.")
a.cabecera(["Qué incluir", "", "", "", "Por qué"])
a.fila("Escandallo con la mano de obra ponderada", nota="Coste real de 10,548 € frente a los "
       "10,52 € del Anexo 3, y margen de −0,548 €/palet")
a.fila("Margen sin mano de obra: −0,10 €/palet", nota="EL ARGUMENTO DEFINITIVO: aunque el "
       "operario fuese gratis, el contrato pierde")
a.fila("Cuenta de resultados incremental completa", nota="Del margen de contribución al "
       "impacto anual de −191.398 €, pasando por campa, leasing y circulante")
a.fila("Los cuatro niveles de precio de equilibrio", nota="Y el resultado: 12,17 €")
a.fila("Tabla de sensibilidad precio × madera", nota="Demuestra que ninguna palanca por "
       "separado salva el contrato")
a.hueco()

a.barra("ANEXO 3 — CAPACIDAD PRODUCTIVA Y ENCAJE DEL CONTRATO", "½ carilla")
a.explicar("Sale de la hoja 4. Es el anexo que desmonta el argumento de urgencia de Esteban y "
           "que revela que la máquina no llega a las 90.000 unidades.")
a.cabecera(["Qué incluir", "", "", "", "Por qué"])
a.fila("Capacidad anual de los tres sistemas", nota="Línea 2019, manual y robot, con el "
       "cálculo a la vista")
a.fila("Grado de utilización actual: 81%", nota="Desmonta el «no llegamos». Hay 40.876 palets "
       "de holgura sin invertir")
a.fila("Capacidad del robot: oficial 90% y realista 80%", nota="En NINGÚN escenario cubre las "
       "90.000 por sí solo")
a.fila("Recorrido de la línea de palet usado", nota="99.670 unidades bloqueadas por espacio, "
       "no por demanda. Es la alternativa creativa")
a.hueco()

a.barra("ANEXO 4 — ESCENARIOS Y ANÁLISIS DE LA INVERSIÓN", "1 carilla")
a.explicar("Sale de las hojas 7 y 8. Es el anexo que respalda la recomendación y demuestra "
           "que se han valorado alternativas, que es el criterio 8 de evaluación "
           "(«creatividad para generar soluciones y alternativas»).")
a.cabecera(["Qué incluir", "", "", "", "Por qué"])
a.fila("Tabla comparativa de las cinco opciones", nota="Impacto anual, resultado proyectado y "
       "necesidad de caja del primer año")
a.fila("Comparación cualitativa frente a los criterios", nota="Sin despidos, turno único, "
       "crecimiento, seguridad de los fundadores, riesgo de cliente")
a.fila("VAN y plazo de recuperación de la máquina", nota="Los cinco usos posibles. Demuestra "
       "que la máquina es buena y el contrato malo, por separado")
a.fila("Términos propuestos para la renegociación", nota="Las tres palancas con su objetivo y "
       "las cláusulas innegociables")
a.hueco()

a.barra("QUÉ DEJARÍA FUERA Y POR QUÉ")
for t in [
    "La reconciliación entre el Anexo 1 y el Anexo 3 (los 288 € de desviación). Es "
    "imprescindible haberla hecho, pero es control de calidad interno. Se menciona en una "
    "frase dentro del cuerpo y se guarda para el turno de preguntas.",
    "El endeudamiento implícito (587.280 €, 1,30x EBITDA que pasaría a 4,13x). Es un buen "
    "argumento, pero cabe en dos líneas del cuerpo del informe. Al anexo solo iría si sobrara "
    "espacio.",
    "El coste por camión (134.109 € frente a los 47.700 € que asigna el Anexo 3). Es una "
    "inferencia propia, no un dato del caso. Mejor guardarla para la defensa oral, donde "
    "demuestra profundidad sin comprometer el rigor del informe escrito.",
    "El análisis del Anexo 2 (rentabilidad por líneas). Sus conclusiones caben en tres filas "
    "dentro del cuerpo del informe: no merece un anexo propio.",
    "Las tablas de sensibilidad completas al plazo de cobro. Basta con la de precio × madera, "
    "que es la que de verdad discrimina.",
]:
    a.texto(t)

wb.save("/home/user/DANISANTELMO/Alcopalet_Analisis_Completo.xlsx")
print("Guardado:", [x.title for x in wb.worksheets])
