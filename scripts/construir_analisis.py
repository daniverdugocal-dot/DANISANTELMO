"""
ALCOPALET (DTI-1399) — Análisis del caso.

Tres hojas:
    RESUMEN   la decisión y los ocho números que la sostienen
    CÁLCULOS  el detalle, en orden, cada bloque una pregunta
    ANEXOS    los cuatro anexos del informe final

Respeta las notas del Anexo 3: MOD de Persán = 38.000 € / 90.000 ud con un solo
operario (nota 1); transporte 0,53 € con el tráiler propio a plena carga (nota 2);
otros variables 1,30 € proporcionales al volumen (nota 3); y 11 meses × 22 días
= 242 días productivos al año (nota final).
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FU = "Arial"
TIT = Font(name=FU, size=16, bold=True, color="1F3864")
SUB = Font(name=FU, size=9, italic=True, color="595959")
SEC = Font(name=FU, size=11, bold=True, color="FFFFFF")
CAB = Font(name=FU, size=10, bold=True, color="FFFFFF")
ETI = Font(name=FU, size=10)
AZUL = Font(name=FU, size=10, color="0000FF")
NEG = Font(name=FU, size=10)
TOT = Font(name=FU, size=10, bold=True)
BIG = Font(name=FU, size=14, bold=True, color="C00000")
VER = Font(name=FU, size=14, bold=True, color="006100")
MINI = Font(name=FU, size=9, color="404040")

F_SEC = PatternFill("solid", fgColor="1F3864")
F_CAB = PatternFill("solid", fgColor="4472C4")
F_TOT = PatternFill("solid", fgColor="D9E2F3")
F_MAL = PatternFill("solid", fgColor="FCE4E4")
F_BIEN = PatternFill("solid", fgColor="E2EFDA")
F_CLA = PatternFill("solid", fgColor="FFF2CC")
F_HIP = PatternFill("solid", fgColor="FFFF00")
F_GRIS = PatternFill("solid", fgColor="F2F2F2")

_s = Side(style="thin", color="BFBFBF")
BOR = Border(left=_s, right=_s, top=_s, bottom=_s)

EUR = '#,##0 "€";(#,##0) "€";"-"'
EU2 = '#,##0.00 "€";(#,##0.00) "€";"-"'
PCT = '0.0%;(0.0%);"-"'
NUM = '#,##0;(#,##0);"-"'
NU2 = '#,##0.00;(#,##0.00);"-"'
ANCHOS = {"A": 46, "B": 15, "C": 14, "D": 14, "E": 46}


class H:
    def __init__(self, wb, nombre, titulo, sub):
        self.ws = wb.create_sheet(nombre)
        ws = self.ws
        ws.sheet_view.showGridLines = False
        for c, w in ANCHOS.items():
            ws.column_dimensions[c].width = w
        ws["A1"] = titulo
        ws["A1"].font = TIT
        ws["A2"] = sub
        ws["A2"].font = SUB
        ws.merge_cells("A2:E2")
        ws.row_dimensions[2].height = 12.5 * (len(sub) // 128 + 1) + 4
        self.f = 4
        ws.freeze_panes = "A4"

    def sec(self, texto, etiq=None):
        ws, f = self.ws, self.f
        ws.cell(f, 1, texto).font = SEC
        for c in range(1, 6):
            ws.cell(f, c).fill = F_SEC
        if etiq:
            c = ws.cell(f, 5, etiq)
            c.font = Font(name=FU, size=9, bold=True, color="FFFFFF")
            c.alignment = Alignment(horizontal="right")
        self.f += 1

    def cab(self, cols):
        ws, f = self.ws, self.f
        for i, t in enumerate(cols):
            c = ws.cell(f, 1 + i, t)
            c.font = CAB
            c.fill = F_CAB
            c.alignment = Alignment(horizontal="left" if i == 0 else "center",
                                    wrap_text=True, vertical="center")
            c.border = BOR
        ws.row_dimensions[f].height = 24
        self.f += 1

    def fila(self, et, *vals, fmt=EUR, fte=NEG, nota=None, fill=None, big=None, fmts=None):
        ws, f = self.ws, self.f
        ws.cell(f, 1, et).font = TOT if (fill or big) else ETI
        for i, v in enumerate(vals):
            if v is None or v == "":
                continue
            c = ws.cell(f, 2 + i, v)
            c.font = big if big else (TOT if fill else fte)
            c.number_format = fmts[i] if fmts else fmt
            c.alignment = Alignment(horizontal="center")
        if nota:
            n = ws.cell(f, 5, nota)
            n.font = MINI
            n.alignment = Alignment(wrap_text=True, vertical="center")
        if fill:
            for c in range(1, 6):
                ws.cell(f, c).fill = fill
        self.f += 1
        return f

    def txt(self, t, fill=None):
        ws, f = self.ws, self.f
        c = ws.cell(f, 1, t)
        c.font = TOT if fill in (F_BIEN, F_MAL, F_CLA) else ETI
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=5)
        if fill:
            for col in range(1, 6):
                ws.cell(f, col).fill = fill
        ws.row_dimensions[f].height = 12.5 * (len(t) // 128 + 1) + 6
        self.f += 1

    def gap(self, n=1):
        self.f += n


wb = Workbook()
wb.remove(wb.active)

# ══════════════════════════════════════════════════════════════ RESUMEN
r = H(wb, "RESUMEN", "ALCOPALET — La decisión",
      "Caso DTI-1399 · Lydes 2026 · Azul = dato del caso · Negro = calculado · "
      "Amarillo = hipótesis propia")

r.sec("QUÉ HACER")
r.txt("1.  COMPRAR LA MÁQUINA HOY, PERO SIN EL TRÁILER.  El fabricante pide respuesta hoy; "
      "los 150.000 € del camión solo hacen falta si se firma Persán. Financiar solo los "
      "40.000 € baja el leasing de 48.000 € a 9.618 € al año.", F_BIEN)
r.txt("2.  NO FIRMAR PERSÁN A 10 €.  El contrato todavía es un borrador: no hay que decidirlo "
      "hoy. A ese precio pierde dinero en cada palet, y el daño a cinco años es de 944.000 €.",
      F_MAL)
r.txt("3.  VOLVER A NEGOCIAR CON TRES PALANCAS: precio, especificación de la madera y plazo de "
      "cobro. Con las tres, el contrato pasa de −188.877 € a +117.964 € al año.", F_CLA)
r.gap()

r.sec("LOS OCHO NÚMEROS QUE LO SOSTIENEN")
r.cab(["Número", "Valor", "", "", "Qué significa"])
r.fila("Margen de contribución de Persán", -0.52, fmt=EU2, big=BIG, fill=F_MAL,
       nota="€ por palet. NEGATIVO: no cubre ni lo que consume fabricarlo. Ningún volumen lo "
            "arregla, y ninguna máquina tampoco.")
r.fila("Margen sin contar la mano de obra", -0.10, fmt=EU2, big=BIG, fill=F_MAL,
       nota="EL ARGUMENTO DEFINITIVO: aunque el operario fuese gratis, seguiría perdiendo. "
            "8,27 + 0,53 + 1,30 = 10,10 € frente a un precio de 10 €.")
r.fila("Precio de equilibrio", 12.17, fmt=EU2, big=BIG, fill=F_CLA,
       nota="€ por palet. Por debajo de aquí el contrato destruye valor, con máquina o sin ella.")
r.fila("Impacto anual del contrato", -188877, fmt=EUR, big=BIG, fill=F_MAL,
       nota="Margen negativo + campa + leasing + coste del cobro a 180 días. Equivale al 68% "
            "del beneficio de toda la empresa.")
r.fila("Utilización actual de la fábrica", 0.810, fmt=PCT, big=BIG, fill=F_CLA,
       nota="La fábrica NO está llena: sobran 40.876 palets al año. El «no llegamos» de "
            "Esteban no se sostiene. La máquina se compra por flexibilidad, no por capacidad.")
r.fila("Caja operativa real de 2025", -60555, fmt=EUR, big=BIG, fill=F_MAL,
       nota="Frente a 279.364 € de beneficio contable. El beneficio está en el almacén, no en "
            "el banco: no hay caja ni para invertir ni para el dividendo.")
r.fila("Caja que exige Persán el primer año", -643232, fmt=EUR, big=BIG, fill=F_MAL,
       nota="Circulante del cobro a 180 días + campa + leasing + margen negativo.")
r.fila("Contrato renegociado con las tres palancas", 117964, fmt=EUR, big=VER, fill=F_BIEN,
       nota="12,50 € de precio, madera re-especificada a 7,50 € y cobro a 90 días.")
r.gap()

r.sec("EL CONTRATO IDEAL")
r.cab(["Variable", "Hoy", "Mínimo", "Objetivo", "Por qué"])
r.fila("Precio por palet", 10.00, 12.17, 12.50, fmt=EU2, fte=TOT,
       nota="El mínimo es el equilibrio exacto. Es la palanca más difícil: Persán lleva meses "
            "negociando este precio.")
r.fila("Coste de la madera", 8.27, 8.00, 7.50, fmt=EU2, fte=TOT, fill=F_CLA,
       nota="LA QUE NADIE MIRA. Los 8,27 € son el estándar Tetra Pak (madera gallega premium, "
            "seca, blanca, lijada), nacido de un incidente de moho en una aduana china. Persán "
            "fabrica detergentes en Sevilla: casi con seguridad no necesita esa calidad.")
r.fila("Plazo de cobro (días)", 180, 120, 90, fmt=NUM, fte=TOT,
       nota="La más fácil de conseguir. Reduce a la mitad el circulante inmovilizado.")
r.fila("Revisión anual del precio de la madera", "NO", "SÍ", "SÍ", fmt=NUM, fte=TOT,
       fmts=[NUM, NUM, NUM], fill=F_CLA,
       nota="INNEGOCIABLE. La madera es el 83% del precio de venta y el contrato dura 5 años.")
r.gap()
r.cab(["Lo que se consiga en la negociación", "Impacto anual", "", "", "Lectura"])
r.fila("Contrato actual (10 € · 8,27 € · 180 días)", -188877, fmt=EUR, fte=TOT, fill=F_MAL,
       nota="Destruye valor")
r.fila("Solo el plazo de cobro (10 € · 8,27 € · 90 d)", -172233, fmt=EUR, fte=TOT, fill=F_MAL,
       nota="Sigue perdiendo: ninguna palanca sola basta")
r.fila("Solo la madera (10 € · 7,50 € · 180 días)", -119519, fmt=EUR, fte=TOT, fill=F_MAL,
       nota="Tampoco basta")
r.fila("Mínimo aceptable (12,17 € · 8,27 € · 180 d)", -800, fmt=EUR, fte=TOT, fill=F_CLA,
       nota="Ni gana ni pierde. No merece la pena por sí solo")
r.fila("CONTRATO IDEAL (12,50 € · 7,50 € · 90 días)", 117964, fmt=EUR, big=VER, fill=F_BIEN,
       nota="Las tres juntas. Este es el objetivo de la reunión")
r.gap()

r.sec("SI PERSÁN DICE QUE NO")
r.txt("No pasa nada. La capacidad ociosa de 40.876 palets al año sigue ahí y la máquina se "
      "justifica igual por flexibilidad. El coste de equivocarse es asimétrico: rechazar "
      "cuesta una oportunidad, firmar cuesta 944.383 € en cinco años.")
r.txt("Y a los padres: rechazar Persán a 10 € es precisamente lo que protege su jubilación. "
      "Si se firma, el beneficio cae a 90.487 € y no habrá dividendo en cinco años.")

# ══════════════════════════════════════════════════════════════ CÁLCULOS
c = H(wb, "CÁLCULOS", "Los cálculos, en orden",
      "Cada bloque responde a una pregunta. Se lee de arriba abajo.")

c.sec("LOS DATOS QUE SE USAN")
c.cab(["Dato", "Valor", "", "", "De dónde sale"])
r_dias = c.fila("Días productivos al año", "=11*22", fmt=NUM, fte=TOT, fill=F_TOT,
                nota="NOTA FINAL DEL ANEXO 3: 11 meses × 22 días")
c.fila("Horas por jornada", 8, fmt=NUM, fte=AZUL, nota="Turno único de 7:00 a 15:00")
c.fila("Coste de un operario (€/año)", 38000, fmt=EUR, fte=AZUL, nota="Apartado (d) del caso")
r_int = c.fila("Tipo de interés del banco", 0.075, fmt=PCT, fte=AZUL, nota="Apartado (c)")
r_vol = c.fila("Volumen del contrato (palets/año)", 90000, fmt=NUM, fte=AZUL, nota="Borrador")
r_pre = c.fila("Precio de Persán (€/palet)", 10.00, fmt=EU2, fte=AZUL,
               nota="FIJO durante 5 años, sin cláusula de revisión")
r_cob = c.fila("Plazo de cobro (días)", 180, fmt=NUM, fte=AZUL)
r_ret = c.fila("Retén exigido (palets)", 1000, fmt=NUM, fte=AZUL,
               nota="Dos camiones en nave para servir en menos de 24 horas")
r_cam = c.fila("Alquiler de la campa (€/año)", "=5000*12", fmt=EUR,
               nota="5.000 €/mes, apartado (a)")
r_lea = c.fila("Leasing máquina + camión (€/año)", "=4000*12", fmt=EUR,
               nota="4.000 €/mes, apartado (c). Cubre robot (40.000 €) y tráiler (150.000 €)")
c.gap()

c.sec("1. ¿CUÁNTO DEJA CADA PALET?", "→ VA AL INFORME")
c.txt("Del Anexo 3. La última fila es la que decide: lo que queda tras pagar lo que ese palet "
      "consume al fabricarse.", F_GRIS)
c.cab(["Concepto (€/palet)", "Palet nuevo", "Palet usado", "PERSÁN", "Nota del anexo"])
fi = c.f
c.fila("Ingreso medio", 15.16, 7.45, 10.00, fmt=EU2, fte=AZUL,
       nota="Persán paga un 34% menos que el cliente medio de palet nuevo")
c.fila("Materia prima", 8.27, 4.80, 8.27, fmt=EU2, fte=AZUL,
       nota="IDÉNTICA en manual y Persán: el robot no abarata la madera")
c.fila("Mano de obra directa", 2.18, 1.02, 0.42, fmt=EU2, fte=AZUL,
       nota="NOTA 1: 38.000 € ÷ 90.000 ud, con un solo operario dedicado en exclusiva")
c.fila("Transporte", 0.62, 0.62, 0.53, fmt=EU2, fte=AZUL,
       nota="NOTA 2: eficiencia del tráiler propio a plena carga en las rutas de Persán")
c.fila("Otros costes variables", 1.30, 0.58, 1.30, fmt=EU2, fte=AZUL,
       nota="NOTA 3: varían proporcionalmente al volumen de fabricación")
r_cos = c.fila("COSTE TOTAL", f"=SUM(B{fi+1}:B{fi+4})", f"=SUM(C{fi+1}:C{fi+4})",
               f"=SUM(D{fi+1}:D{fi+4})", fmt=EU2, fte=TOT, fill=F_TOT)
r_mar = c.fila("MARGEN DE CONTRIBUCIÓN", f"=B{fi}-B{r_cos}", f"=C{fi}-C{r_cos}",
               f"=D{fi}-D{r_cos}", fmt=EU2, fte=TOT, fill=F_MAL,
               nota="HALLAZGO: Persán es el único negativo de los tres")
c.fila("Peso de la madera sobre el precio", f"=B{fi+1}/B{fi}", f"=C{fi+1}/C{fi}",
       f"=D{fi+1}/D{fi}", fmt=PCT,
       nota="El 83% del precio de Persán es madera, y el contrato no la revisa en 5 años")
c.fila("MARGEN SI LA MANO DE OBRA FUESE GRATIS", f"=B{fi}-B{fi+1}-B{fi+3}-B{fi+4}",
       f"=C{fi}-C{fi+1}-C{fi+3}-C{fi+4}", f"=D{fi}-D{fi+1}-D{fi+3}-D{fi+4}",
       fmt=EU2, fte=TOT, fill=F_MAL,
       nota="Sigue siendo negativo: el problema no es la productividad, es el precio")
c.gap()

c.sec("2. COSTES FIJOS Y VARIABLES DE LA EMPRESA", "→ VA AL INFORME")
c.txt("Del Anexo 1, reordenado por comportamiento en vez de por naturaleza. Sin esta "
      "separación no se puede juzgar un pedido nuevo. Reconstruye el EBITDA exacto.", F_GRIS)
c.cab(["Concepto", "Importe 2025", "% s/ventas", "", "Comentario"])
_rv = c.f
r_ven = c.fila("VENTAS NETAS", 4585267, f"=B{_rv}/B{_rv}", fmt=EUR, fte=AZUL, fill=F_TOT,
               fmts=[EUR, PCT])
r_cns = c.fila("Materia prima consumida", "=3068506-376062", f"=B{c.f}/B{r_ven}", fmt=EUR,
               fmts=[EUR, PCT], nota="Aprovisionamientos menos el aumento de existencias")
r_trp = c.fila("Transporte", 268218, f"=B{c.f}/B{r_ven}", fmt=EUR, fte=AZUL, fmts=[EUR, PCT])
r_otr = c.fila("Otros costes de explotación", 378979, f"=B{c.f}/B{r_ven}", fmt=EUR, fte=AZUL,
               fmts=[EUR, PCT])
r_cv = c.fila("TOTAL COSTES VARIABLES", f"=SUM(B{r_cns}:B{r_otr})", f"=B{c.f}/B{r_ven}",
              fmt=EUR, fte=TOT, fill=F_TOT, fmts=[EUR, PCT], nota="Crecen si se fabrica más")
r_mc = c.fila("MARGEN DE CONTRIBUCIÓN", f"=B{r_ven}-B{r_cv}", f"=B{c.f}/B{r_ven}", fmt=EUR,
              fte=TOT, fill=F_BIEN, fmts=[EUR, PCT],
              nota="De cada euro vendido quedan 27 céntimos")
r_pp = c.fila("Personal de producción", 481660, f"=B{c.f}/B{r_ven}", fmt=EUR, fte=AZUL,
              fmts=[EUR, PCT], nota="FIJO en la práctica: Esteban no quiere despidos")
r_pe = c.fila("Personal de estructura", 311292, f"=B{c.f}/B{r_ven}", fmt=EUR, fte=AZUL,
              fmts=[EUR, PCT])
r_cf = c.fila("TOTAL COSTES FIJOS", f"=B{r_pp}+B{r_pe}", f"=B{c.f}/B{r_ven}", fmt=EUR,
              fte=TOT, fill=F_TOT, fmts=[EUR, PCT], nota="Se pagan se fabrique o no")
c.fila("EBITDA reconstruido", f"=B{r_mc}-B{r_cf}", fmt=EUR, fte=TOT)
c.fila("Diferencia con el EBITDA del Anexo 1", f"=B{c.f-1}-452674", fmt=EUR, fte=TOT,
       fill=F_BIEN, nota="Cuadra al euro: la reordenación es indiscutible")
c.gap()
c.cab(["Punto de equilibrio", "Hoy", "Con Persán", "", "Comentario"])
r_rat = c.fila("Ratio de margen de contribución", f"=B{r_mc}/B{r_ven}", fmt=PCT, fte=TOT)
r_pm = c.fila("PUNTO DE EQUILIBRIO (resultado = 0)", f"=(B{r_cf}+36143+44046)/B{r_rat}",
              f"=(B{r_cf}+108000+36143+44046)/B{r_rat}", fmt=EUR, fte=TOT, fill=F_TOT,
              nota="Persán sube los costes fijos 108.000 € (campa + leasing) sin aportar margen")
c.fila("MARGEN DE SEGURIDAD", f"=(B{r_ven}-B{r_pm})/B{r_ven}", f"=(B{r_ven}-C{r_pm})/B{r_ven}",
       fmt=PCT, fte=TOT, fill=F_BIEN,
       nota="Cuánto puede caer la facturación antes de entrar en pérdidas")
c.gap()
c.cab(["¿Hay caja para invertir o repartir?", "Importe", "", "", "Comentario"])
c.fila("EBITDA 2025", 452674, fmt=EUR, fte=AZUL)
c.fila("(−) Aumento de existencias", -376062, fmt=EUR, fte=AZUL,
       nota="Madera comprada que sigue en el almacén")
c.fila("(−) Gastos financieros e impuestos", -137167, fmt=EUR, fte=AZUL)
c.fila("CAJA OPERATIVA REAL", f"=SUM(B{c.f-3}:B{c.f-1})", fmt=EUR, big=BIG, fill=F_MAL,
       nota="NEGATIVA, con 279.364 € de beneficio contable")
c.gap()

c.sec("3. ¿CABE PERSÁN EN LA FÁBRICA?", "→ VA AL INFORME")
c.txt("El caso da rendimientos por hora y por jornada, pero nunca la capacidad anual. "
      "Sin ella no se sabe si el contrato es viable ni si la máquina hace falta.", F_GRIS)
c.cab(["Sistema", "Palets/día", "Palets/año", "Operarios", "Cálculo"])
_rl = c.f
r_lin = c.fila("Línea automática (2019)", "=65*0.75*8", f"=B{_rl}*$B${r_dias}", 2,
               fmt=NUM, fmts=[NUM, NUM, NUM], nota="65 palets/h × 75% de rendimiento × 8 h")
_rm = c.f
r_man = c.fila("Fabricación manual", 500, f"=B{_rm}*$B${r_dias}", 8, fmt=NUM, fte=AZUL,
               fmts=[NUM, NUM, NUM], nota="500 al día entre 8 operarios = 62,5 cada uno")
r_cap = c.fila("CAPACIDAD ACTUAL", f"=B{r_lin}+B{r_man}", f"=C{r_lin}+C{r_man}",
               f"=D{r_lin}+D{r_man}", fmt=NUM, fte=TOT, fill=F_TOT)
_rr = c.f
r_rob = c.fila("Robot nuevo", "=400*0.9", f"=B{_rr}*$B${r_dias}", 1, fmt=NUM,
               fmts=[NUM, NUM, NUM], nota="400 palets/jornada × 90% de rendimiento")
c.gap()
c.cab(["Concepto", "Palets/año", "", "", "Comentario"])
r_prd = c.fila("Producción real de palet nuevo 2025", 174504, fmt=NUM, fte=AZUL, nota="Anexo 2")
c.fila("UTILIZACIÓN ACTUAL", f"=B{r_prd}/C{r_cap}", fmt=PCT, big=BIG, fill=F_CLA,
       nota="La fábrica NO está llena")
c.fila("Capacidad ociosa disponible hoy", f"=C{r_cap}-B{r_prd}", fmt=NUM, fte=TOT, fill=F_CLA,
       nota="Se puede fabricar esto sin invertir un euro. El «no llegamos» no se sostiene")
c.fila("Capacidad con robot / demanda con Persán", f"=C{r_cap}+C{r_rob}",
       f"=B{r_prd}+B{r_vol}", fmt=NUM, fte=TOT,
       nota="Con el robot cabe. Sin él faltarían 49.124 palets: si se firma, hace falta "
            "capacidad nueva sí o sí")
c.fila("Aviso: el robot solo llega a 87.120", f"=C{r_rob}-B{r_vol}", fmt=NUM, fte=TOT,
       fill=F_HIP,
       nota="La NOTA 1 calcula la mano de obra sobre 90.000 unidades, pero al 90% de "
            "rendimiento la máquina se queda 2.880 corta. Se completa con la línea de 2019, "
            "que tiene holgura. Efecto pequeño, pero conviene decirlo en el informe")
c.gap()

c.sec("4. ¿CUÁNTO CUESTA FIRMAR?", "→ VA AL INFORME")
c.cab(["Concepto", "Importe anual", "", "", "Comentario"])
c.fila("Ingresos", f"=B{r_vol}*B{r_pre}", fmt=EUR, nota="90.000 × 10 €")
c.fila("(−) Costes variables", f"=-B{r_vol}*D{r_cos}", fmt=EUR, nota="90.000 × 10,52 €")
r_mcp = c.fila("MARGEN DE CONTRIBUCIÓN", f"=B{c.f-2}+B{c.f-1}", fmt=EUR, fte=TOT, fill=F_MAL,
               nota="Negativo ANTES de cualquier coste fijo")
c.fila("(−) Alquiler de la campa", f"=-B{r_cam}", fmt=EUR)
c.fila("(−) Leasing de máquina y camión", f"=-B{r_lea}", fmt=EUR)
r_cir = c.fila("Circulante inmovilizado", f"=B{r_vol}*B{r_pre}*B{r_cob}/365+B{r_ret}*D{r_cos}",
               fmt=EUR,
               nota="Cobro a 180 días (443.836 €) más el retén de 1.000 palets (10.520 €)")
r_cfi = c.fila("(−) Coste financiero del circulante", f"=-B{r_cir}*B{r_int}", fmt=EUR,
               nota="Al 7,5%. El escandallo no lo recoge, pero es un coste real")
r_imp = c.fila("IMPACTO ANUAL DEL CONTRATO", f"=B{r_mcp}-B{r_cam}-B{r_lea}+B{r_cfi}",
               fmt=EUR, big=BIG, fill=F_MAL)
c.fila("En % del beneficio de la empresa", f"=B{r_imp}/279364", fmt=PCT, fte=TOT, fill=F_MAL,
       nota="Se lleva por delante dos tercios del beneficio de TODA la empresa")
c.fila("Impacto acumulado a 5 años", f"=B{r_imp}*5", fmt=EUR, fte=TOT, fill=F_MAL,
       nota="El contrato dura 5 años sin revisión de precio")
c.fila("Caja necesaria el primer año", f"=B{r_imp}-B{r_cir}", fmt=EUR, fte=TOT, fill=F_MAL,
       nota="Frente a una caja operativa de −60.555 € en 2025")
c.gap()

c.sec("5. ¿A QUÉ PRECIO SÍ?", "→ VA AL INFORME")
c.cab(["Nivel de cobertura", "Precio mínimo", "Subida", "", "Qué cubre"])
c.fila("1. Solo los costes variables", f"=D{r_cos}", f"=B{c.f}/B{r_pre}-1", fmt=EU2,
       fmts=[EU2, PCT], nota="Madera, mano de obra, transporte y variables")
r_b2 = c.fila("2. + campa y leasing", f"=D{r_cos}+(B{r_cam}+B{r_lea})/B{r_vol}",
              f"=B{c.f}/B{r_pre}-1", fmt=EU2, fmts=[EU2, PCT],
              nota="Añade los costes fijos que el contrato obliga a asumir")
c.fila("3. + coste del circulante",
       f"=(B{r_vol}*D{r_cos}+B{r_cam}+B{r_lea})/(B{r_vol}*(1-B{r_cob}/365*B{r_int}))",
       f"=B{c.f}/B{r_pre}-1", fmt=EU2, fmts=[EU2, PCT], big=BIG, fill=F_CLA,
       nota="PRECIO DE EQUILIBRIO REAL")
c.fila("4. + margen igual al del negocio actual", f"=B{r_b2}+B{r_mar}",
       f"=B{c.f}/B{r_pre}-1", fmt=EU2, fmts=[EU2, PCT],
       nota="Para que Persán rinda como el resto del palet nuevo")
c.gap()
c.cab(["Escenario de negociación", "Precio", "Madera", "Cobro (días)", "Impacto anual"])
NM = 0.42 + 0.53 + 1.30
for et, pr, ma, di, fill in [
    ("Contrato actual", 10.00, 8.27, 180, F_MAL),
    ("Solo el plazo de cobro", 10.00, 8.27, 90, F_MAL),
    ("Solo la madera re-especificada", 10.00, 7.50, 180, F_MAL),
    ("Mínimo aceptable", 12.17, 8.27, 180, F_CLA),
    ("Precio y plazo", 12.50, 8.27, 90, F_BIEN),
    ("CONTRATO IDEAL: las tres palancas", 12.50, 7.50, 90, F_BIEN),
]:
    f = c.f
    ws = c.ws
    ws.cell(f, 1, et).font = TOT if "IDEAL" in et else ETI
    for col, v, fm in ((2, pr, EU2), (3, ma, EU2), (4, di, NUM)):
        cc = ws.cell(f, col, v)
        cc.font = NEG
        cc.number_format = fm
        cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 5, f"=90000*(B{f}-(C{f}+{NM}))-108000"
                       f"-(B{f}*90000*D{f}/365+1000*(C{f}+{NM}))*0.075")
    cc.font = TOT
    cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    for col in range(1, 6):
        ws.cell(f, col).fill = fill
    c.f += 1
c.txt("Ninguna palanca por separado salva el contrato. Hay que subir el precio sí o sí, pero "
      "combinarlo con las otras dos reduce lo que hay que pedir: en vez de exigir 12,17 € a "
      "secas, se cierra en 12,50 € con un palet más barato de fabricar y cobrando antes.",
      F_GRIS)
c.gap()

c.sec("6. LA MÁQUINA, POR SEPARADO", "→ VA AL INFORME")
c.txt("El caso presenta máquina y contrato como una sola decisión. Tienen plazos distintos: "
      "el fabricante pide respuesta HOY, el contrato es todavía un borrador.", F_GRIS)
c.cab(["Concepto", "Manual", "Línea 2019", "Robot", "Comentario"])
c.fila("Palets por operario y día", 62.5, 195, 360, fmt=NU2,
       nota="El robot multiplica por 5,8 la productividad manual")
c.fila("Cambio de formato (horas)", "", 6, 0.42, fmt=NU2, fill=F_CLA,
       nota="LA RAZÓN DE COMPRARLA: 6 horas de parada frente a 25 minutos. Alcopalet gana "
            "dinero en series especiales (114×114 de Don Simón, Tetra Pak). Con 6 horas de "
            "parada una serie corta es inviable; con media hora, es negocio")
c.gap()
c.cab(["Análisis de la inversión", "Importe", "", "", "Comentario"])
c.fila("Inversión", 40000, fmt=EUR, fte=AZUL, nota="Sobre un beneficio de 279.364 €")
r_lm = c.fila("Leasing solo de la máquina (€/año)",
              "=40000*(0.075/12)/(1-(1+0.075/12)^-60)*12", fmt=EUR, fte=TOT, fill=F_BIEN,
              nota="Frente a 48.000 €/año si se financia junto con el tráiler: separar las dos "
                   "compras ahorra 38.382 € al año")
c.fila("Ahorro si el personal liberado se reasigna", 180880, fmt=EUR, fte=TOT, fill=F_HIP,
       nota="4,76 operarios equivalentes × 38.000 €. HIPÓTESIS: solo es real si esas personas "
            "producen algo vendible. Esteban no quiere despidos, así que hay que reasignarlas: "
            "la línea de palet usado necesita 2,6 más y solo le falta espacio")
c.fila("Plazo de recuperación (meses)", "=40000/180880*12", fmt=NU2, fte=TOT, fill=F_BIEN,
       nota="Menos de tres meses en ese escenario")
c.fila("Peor caso: nadie se reasigna", f"=-B{r_lm}", fmt=EUR, fte=TOT,
       nota="Se pierde solo el leasing. Riesgo acotado")
c.gap()

c.sec("7. LAS OPCIONES", "→ VA AL INFORME")
c.cab(["Opción", "Impacto anual", "Resultado neto", "Caja año 1", "Valoración"])
for et, imp, caja, val, fill in [
    ("A — No hacer nada", 0, 0, "Congela la empresa en los 3.000 m² actuales", None),
    ("B — Persán + robot (Esteban)", -188877, -643232,
     "Insostenible: se lleva el 68% del beneficio", F_MAL),
    ("C — Dividendo sin invertir (padres)", 0, -200000,
     "No hay caja: en 2025 se generaron −60.555 €", F_MAL),
    ("D — Robot solo, sin tráiler", -9618, -9618,
     "Riesgo mínimo y gana flexibilidad", F_CLA),
    ("E — Robot + Persán renegociado", 117964, -169183, "RECOMENDADA", F_BIEN),
]:
    f = c.f
    ws = c.ws
    ws.cell(f, 1, et).font = TOT if fill == F_BIEN else ETI
    cc = ws.cell(f, 2, imp); cc.font = TOT; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 3, f"=279364+B{f}"); cc.font = TOT; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    cc = ws.cell(f, 4, caja); cc.font = NEG; cc.number_format = EUR
    cc.alignment = Alignment(horizontal="center")
    ws.cell(f, 5, val).font = MINI
    ws.cell(f, 5).alignment = Alignment(wrap_text=True, vertical="center")
    if fill:
        for col in range(1, 6):
            ws.cell(f, col).fill = fill
    c.f += 1
c.gap()
c.cab(["¿Cumple las restricciones del caso?", "B", "D", "E", "Comentario"])
for crit, b, dd, e, nota in [
    ("El contrato cubre sus costes", "NO", "n/a", "SÍ", "Solo E supera los 12,17 €"),
    ("Cabe en la tesorería", "NO", "SÍ", "AJUSTADO", "B exige 643.000 € el primer año"),
    ("Respeta «sin despidos»", "SÍ", "SÍ", "SÍ", "La capacidad extra se necesita igual"),
    ("Respeta el turno único", "SÍ", "SÍ", "SÍ", None),
    ("Permite crecer hasta 2030", "SÍ", "PARCIAL", "SÍ", None),
    ("Da seguridad a los fundadores", "NO", "PARCIAL", "PARCIAL",
     "B deja el beneficio en 90.487 €: sin dividendo en cinco años"),
]:
    f = c.f
    ws = c.ws
    ws.cell(f, 1, crit).font = ETI
    for col, v in ((2, b), (3, dd), (4, e)):
        cc = ws.cell(f, col, v)
        cc.font = TOT
        cc.alignment = Alignment(horizontal="center")
        cc.fill = F_MAL if v == "NO" else (F_BIEN if v == "SÍ" else F_CLA)
    if nota:
        ws.cell(f, 5, nota).font = MINI
        ws.cell(f, 5).alignment = Alignment(wrap_text=True, vertical="center")
    c.f += 1

# ══════════════════════════════════════════════════════════════ ANEXOS
a = H(wb, "ANEXOS", "Los anexos del informe final",
      "Máximo 10 carillas INCLUYENDO anexos, portada e índice. Reparto: portada e índice 1 + "
      "informe 5 + anexos 3,5 = 9,5 carillas.")

a.sec("ANEXO 1 — ESTRUCTURA DE COSTES Y PUNTO DE EQUILIBRIO", "½ carilla")
a.cab(["Qué incluir", "", "", "", "Por qué"])
a.fila("Tabla de costes fijos y variables sobre ventas",
       nota="Con el margen de contribución del 27,2%. Sin esto no se puede juzgar el contrato")
a.fila("La línea de comprobación del EBITDA",
       nota="Diferencia cero con el Anexo 1: hace la reordenación indiscutible")
a.fila("Punto de equilibrio y margen de seguridad",
       nota="3.214.115 € y 29,9%. Demuestra que la empresa está sana HOY")
a.fila("Efecto de Persán sobre el punto de equilibrio",
       nota="Sube casi 400.000 € por los 108.000 € de costes fijos adicionales")
a.gap()

a.sec("ANEXO 2 — ANÁLISIS ECONÓMICO DEL CONTRATO PERSÁN", "1 carilla · EL CENTRAL")
a.cab(["Qué incluir", "", "", "", "Por qué"])
a.fila("Escandallo comparado: nuevo, usado y Persán",
       nota="El margen de −0,52 €/palet sale del propio Anexo 3 del caso")
a.fila("Margen sin mano de obra: −0,10 €/palet",
       nota="EL ARGUMENTO DEFINITIVO: aunque el operario fuese gratis, el contrato pierde")
a.fila("Cuenta de resultados incremental completa",
       nota="Del margen al impacto anual de −188.877 €, con campa, leasing y circulante")
a.fila("Los cuatro niveles de precio de equilibrio",
       nota="Y el resultado: 12,17 €, frente a los 10 € ofrecidos")
a.fila("Tabla de escenarios de negociación",
       nota="Demuestra que ninguna palanca por separado salva el contrato")
a.gap()

a.sec("ANEXO 3 — CAPACIDAD PRODUCTIVA", "½ carilla")
a.cab(["Qué incluir", "", "", "", "Por qué"])
a.fila("Capacidad anual de los tres sistemas",
       nota="Con el cálculo a la vista: 242 días productivos según la nota final del Anexo 3")
a.fila("Grado de utilización actual: 81%",
       nota="Desmonta el «no llegamos». Hay 40.876 palets de holgura sin invertir un euro")
a.fila("Encaje de Persán con y sin robot",
       nota="Sin máquina faltan 49.124 palets: si se firma, hace falta capacidad sí o sí")
a.fila("Recorrido de la línea de palet usado",
       nota="99.670 unidades bloqueadas por ESPACIO, no por demanda. La campa que se "
            "alquilaría para el retén de Persán desbloquearía una línea que sí gana dinero")
a.gap()

a.sec("ANEXO 4 — ESCENARIOS Y TÉRMINOS DE LA RENEGOCIACIÓN", "1 carilla")
a.cab(["Qué incluir", "", "", "", "Por qué"])
a.fila("Tabla comparativa de las cinco opciones",
       nota="Impacto anual, resultado proyectado y caja del primer año")
a.fila("Cumplimiento de las restricciones del caso",
       nota="Sin despidos, turno único, crecimiento y seguridad de los fundadores")
a.fila("Análisis de la inversión de la máquina",
       nota="40.000 €, leasing de 9.618 €/año sin el tráiler, recuperación en menos de 3 meses")
a.fila("Los términos del contrato ideal",
       nota="Las tres palancas con su objetivo y la cláusula de revisión de la madera. "
            "Responde al criterio 8 de evaluación: creatividad para generar alternativas")
a.gap()

a.sec("QUÉ DEJARÍA FUERA, Y POR QUÉ")
for t in [
    "La reconciliación entre el Anexo 1 y el Anexo 3 (288 € de desviación sobre 2,7 millones). "
    "Imprescindible haberla hecho, pero es control interno: una frase en el cuerpo del informe "
    "y guardada para el turno de preguntas.",
    "El endeudamiento implícito (587.280 €, que pasaría de 1,30x a 4,13x EBITDA). Buen "
    "argumento, pero cabe en dos líneas del cuerpo.",
    "El coste por camión (134.109 € frente a los 47.700 € que asigna el Anexo 3). Es una "
    "inferencia propia, no un dato del caso: mejor para la defensa oral.",
    "El análisis del Anexo 2 por líneas de negocio. Sus conclusiones caben en tres filas "
    "dentro del cuerpo: no merece un anexo propio.",
    "Los escenarios de rendimiento de la máquina al 85%, 80% y 75%. La nota 1 del caso manda "
    "calcular sobre 90.000 unidades; el matiz va en una línea, no en un anexo.",
]:
    a.txt("• " + t)

wb.save("/home/user/DANISANTELMO/Alcopalet_Analisis.xlsx")
print("Guardado:", [x.title for x in wb.worksheets])
