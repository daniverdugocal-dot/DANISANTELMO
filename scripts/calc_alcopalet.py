"""Alcopalet DTI-1399 — verificación numérica del caso."""

# --- Datos del caso ---
DIAS = 11 * 22          # 242 días productivos/año
JORNADA_H = 8           # 7:00 a 15:00
COSTE_MO = 38_000       # €/año por operario
COSTE_DESPIDO = 30_000

# Anexo 1 — P&G 2025
VENTAS = 4_585_267
VAR_EXIST = 376_062
APROV = 3_068_506
MB = 1_892_822
PERS_PROD = 481_660
TRANSP = 268_218
PERS_ESTR = 311_292
OTROS = 378_979
EBITDA = 452_674
AMORT = 36_143
EBIT = 416_531
FIN = 44_046
BAI = 372_485
IMP = 93_121
RN = 279_364

# Anexo 2 — líneas
V_NUEVO, RN_NUEVO, U_NUEVO = 2_645_224, 158_882, 174_504
V_USADO, RN_USADO, U_USADO = 1_940_042, 120_482, 260_330

# Anexo 3 — escandallo €/unidad
man = dict(ing=15.16, mp=8.27, mod=2.18, tr=0.62, otros=1.30)
usa = dict(ing=7.45,  mp=4.80, mod=1.02, tr=0.62, otros=0.58)
per = dict(ing=10.00, mp=8.27, mod=0.42, tr=0.53, otros=1.30)

def coste(d): return d['mp'] + d['mod'] + d['tr'] + d['otros']
def margen(d): return d['ing'] - coste(d)

print("=" * 68)
print("1. VERIFICACIÓN DE COHERENCIA DE LOS ANEXOS")
print("=" * 68)
print(f"Precio medio nuevo : {V_NUEVO/U_NUEVO:6.2f} €  (anexo 3: 15,16)")
print(f"Precio medio usado : {V_USADO/U_USADO:6.2f} €  (anexo 3:  7,45)")
print(f"Suma ventas líneas : {V_NUEVO+V_USADO:,} vs P&G {VENTAS:,}")
print(f"Suma rdo. líneas   : {RN_NUEVO+RN_USADO:,} vs P&G {RN:,}")
print()
for n, d in (("Manual", man), ("Usado", usa), ("Persán", per)):
    print(f"{n:8s} coste unit. {coste(d):6.2f} €   margen contrib. {margen(d):+6.2f} €")
print()

# Reconciliación escandallo vs P&G
mp_tot = man['mp']*U_NUEVO + usa['mp']*U_USADO
tr_tot = man['tr']*U_NUEVO + usa['tr']*U_USADO
ot_tot = man['otros']*U_NUEVO + usa['otros']*U_USADO
mod_tot = man['mod']*U_NUEVO + usa['mod']*U_USADO
print(f"Materia prima escandallo : {mp_tot:12,.0f} €")
print(f"Consumo real P&G (Aprov - Var.exist): {APROV-VAR_EXIST:12,.0f} €  -> desvío {mp_tot-(APROV-VAR_EXIST):+,.0f}")
print(f"Transporte escandallo    : {tr_tot:12,.0f} € vs P&G {TRANSP:,}  -> desvío {tr_tot-TRANSP:+,.0f}")
print(f"Otros costes escandallo  : {ot_tot:12,.0f} € vs P&G {OTROS:,}  -> desvío {ot_tot-OTROS:+,.0f}")
print(f"  % otros costes de palet nuevo: {man['otros']*U_NUEVO/ot_tot:.1%}  (caso dice 60%)")
print(f"MOD escandallo           : {mod_tot:12,.0f} € vs P&G Pers.Prod {PERS_PROD:,}  -> desvío {mod_tot-PERS_PROD:+,.0f}")
print(f"  MOD nuevo {man['mod']*U_NUEVO:,.0f} = 10 operarios x 38.000 = {10*COSTE_MO:,}")
print(f"  MOD usado {usa['mod']*U_USADO:,.0f} =  7 operarios x 38.000 = {7*COSTE_MO:,}")
print()

print("=" * 68)
print("2. CAPACIDAD PRODUCTIVA (palet nuevo)")
print("=" * 68)
linea_h = 65 * 0.75
linea_dia = linea_h * JORNADA_H
linea_anio = linea_dia * DIAS
manual_dia = 500
manual_anio = manual_dia * DIAS
print(f"Línea 2019 : {65} p/h x 75% = {linea_h:.2f} p/h x {JORNADA_H}h = {linea_dia:.0f}/día -> {linea_anio:,.0f}/año (2 operarios)")
print(f"Manual     : {manual_dia}/día ({manual_dia/8:.1f} p/operario) -> {manual_anio:,.0f}/año (8 operarios)")
cap_actual = linea_anio + manual_anio
print(f"CAPACIDAD ACTUAL TOTAL: {cap_actual:,.0f}/año")
print(f"Producción real 2025  : {U_NUEVO:,} -> utilización {U_NUEVO/cap_actual:.1%}")
print(f"Capacidad ociosa      : {cap_actual-U_NUEVO:,.0f} palets/año")
print()
robot_dia = 400 * 0.90
robot_anio = robot_dia * DIAS
print(f"Robot nuevo: 400 x 90% = {robot_dia:.0f}/día -> {robot_anio:,.0f}/año (1 operario)")
print(f"Persán exige 90.000/año = {90_000/DIAS:.1f}/día")
print(f"  -> DÉFICIT del robot en solitario: {robot_anio-90_000:+,.0f} palets/año ({(robot_anio-90_000)/90_000:+.1%})")
print()

print("=" * 68)
print("3. IMPACTO ECONÓMICO DEL CONTRATO PERSÁN (90.000 ud/año)")
print("=" * 68)
Q = 90_000
ing = Q * per['ing']
cv = Q * coste(per)
mc = Q * margen(per)
campa = 5_000 * 12
leasing = 4_000 * 12
print(f"Ingresos                        : {ing:+12,.0f} €")
print(f"Costes variables                : {-cv:+12,.0f} €")
print(f"MARGEN DE CONTRIBUCIÓN          : {mc:+12,.0f} €   <-- NEGATIVO")
print(f"Alquiler campa (5.000 x 12)     : {-campa:+12,.0f} €")
print(f"Leasing máquina+camión (4.000x12): {-leasing:+12,.0f} €")
sub = mc - campa - leasing
print(f"RESULTADO ANTES DE CIRCULANTE   : {sub:+12,.0f} €")
# circulante: cobro a 180 días
cta_cli = ing * 180 / 365
coste_circ = cta_cli * 0.075
reten = 1_000 * coste(per)
print()
print(f"Cuenta a cobrar (180 días)      : {cta_cli:12,.0f} €")
print(f"Retén 1.000 palets (a coste)    : {reten:12,.0f} €")
print(f"Coste financiero circulante 7,5%: {-coste_circ:+12,.0f} €")
tot = sub - coste_circ
print(f"IMPACTO ANUAL TOTAL             : {tot:+12,.0f} €")
print(f"  = {tot/RN:+.1%} del resultado neto actual ({RN:,} €)")
print(f"  A 5 años: {tot*5:+,.0f} €")
print()
# break-even
be_var = coste(per)
be_tot = coste(per) + (campa + leasing) / Q
be_circ = coste(per) + (campa + leasing + coste_circ) / Q
print(f"Precio de equilibrio (solo variables)      : {be_var:6.2f} €/ud")
print(f"Precio de equilibrio (+ campa y leasing)   : {be_tot:6.2f} €/ud")
print(f"Precio de equilibrio (+ coste circulante)  : {be_circ:6.2f} €/ud")
print(f"Precio para igualar margen actual (2,80 €) : {be_tot+2.80:6.2f} €/ud")
print(f"Precio ofrecido por Persán                 : {per['ing']:6.2f} €/ud")
print(f"  -> déficit sobre equilibrio: {per['ing']-be_circ:+.2f} €/ud")
print()
print(f"Peso de la madera sobre el precio Persán: {per['mp']/per['ing']:.1%}")
for infl in (0.05, 0.10, 0.15):
    nm = per['ing'] - (coste(per) + per['mp']*infl)
    print(f"  Si la madera sube {infl:.0%}: margen unitario {nm:+.2f} € -> {nm*Q:+,.0f} €/año")
print()

print("=" * 68)
print("4. EL ROBOT COMO DECISIÓN INDEPENDIENTE")
print("=" * 68)
prod_op_manual = manual_dia / 8
op_equiv = robot_anio / (prod_op_manual * DIAS)
print(f"Productividad manual: {prod_op_manual:.1f} palets/operario/día")
print(f"Robot ({robot_anio:,.0f}/año) equivale a {op_equiv:.2f} operarios manuales")
ahorro = (op_equiv - 1) * COSTE_MO
print(f"Ahorro bruto de mano de obra: ({op_equiv:.2f} - 1) x 38.000 = {ahorro:,.0f} €/año")
leasing_maq = 40_000 * (0.075/12) / (1 - (1+0.075/12)**-60) * 12
print(f"Leasing solo máquina (40.000 € al 7,5%, 5 años): {leasing_maq:,.0f} €/año")
print(f"Ahorro neto: {ahorro-leasing_maq:+,.0f} €/año")
print(f"Payback simple sobre 40.000 €: {40_000/ahorro*12:.1f} meses")
print()
# margen del robot a precio de mercado
robot_mercado = dict(ing=15.16, mp=8.27, mod=0.42, tr=0.62, otros=1.30)
print(f"Margen del robot a PRECIO DE MERCADO (15,16 €): {margen(robot_mercado):+.2f} €/ud")
print(f"Margen del robot a PRECIO PERSÁN     (10,00 €): {margen(per):+.2f} €/ud")
print(f"Diferencia por unidad: {margen(robot_mercado)-margen(per):+.2f} €")
print(f"COSTE DE OPORTUNIDAD sobre {robot_anio:,.0f} ud/año: {(margen(robot_mercado)-margen(per))*robot_anio:+,.0f} €/año")
print()

print("=" * 68)
print("5. CAPACIDAD DE PAGO: ¿HAY CAJA PARA DIVIDENDOS?")
print("=" * 68)
print(f"EBITDA 2025                       : {EBITDA:+12,.0f} €")
print(f"Aumento de existencias (salida)   : {-VAR_EXIST:+12,.0f} €")
print(f"Gastos financieros                : {-FIN:+12,.0f} €")
print(f"Impuestos                         : {-IMP:+12,.0f} €")
fcf = EBITDA - VAR_EXIST - FIN - IMP
print(f"CAJA OPERATIVA APROXIMADA 2025    : {fcf:+12,.0f} €")
print(f"(resultado neto contable: {RN:,} €)")
print()
print("Necesidades de caja del proyecto Persán en el año 1:")
print(f"  Circulante por cobro a 180 días : {-cta_cli:+12,.0f} €")
print(f"  Campa + leasing                 : {-(campa+leasing):+12,.0f} €")
print(f"  Margen de contribución negativo : {mc:+12,.0f} €")
print(f"  TOTAL                           : {-(cta_cli+campa+leasing)+mc:+12,.0f} €")
print()

print("=" * 68)
print("6. LÍNEA DE PALET USADO: RECORRIDO BLOQUEADO POR ESPACIO")
print("=" * 68)
techo_anual = 30_000 * 12
print(f"Producción actual : {U_USADO:,} ud/año ({U_USADO/12:,.0f}/mes)")
print(f"Techo logístico   : {techo_anual:,} ud/año (30.000/mes)")
rec = techo_anual - U_USADO
print(f"Recorrido         : {rec:,} ud/año -> contribución {rec*margen(usa):+,.0f} €/año")
ops = techo_anual / DIAS / 155
print(f"Operarios necesarios para el techo: {ops:.1f} (hoy 7) -> +{ops-7:.1f} operarios")
print(f"Coste de la campa: {campa:,} €/año  -> {'NO cubre' if rec*margen(usa) < campa else 'cubre'} el alquiler por sí sola")
