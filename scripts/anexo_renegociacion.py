"""
Genera la imagen del anexo «Escenario de renegociación con Persán».

Salida: anexo_renegociacion_persan.png, lista para pegar en el informe.
Todas las cifras se calculan aquí, no se escriben a mano.
"""
from pathlib import Path

# ---------------------------------------------------------------- cálculo
Q = 90_000          # unidades/año del contrato
I = 0.075           # tipo de interés del banco
CAMPA = 5_000 * 12
LEASING = 4_000 * 12
RETEN = 1_000       # palets de retén permanente
RN_ACTUAL = 279_364  # resultado neto 2025

# Componentes del escandallo (Anexo 3, notas 1, 2 y 3)
MOD, TRANSPORTE, OTROS = 0.42, 0.53, 1.30


def escenario(precio, madera, dias):
    coste = madera + MOD + TRANSPORTE + OTROS
    margen = precio - coste
    mc = margen * Q
    circulante = precio * Q * dias / 365 + RETEN * coste
    financiero = circulante * I
    impacto = mc - CAMPA - LEASING - financiero
    return dict(precio=precio, madera=madera, dias=dias, coste=coste, margen=margen,
                mc=mc, circulante=circulante, financiero=financiero, impacto=impacto,
                resultado=RN_ACTUAL + impacto)


ACT = escenario(10.00, 8.27, 180)
REN = escenario(12.50, 7.50, 90)

ESCALERA = [
    ("Contrato actual", escenario(10.00, 8.27, 180)["impacto"]),
    ("Solo se consigue el plazo de cobro (90 días)", escenario(10.00, 8.27, 90)["impacto"]),
    ("Solo se consigue re-especificar la madera", escenario(10.00, 7.50, 180)["impacto"]),
    ("Solo se consigue el precio de equilibrio (12,17 €)",
     escenario(12.17, 8.27, 180)["impacto"]),
    ("OBJETIVO: las tres palancas juntas", escenario(12.50, 7.50, 90)["impacto"]),
]


def eur(v, dec=0):
    s = f"{abs(v):,.{dec}f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return ("−" if v < 0 else "") + s + " €"


def num(v, dec=0):
    s = f"{abs(v):,.{dec}f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return ("−" if v < 0 else "") + s


# ---------------------------------------------------------------- HTML
def fila_esc(nombre, valor, destacado=False):
    ancho = min(100, abs(valor) / 200_000 * 100)
    color = "#c00000" if valor < 0 else "#2e7d32"
    lado = "right: 50%;" if valor < 0 else "left: 50%;"
    cls = ' class="destacado"' if destacado else ""
    return f"""
      <tr{cls}>
        <td class="nombre">{nombre}</td>
        <td class="barra">
          <div class="eje"></div>
          <div class="rect" style="{lado} width:{ancho / 2:.1f}%; background:{color};"></div>
        </td>
        <td class="cifra" style="color:{color}">{eur(valor)}</td>
      </tr>"""


HTML = f"""<!doctype html>
<meta charset="utf-8">
<style>
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0; padding: 34px 38px;
    font-family: Carlito, Calibri, "Liberation Sans", Arial, sans-serif;
    color: #1a1a1a; background: #fff; width: 1180px;
  }}
  h1 {{ font-size: 21px; margin: 0 0 3px; color: #1f3864; letter-spacing: -.2px; }}
  .sub {{ font-size: 12px; color: #666; margin-bottom: 22px; }}
  h2 {{
    font-size: 12.5px; text-transform: uppercase; letter-spacing: .6px;
    color: #fff; background: #1f3864; padding: 6px 10px; margin: 20px 0 0;
  }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th {{
    background: #4472c4; color: #fff; font-weight: 600; text-align: center;
    padding: 7px 10px; border: 1px solid #cfd8e8; font-size: 12.5px;
  }}
  th.izq, td.izq {{ text-align: left; }}
  td {{ padding: 6px 10px; border: 1px solid #dde3ee; text-align: center; }}
  td.izq {{ width: 40%; }}
  tr.total td {{ background: #d9e2f3; font-weight: 700; }}
  tr.mal td {{ background: #fce4e4; }}
  tr.bien td {{ background: #e2efda; }}
  .rojo {{ color: #c00000; font-weight: 700; }}
  .verde {{ color: #2e7d32; font-weight: 700; }}
  .grande {{ font-size: 16px; }}
  .nota {{ font-size: 11px; color: #666; font-style: italic; margin-top: 6px; }}
  .cols {{ display: flex; gap: 22px; }}
  .cols > div {{ flex: 1; }}

  table.escalera td {{ border: none; padding: 5px 8px; }}
  td.nombre {{ text-align: left; width: 42%; font-size: 12.5px; }}
  td.barra {{ position: relative; height: 22px; width: 40%; }}
  td.cifra {{ text-align: right; width: 18%; font-weight: 700; font-size: 13px; }}
  .eje {{ position: absolute; left: 50%; top: 2px; bottom: 2px; width: 1px; background: #b0b0b0; }}
  .rect {{ position: absolute; top: 5px; height: 13px; border-radius: 2px; }}
  tr.destacado td {{ background: #e2efda; }}
  tr.destacado td.nombre {{ font-weight: 700; }}
</style>

<h1>Anexo — Escenario de renegociación del contrato con Persán</h1>
<div class="sub">Alcopalet (DTI&#8209;1399) · 90.000 palets/año · Cifras anuales ·
Coste unitario según las notas 1, 2 y 3 del Anexo 3 del caso</div>

<h2>1. Los términos que hay que negociar</h2>
<table>
  <tr>
    <th class="izq">Variable</th><th>Contrato actual</th><th>Mínimo aceptable</th>
    <th>Objetivo</th><th class="izq">Justificación</th>
  </tr>
  <tr>
    <td class="izq">Precio por palet</td>
    <td class="rojo">{eur(ACT['precio'], 2)}</td><td>12,17 €</td>
    <td class="verde">{eur(REN['precio'], 2)}</td>
    <td class="izq">El mínimo es el punto de equilibrio exacto del contrato</td>
  </tr>
  <tr>
    <td class="izq">Coste de la madera</td>
    <td class="rojo">{eur(ACT['madera'], 2)}</td><td>8,00 €</td>
    <td class="verde">{eur(REN['madera'], 2)}</td>
    <td class="izq">Re-especificar el palet: el estándar Tetra&nbsp;Pak (madera gallega
        premium, seca y lijada) no lo exige el destino de Persán</td>
  </tr>
  <tr>
    <td class="izq">Plazo de cobro</td>
    <td class="rojo">{num(ACT['dias'])} días</td><td>120 días</td>
    <td class="verde">{num(REN['dias'])} días</td>
    <td class="izq">Reduce a la mitad el circulante inmovilizado</td>
  </tr>
  <tr>
    <td class="izq">Revisión del precio de la madera</td>
    <td class="rojo">No prevista</td><td>Anual</td>
    <td class="verde">Anual</td>
    <td class="izq">Innegociable: la madera es el 83&nbsp;% del precio y el contrato dura
        5 años</td>
  </tr>
</table>

<div class="cols">
<div>
<h2>2. Efecto sobre el margen unitario</h2>
<table>
  <tr><th class="izq">€ por palet</th><th>Actual</th><th>Renegociado</th></tr>
  <tr><td class="izq">Ingreso por palet</td><td>{eur(ACT['precio'], 2)}</td>
      <td>{eur(REN['precio'], 2)}</td></tr>
  <tr><td class="izq">Materia prima (madera y clavos)</td>
      <td>{eur(ACT['madera'], 2)}</td><td>{eur(REN['madera'], 2)}</td></tr>
  <tr><td class="izq">Mano de obra directa</td><td>{eur(MOD, 2)}</td>
      <td>{eur(MOD, 2)}</td></tr>
  <tr><td class="izq">Transporte</td><td>{eur(TRANSPORTE, 2)}</td>
      <td>{eur(TRANSPORTE, 2)}</td></tr>
  <tr><td class="izq">Otros costes variables</td><td>{eur(OTROS, 2)}</td>
      <td>{eur(OTROS, 2)}</td></tr>
  <tr class="total"><td class="izq">Coste total</td><td>{eur(ACT['coste'], 2)}</td>
      <td>{eur(REN['coste'], 2)}</td></tr>
  <tr class="total"><td class="izq">MARGEN DE CONTRIBUCIÓN</td>
      <td class="rojo grande">{eur(ACT['margen'], 2)}</td>
      <td class="verde grande">+{eur(REN['margen'], 2)}</td></tr>
</table>
<div class="nota">El margen actual es negativo antes de cualquier coste fijo: el precio no
cubre lo que consume fabricar el palet.</div>
</div>

<div>
<h2>3. Efecto sobre el resultado anual</h2>
<table>
  <tr><th class="izq">Concepto</th><th>Actual</th><th>Renegociado</th></tr>
  <tr><td class="izq">Margen de contribución total</td>
      <td>{eur(ACT['mc'])}</td><td>{eur(REN['mc'])}</td></tr>
  <tr><td class="izq">Alquiler de la campa</td><td>{eur(-CAMPA)}</td>
      <td>{eur(-CAMPA)}</td></tr>
  <tr><td class="izq">Leasing de máquina y camión</td><td>{eur(-LEASING)}</td>
      <td>{eur(-LEASING)}</td></tr>
  <tr><td class="izq">Coste financiero del circulante</td>
      <td>{eur(-ACT['financiero'])}</td><td>{eur(-REN['financiero'])}</td></tr>
  <tr class="total"><td class="izq">IMPACTO ANUAL DEL CONTRATO</td>
      <td class="rojo grande">{eur(ACT['impacto'])}</td>
      <td class="verde grande">+{eur(REN['impacto'])}</td></tr>
  <tr class="total"><td class="izq">Resultado neto de la empresa</td>
      <td class="rojo">{eur(ACT['resultado'])}</td>
      <td class="verde">{eur(REN['resultado'])}</td></tr>
  <tr><td class="izq">Circulante inmovilizado</td>
      <td>{eur(ACT['circulante'])}</td><td>{eur(REN['circulante'])}</td></tr>
</table>
<div class="nota">Resultado neto de 2025 antes del contrato: 279.364 €.</div>
</div>
</div>

<h2>4. Ninguna palanca por separado salva el contrato</h2>
<table class="escalera">
  {''.join(fila_esc(n, v, n.startswith('OBJETIVO')) for n, v in ESCALERA)}
</table>
<div class="nota">Impacto anual sobre el resultado según lo que se consiga en la
negociación. Hay que subir el precio necesariamente, pero combinarlo con las otras dos
palancas reduce la subida exigible: en lugar de imponer 12,17 €, se cierra en 12,50 € con
un palet más barato de fabricar y cobrando a la mitad de plazo.</div>
"""

ruta_html = Path("/tmp/anexo_renegociacion.html")
ruta_html.write_text(HTML, encoding="utf-8")

from playwright.sync_api import sync_playwright

salida = "/home/user/DANISANTELMO/anexo_renegociacion_persan.png"
with sync_playwright() as pw:
    navegador = pw.chromium.launch(
        executable_path="/opt/pw-browsers/chromium-1194/chrome-linux/chrome")
    pagina = navegador.new_page(viewport={"width": 1180, "height": 400},
                                device_scale_factor=2)
    pagina.goto(ruta_html.as_uri())
    pagina.wait_for_timeout(400)
    pagina.screenshot(path=salida, full_page=True)
    navegador.close()

print("Imagen:", salida)
print(f"  Contrato actual ..... {ACT['impacto']:>+12,.0f} €   margen {ACT['margen']:+.2f} €/ud")
print(f"  Renegociado ......... {REN['impacto']:>+12,.0f} €   margen {REN['margen']:+.2f} €/ud")
print(f"  Mejora .............. {REN['impacto'] - ACT['impacto']:>+12,.0f} €")
