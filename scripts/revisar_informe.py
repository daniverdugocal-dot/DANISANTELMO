"""
Regenera el informe de Alcopalet con las correcciones marcadas (subrayadas)
y le añade las páginas de anexos del PDF original, sin modificarlas.

Salida: Alcopalet_Daniel_Verdugo_Calvo_REVISADO.pdf
"""
import base64
from pathlib import Path

import pymupdf
from playwright.sync_api import sync_playwright

ORIGINAL = "/root/.claude/uploads/9303fdae-b74e-58a3-9bb1-e752b7e896c0/233c5763-AlcopaletDaniel_Verdugo_Calvo.pdf"
SALIDA = "/home/user/DANISANTELMO/Alcopalet_Daniel_Verdugo_Calvo_REVISADO.pdf"
TMP = Path("/tmp/rev_cuerpo.pdf")

# ---------------------------------------------------------------- logo
doc = pymupdf.open(ORIGINAL)
# Se recorta el logo de la página renderizada: extraerlo como imagen suelta
# pierde la máscara de transparencia y sale con fondo negro.
_pag = doc[2]
_rect = _pag.get_image_rects(_pag.get_images(full=True)[0][0])[0]
_pix = _pag.get_pixmap(clip=_rect, dpi=400, colorspace=pymupdf.csRGB, alpha=False)
LOGO_B64 = base64.b64encode(_pix.tobytes("png")).decode()


def c(t):
    """Marca un cambio: subrayado."""
    return f'<u class="cambio">{t}</u>'


# ---------------------------------------------------------------- contenido
INTRO = (
    "Alcopalet es una empresa familiar dedicada a la fabricación y reparación de palés, "
    "actualmente dirigida por Esteban Ojeda. Tras años de crecimiento, la compañía se "
    "enfrenta a una nueva oportunidad con el contrato de Persán, que exige aumentar su "
    "capacidad productiva y logística. Para afrontarlo, Esteban plantea invertir en una "
    "nueva máquina que aporte mayor flexibilidad y reduzca la dependencia de la mano de "
    "obra. Sin embargo, sus padres, fundadores de la empresa, desean priorizar el reparto "
    "de beneficios de cara a su jubilación. Esteban debe decidir cómo afrontar esta "
    "oportunidad de crecimiento sin comprometer la estabilidad y el futuro de Alcopalet."
)

RESUMEN = (
    "Alcopalet se encuentra ante una oportunidad de crecimiento que exige decidir cómo "
    "afrontar el contrato con Persán sin comprometer la rentabilidad y la estabilidad de "
    "la empresa. Tras analizar las distintas alternativas, se propone adquirir el robot y "
    "renegociar el contrato, estableciendo como objetivo un precio de 12,50 € por palé y "
    "un plazo de cobro de 90 días. Con estas condiciones, el contrato generaría un "
    "resultado anual positivo de 48.606 €, frente a los -188.877 € de la propuesta "
    "inicial. Si Persán no acepta la renegociación, Alcopalet deberá rechazar el contrato "
    "y mantener la inversión en el robot, orientando la mejora de capacidad y "
    "productividad hacia el crecimiento rentable del negocio."
)

PROBLEMA = (
    "El contrato con Persán, tal como está planteado, no cubre su coste: 10 € de precio "
    "frente a 10,52 € de coste unitario, lo que resta 188.877 € anuales "
    + c("(véanse los Anexos 3 y 6)") + ". A ello se añade que la capacidad productiva y "
    "logística actual no permite asumir el nuevo volumen sin nuevas inversiones "
    + c("(véase el Anexo 4)") + ", y el deseo de los fundadores de comenzar a recoger los "
    "frutos del negocio. Esteban debe decidir si apostar por el crecimiento y cómo hacerlo "
    "sin comprometer la rentabilidad, la liquidez y la estabilidad de la empresa."
)

OPC_INTRO = ("Ante la situación planteada, Esteban dispone de cuatro alternativas para "
             "afrontar el crecimiento de Alcopalet:")

OPC1 = (
    "<b>Aceptar el contrato de Persán en las condiciones actuales.</b> Esta alternativa "
    "permitiría incorporar un cliente de gran volumen, aumentar las ventas y reforzar la "
    "posición de Alcopalet en el mercado. El contrato ofrece estabilidad comercial durante "
    "cinco años, aunque fija el precio de venta y establece un plazo de cobro de 180 días. "
    "El coste unitario estimado asciende a 10,52 €, frente a un precio de 10 €, lo que "
    "genera un margen negativo de 0,52 € por palé " + c("(véase el Anexo 3)") + ". Al "
    "incluir el alquiler de la campa, el leasing y el coste del circulante, el impacto "
    "anual estimado sería de -188.877 € " + c("(véase el Anexo 6)") + "."
)

OPC2 = (
    "<b>Rechazar el contrato y mantener la situación actual.</b> Esta opción permitiría "
    "conservar la estructura actual, evitar nuevas inversiones y preservar la rentabilidad "
    "y la liquidez alcanzadas. También facilitaría el reparto de beneficios solicitado por "
    "los fundadores. Alcopalet obtuvo un resultado neto de 279.364 € y un margen neto del "
    "6,1 % en 2025 " + c("(véase el Anexo 1)") + ". Al mismo tiempo, la empresa "
    "renunciaría al crecimiento asociado a Persán y continuaría operando con las "
    "limitaciones actuales de capacidad y espacio."
)

OPC3 = (
    "<b>Adquirir el robot como inversión independiente, sin vincularlo inicialmente a "
    "Persán.</b> El robot permitiría elevar la productividad, reducir el coste de mano de "
    "obra por palé y disponer de una mayor flexibilidad en los cambios de formato. Su "
    "ahorro teórico neto asciende a 171.262 € anuales y el plazo estimado de recuperación "
    "es de 2,65 meses " + c("(véase el Anexo 7)") + ". Como Alcopalet no contempla "
    "despedir trabajadores, la inversión requeriría recolocar al personal liberado en "
    "otras actividades para transformar la mejora productiva en un beneficio real. Una "
    "posibilidad sería reforzar la línea de palé usado, que todavía dispone de recorrido "
    "de crecimiento, aunque está condicionada por el espacio disponible "
    + c("(véase el Anexo 5)") + "."
)

OPC4 = (
    "<b>Renegociar las condiciones con Persán antes de aceptar el contrato.</b> Esta "
    "alternativa consistiría en proponer un precio de 12,50 € por palé y reducir el plazo "
    "de cobro a 90 días. Bajo estas condiciones, el margen de contribución alcanzaría "
    "1,98 € por unidad y el contrato generaría un resultado anual positivo de 48.606 €, "
    "mejorando en 237.483 € el escenario inicial " + c("(véase el Anexo 8)") + ". La "
    "opción combinaría el crecimiento comercial con la modernización productiva y una "
    "menor presión sobre la liquidez, aunque su puesta en marcha dependería de la "
    "aceptación de Persán y de la adecuada coordinación de las inversiones y los recursos "
    "necesarios."
)

CRIT_INTRO = (
    "Para seleccionar la alternativa más adecuada, Alcopalet debe valorar no solo la "
    "oportunidad de crecimiento que representa Persán, sino también las consecuencias que "
    "cada opción puede tener sobre el conjunto de la empresa. La decisión se analizará "
    "desde tres perspectivas complementarias: económico-financiera, estratégica y de "
    "mercado, y organizativa, comercial y humana."
)

CRIT1 = (
    "Desde el punto de vista <b>económico-financiero</b>, se tendrá en cuenta la "
    "rentabilidad que puede aportar cada alternativa, los recursos necesarios para "
    "llevarla a cabo y su efecto sobre la liquidez de la compañía. Este aspecto resulta "
    "especialmente relevante ante las necesidades de circulante que genera Persán y el "
    "impacto que tendría el contrato sobre los resultados de Alcopalet, analizados "
    + c("en el Anexo 6") + ". También se valorará que el crecimiento no deteriore la "
    "rentabilidad que actualmente genera la empresa."
)

CRIT2 = (
    "En la dimensión <b>estratégica y de mercado</b>, se analizará en qué medida cada "
    "opción contribuye al crecimiento futuro de Alcopalet y permite incorporar clientes "
    "que aporten volumen y rentabilidad. Asimismo, se tendrá en cuenta la capacidad "
    "disponible para acompañar ese crecimiento, recogida " + c("en el Anexo 4") + ", y el "
    "riesgo de que Persán adquiera un peso excesivo dentro del negocio."
)

CRIT3 = (
    "Finalmente, desde la perspectiva <b>organizativa, comercial y humana</b>, se valorará "
    "la capacidad para asumir un mayor volumen manteniendo el nivel de servicio a los "
    "clientes actuales, así como los cambios necesarios en producción y logística. También "
    "será importante aprovechar adecuadamente la plantilla sin recurrir a despidos, "
    "especialmente ante la capacidad de trabajo que podría liberar la automatización, "
    "analizada " + c("en el Anexo 7") + ", y compatibilizar la decisión con los intereses "
    "de Esteban y de los fundadores."
)

REC_INTRO = (
    "Una vez analizadas las cuatro alternativas conforme a los criterios establecidos, se "
    "descartan aquellas que no ofrecen un equilibrio adecuado entre crecimiento, "
    "rentabilidad y estabilidad. El rechazo de estas opciones permite centrar la decisión "
    "en la alternativa que mejor responde a la situación actual y al futuro de Alcopalet:"
)

REC1 = (
    "<b>Rechazo de la alternativa 1</b>" + c(",") + " consistente en aceptar el contrato "
    "de Persán en las condiciones actuales. Aunque permitiría aumentar las ventas y "
    "asegurar un volumen de actividad durante cinco años, se rechaza porque supondría "
    "crecer destruyendo valor. El precio de 10 € por palé es inferior al coste unitario de "
    "10,52 € y, al considerar las inversiones adicionales y las necesidades de circulante, "
    "el impacto anual estimado asciende a -188.877 € " + c("(véase el Anexo 6)") + ". "
    + c("El precio que equilibraría el contrato asciende a 12,17 € por palé, muy por "
        "encima del pactado.") + " Por tanto, no cumple los criterios de rentabilidad y "
    "liquidez establecidos."
)

REC2 = (
    "<b>Rechazo de la alternativa 2</b>" + c(",") + " consistente en mantener la situación "
    "actual. Esta opción permitiría preservar los resultados actuales y facilitar el "
    "reparto de beneficios a los fundadores, pero no resuelve las limitaciones que pueden "
    "frenar el crecimiento futuro de Alcopalet. La empresa mantendría sus restricciones de "
    "capacidad y espacio y renunciaría tanto a Persán como a otras oportunidades "
    "comerciales. Además, la línea de palé usado "
    + c("presenta un margen neto superior al del palé nuevo (6,21 % frente a 6,01 %) y ")
    + "todavía dispone de un recorrido potencial de 99.670 unidades anuales "
    + c("(véanse los Anexos 2 y 5)") + ". Se rechaza, por tanto, por su escaso potencial "
    "de crecimiento y por no preparar a la empresa para afrontar mayores volúmenes."
)

REC3 = (
    "<b>Rechazo de la alternativa 3</b>" + c(",") + " consistente en adquirir el robot sin "
    "vincularlo inicialmente a Persán. La alternativa es favorable desde el punto de vista "
    "productivo, ya que mejora la flexibilidad y productividad y presenta un plazo "
    "estimado de recuperación de 2,65 meses " + c("(véase el Anexo 7)") + ". Sin embargo, "
    "se descarta como primera opción porque no aprovecha plenamente la oportunidad "
    "comercial que representa Persán si existe la posibilidad de alcanzar unas condiciones "
    "rentables. No obstante, se mantendría como alternativa de respaldo si la negociación "
    "con Persán no prospera."
)

REC_FIN = (
    "Por tanto, descartadas las alternativas anteriores, la mejor opción es invertir en el "
    "robot y renegociar el contrato con Persán antes de aceptarlo. Esta alternativa "
    "permite aprovechar la oportunidad de crecimiento sin asumir las condiciones "
    "económicas iniciales, combinando la mejora de la capacidad productiva con un contrato "
    "rentable y una menor presión sobre la liquidez. Bajo las condiciones propuestas, el "
    "contrato generaría un resultado anual positivo de 48.606 € "
    + c("(véase el Anexo 8)") + ", frente a las pérdidas que produciría el acuerdo inicial."
)

CONS1 = (
    "La alternativa elegida permite afrontar el crecimiento de Alcopalet sin asumir las "
    "pérdidas derivadas de las condiciones iniciales de Persán. Con la renegociación "
    "propuesta (12,50 € por palé y cobro a 90 días), el contrato pasaría de generar un "
    "impacto anual de -188.877 € a aportar un resultado positivo de 48.606 €, lo que "
    "supone una mejora de 237.483 € anuales " + c("(véase el Anexo 8)") + ". Además, la "
    "inversión en el robot mejoraría la productividad y la flexibilidad productiva, con un "
    "plazo estimado de recuperación de 2,65 meses " + c("(véase el Anexo 7)") + "."
)

CONS2 = (
    "Más allá del efecto financiero, la decisión permitiría aumentar la capacidad y "
    "reducir la dependencia de la mano de obra. Al no contemplarse despidos, el personal "
    "liberado deberá recolocarse en otras actividades con potencial de crecimiento, como "
    "la línea de palé usado. Comercialmente, Persán aportaría un volumen estable durante "
    "cinco años, aunque Alcopalet deberá evitar una dependencia excesiva y mantener el "
    "nivel de servicio a sus clientes actuales. "
    + c("En cuanto a los fundadores, la mejora del resultado permitiría abordar el reparto "
        "de beneficios de forma progresiva, una vez que el contrato renegociado empiece a "
        "generar caja.")
)

CONS3 = (
    "Si Persán no acepta la renegociación, Alcopalet debería rechazar el contrato y "
    "mantener la inversión en el robot como una decisión independiente. De este modo "
    "evitaría asumir un contrato deficitario, pero conservaría las mejoras de "
    "productividad, capacidad y flexibilidad necesarias para atender el crecimiento futuro "
    "y captar nuevos clientes en mejores condiciones. La inversión dejaría así de depender "
    "de Persán y pasaría a formar parte de la modernización de la estructura productiva de "
    "Alcopalet."
)

ESF_INTRO = ("La alternativa propuesta exigirá un esfuerzo coordinado en distintas áreas "
             "de Alcopalet:")
ESF1 = ("En el ámbito <b>comercial</b>, la prioridad será renegociar con Persán unas "
        "condiciones que aseguren la rentabilidad del contrato, especialmente el precio y "
        "el plazo de cobro, sin descuidar la cartera actual ni generar una dependencia "
        "excesiva de este cliente.")
ESF2 = ("En el área <b>productiva y logística</b>, será necesario incorporar el robot, "
        "adaptar la planificación de la producción y asegurar los medios logísticos y el "
        "espacio necesarios para cumplir los compromisos de servicio.")
ESF3 = ("Por último, en el ámbito <b>humano y organizativo</b>, la automatización "
        "permitirá liberar aproximadamente cinco operarios, que deberán ser recolocados en "
        "tareas donde puedan aportar mayor valor, especialmente apoyando el crecimiento de "
        "la línea de palé usado.")
ESF_FIN = ("De esta forma, la inversión no se plantea como una reducción de plantilla, "
           "sino como una oportunidad para aumentar la productividad y aprovechar mejor "
           "los recursos actuales de Alcopalet.")

DECISION = (
    "Alcopalet debe invertir en el robot y renegociar el contrato con Persán, fijando como "
    "objetivo un precio de 12,50 € por palé y un plazo de cobro de 90 días. Si Persán no "
    "acepta estas condiciones, deberá rechazarse el contrato, manteniendo la inversión en "
    "el robot como apuesta por la productividad y el crecimiento futuro de la empresa."
)

PLAN_INTRO = (
    "La ejecución de la propuesta se realizará de forma progresiva, evitando que la "
    "incorporación del robot y, en su caso, del contrato con Persán afecten al "
    "funcionamiento habitual de Alcopalet."
)

FASE1 = (
    "<b>Fase 1. Negociación y toma de decisión.</b> Esteban deberá confirmar la "
    "adquisición del robot y negociar con Persán las nuevas condiciones del contrato, "
    "estableciendo como referencia un precio de 12,50 € por palé y un plazo de cobro de 90 "
    "días" + c(", con un límite de retirada de 12,17 € por palé") + ". La aceptación del "
    "contrato quedará condicionada a alcanzar unas condiciones que garanticen su "
    "rentabilidad. Si Persán mantiene la propuesta inicial, Alcopalet rechazará el acuerdo "
    "y continuará con la inversión en el robot de manera independiente."
)

FASE2 = (
    "<b>Fase 2. Implantación productiva y logística.</b> Una vez recibido el robot, se "
    "procederá a su instalación, puesta a punto y realización de pruebas con los diferentes "
    "formatos antes de incorporarlo plenamente a la producción. Se designará un operario "
    "responsable, que recibirá la formación necesaria, y durante el primer mes se "
    "controlarán diariamente la producción, las paradas, los tiempos de cambio y las "
    "incidencias. Si finalmente se alcanza un acuerdo con Persán, se contratará la campa, "
    "se dispondrá de los medios de transporte necesarios y se organizará el stock "
    "permanente de 1.000 palés para garantizar el nivel de servicio exigido."
)

FASE3 = (
    "<b>Fase 3. Reorganización de los recursos.</b> Tras estabilizar el funcionamiento del "
    "robot, se reorganizará la plantilla manteniendo el compromiso de no realizar despidos. "
    "Los trabajadores cuya carga de trabajo se reduzca serán reasignados progresivamente a "
    "otras actividades, especialmente a la línea de palé usado. Paralelamente, deberá "
    "revisarse la distribución del espacio disponible para aliviar el actual cuello de "
    "botella logístico y aprovechar el potencial de crecimiento de esta línea."
)

FASE4 = (
    "<b>Fase 4. Seguimiento y control.</b> Esteban implantará un cuadro de control mensual "
    "que recoja producción, coste por palé, margen por cliente, utilización de la "
    "capacidad, entregas a tiempo y días de cobro. A los tres y seis meses se realizará una "
    "revisión específica para comprobar la rentabilidad real de Persán, las mejoras de "
    "productividad obtenidas con el robot y el aprovechamiento del personal recolocado. "
    "Cualquier desviación relevante deberá traducirse en medidas concretas sobre precios, "
    "costes, producción o asignación de recursos."
)

FASE5 = (
    "<b>Fase 5. Consolidación del crecimiento.</b> Una vez estabilizada la nueva "
    "estructura, Alcopalet deberá utilizar la capacidad adicional para seguir desarrollando "
    "su cartera de clientes y la línea de palé usado, evitando concentrar su crecimiento en "
    "Persán. Si finalmente no se alcanza un acuerdo con este cliente, la capacidad liberada "
    "por el robot se orientará directamente a captar nuevos clientes rentables y aumentar "
    "el volumen de las líneas actuales, manteniendo la inversión como base para el "
    "crecimiento futuro de la empresa."
)

# ---------------------------------------------------------------- HTML
def sec(n, t):
    return f'<div class="sec">{n}. {t}</div>'


def p(t, sangria=True):
    return f'<p class="{"ind" if sangria else "noind"}">{t}</p>'


def li(t, num=None):
    marca = f'<span class="num">{num}.</span> ' if num else '<span class="num">-</span> '
    return f'<p class="li">{marca}{t}</p>'


HTML = f"""<!doctype html><meta charset="utf-8">
<style>
  @page {{ size: A4; margin: 3.1cm 2.5cm 2.2cm 2.5cm; }}
  body {{ font-family: Carlito, Calibri, "Liberation Sans", sans-serif;
          font-size: 11pt; line-height: 1.32; color: #000; margin: 0; text-align: justify; }}
  .sec {{ background: #dbead5; font-weight: bold; font-size: 11.5pt;
          letter-spacing: .7px; padding: 4pt 8pt; margin: 13pt 0 7pt;
          text-align: left; }}
  p {{ margin: 0 0 8pt; }}
  p.ind {{ text-indent: 1.25cm; }}
  p.li {{ margin-left: 0; }}
  .num {{ font-weight: normal; }}
  u.cambio {{ text-decoration: underline; text-underline-offset: 2px; }}
  .sub {{ font-weight: bold; margin: 9pt 0 6pt; text-align: left; }}
  .portada {{ text-align: center; }}
  .portada .t1 {{ font-size: 26pt; margin-top: 5cm; }}
  .portada .t2 {{ font-size: 24pt; font-weight: bold; margin-top: .5cm; }}
  .portada .t3 {{ font-size: 15pt; font-style: italic; margin-top: .4cm; }}
  .portada .t4 {{ font-size: 11pt; font-style: italic; margin-top: 1.4cm; }}
  .portada .t5 {{ font-size: 11pt; font-style: italic; margin-top: .3cm; }}
  .portada .aut {{ font-size: 11pt; margin-top: 5cm; }}
  .idx {{ text-align: left; }}
  .idx div {{ margin-bottom: 6pt; }}
  .salto {{ page-break-after: always; }}
</style>

<div class="portada salto">
  <div class="t1">CASO PRÁCTICO</div>
  <div class="t2">ALCOPALET</div>
  <div class="t3">El precio del crecimiento</div>
  <div class="t4">Programa LYDES – SAN TELMO Business School</div>
  <div class="t5">31 de agosto de 2026</div>
  <div class="aut">Trabajo realizado por:<br>Daniel Verdugo Calvo</div>
</div>

<div class="salto">
  <div class="sec">CONTENIDO</div>
  <div class="idx">
    <div>1. Introducción: Breve resumen del “argumento” del Caso. ........................................ 3</div>
    <div>2. Resumen ejecutivo. ................................................................................................ 3</div>
    <div>3. Identificación del problema planteado. ................................................................. 3</div>
    <div>4. Presentación y análisis de opciones. ...................................................................... 3</div>
    <div>5. Fijación de criterios. ............................................................................................... 4</div>
    <div>6. Elección justificada de una opción. ........................................................................ 5</div>
    <div>7. Conclusión. ............................................................................................................ 6</div>
    <div>8. Anexos. ................................................................................................................. 8</div>
  </div>
</div>

{sec(1, 'INTRODUCCIÓN: BREVE RESUMEN DEL “ARGUMENTO” DEL CASO.')}
{p(INTRO)}
{sec(2, 'RESUMEN EJECUTIVO.')}
{p(RESUMEN)}
{sec(3, 'IDENTIFICACIÓN DEL PROBLEMA PLANTEADO.')}
{p(PROBLEMA)}
{sec(4, 'PRESENTACIÓN Y ANÁLISIS DE OPCIONES.')}
{p(OPC_INTRO)}
{li(OPC1, 1)}
{li(OPC2, 2)}
{li(OPC3, 3)}
{li(OPC4, 4)}
{sec(5, 'FIJACIÓN DE CRITERIOS.')}
{p(CRIT_INTRO)}
{li(CRIT1)}
{li(CRIT2)}
{li(CRIT3)}
{sec(6, 'ELECCIÓN JUSTIFICADA DE UNA OPCIÓN.')}
<div class="sub">a) Rechazo justificado de otras.</div>
{p(REC_INTRO)}
{p(REC1)}
{p(REC2)}
{p(REC3)}
{p(REC_FIN)}
<div class="sub">b) Resumen de las consecuencias financieras y de otro tipo.</div>
{p(CONS1)}
{p(CONS2)}
{p(CONS3)}
<div class="sub">c) Esfuerzo en áreas determinadas: comercial, productiva, personal, …</div>
{p(ESF_INTRO)}
{li(ESF1)}
{li(ESF2)}
{li(ESF3)}
{p(ESF_FIN)}
{sec(7, 'CONCLUSIÓN.')}
<div class="sub">a) Decisión/Propuesta Final</div>
{p(DECISION)}
<div class="sub">b) Plan de Acción: Resumen de la secuenciación temporal de actuaciones.</div>
{p(PLAN_INTRO)}
{p(FASE1)}
{p(FASE2)}
{p(FASE3)}
{p(FASE4)}
{p(FASE5)}
"""

Path("/tmp/rev_cuerpo.html").write_text(HTML, encoding="utf-8")

CABECERA = (f'<div style="width:100%; padding:0 2.5cm; text-align:right;">'
            f'<img src="data:image/png;base64,{LOGO_B64}" style="height:34px;"></div>')
PIE = ('<div style="width:100%; font-family:Carlito,Calibri,sans-serif; font-size:9pt; '
       'text-align:center;"><span class="pageNumber"></span></div>')

with sync_playwright() as pw:
    nav = pw.chromium.launch(
        executable_path="/opt/pw-browsers/chromium-1194/chrome-linux/chrome")
    pag = nav.new_page()
    pag.goto(Path("/tmp/rev_cuerpo.html").as_uri())
    pag.wait_for_timeout(400)
    pag.pdf(path=str(TMP), format="A4", print_background=True,
            display_header_footer=True, header_template=CABECERA, footer_template=PIE,
            margin={"top": "3.1cm", "bottom": "2.2cm", "left": "0", "right": "0"})
    nav.close()

# ---------------------------------------------------------------- unir con anexos
cuerpo = pymupdf.open(str(TMP))
final = pymupdf.open()
final.insert_pdf(cuerpo)
final.insert_pdf(doc, from_page=7, to_page=9)   # páginas 8, 9 y 10 originales
final.save(SALIDA)
print(f"Cuerpo generado: {cuerpo.page_count} páginas")
print(f"Anexos añadidos: 3 páginas del original")
print(f"TOTAL: {final.page_count} páginas -> {SALIDA}")
