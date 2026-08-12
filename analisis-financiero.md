# Análisis financiero de partida — Alcopalet (DTI-1399)

> Cálculos de apoyo para construir el informe. No es el informe: es la munición numérica para
> justificar problema, opciones, criterios y decisión. Script: `scripts/calc_alcopalet.py`.

## 1. Coherencia entre anexos (verificación, no hace falta incluirla en el informe)

- Suma de ventas por línea (2.645.224 + 1.940.042 = 4.585.266) cuadra con el P&G (4.585.267).
  Suma de resultados por línea también cuadra exacto (279.364). **Los anexos son consistentes.**
- Precio medio implícito por línea = ingreso del escandallo: nuevo 2.645.224/174.504 = 15,16 €;
  usado 1.940.042/260.330 = 7,45 €. Coincide con el Anexo 3.
- Materia prima total del escandallo (2.692.732 €) ≈ consumo real del P&G (Aprovisionamientos −
  Variación existencias = 2.692.444 €). Transporte y otros costes también cuadran casi al euro.
- La MOD del escandallo (645.955 €) **no** cuadra con "Personal-Producción" del P&G (481.660 €),
  pero si se traduce a operarios: nuevo = 10 operarios × 38.000 € = 380.000 €; usado = 7 × 38.000
  € = 266.000 €; total 646.000 €. Es decir, el escandallo usa el **coste total de personal
  directo** (incluye la fabricación del usado), mientras que "Personal-Producción" del P&G
  probablemente ya neta alguna partida o clasifica distinto. No es un error del caso: son dos
  agregaciones distintas de la misma plantilla. Útil para explicar de dónde salen los "8
  operarios manual + 2 línea automática = 10" y los "7 operarios usado".

## 2. Capacidad productiva actual (línea nueva) — el dato que falta explicitar en el caso

Supuestos: 242 días productivos/año (11 meses × 22 días), jornada 8h (7:00-15:00).

| Recurso | Cálculo | Capacidad anual |
|---|---|---|
| Línea automática 2019 | 65 p/h × 75% × 8h × 242 días | **94.380 palets/año** |
| Fabricación manual | 500 p/día × 242 días | **121.000 palets/año** |
| **Total actual** | | **215.380 palets/año** |

Producción real 2025 (línea nueva): 174.504 → **utilización 81%**, con ~40.900 palets/año de
holgura ya disponible hoy, sin ninguna inversión.

**Robot solo**: 400 p/día × 90% × 242 días = **87.120 palets/año** con 1 operario.

- Si Persán se suma a la demanda actual (174.504 + 90.000 = 264.504), la capacidad actual
  (215.380) se queda corta en ~49.000 palets/año — **hace falta capacidad nueva de todos
  modos**, con o sin robot.
- El robot solo (87.120) **cubre de sobra el bloque Persán (90.000 casi exacto: -2.880, -3,2%)**
  pero deja el resto de la holgura para seguir creciendo. Sumado a la capacidad actual (215.380 +
  87.120 = 302.500) cubre los 264.504 con margen (~38.000 palets/año de colchón).
- **No hace falta despedir a nadie para operar el robot**: con la demanda total proyectada, los
  10 operarios actuales de la línea nueva siguen ocupados a plena capacidad; el robot requiere
  1 operario adicional (contratación, no despido). El dato de "coste de despido 30.000 €" que da
  el caso parece un distractor: solo sería relevante si se decidiera sustituir la línea manual en
  vez de añadir capacidad — no es la única lectura posible, pero conviene no asumir despidos sin
  justificarlo.

## 3. El contrato con Persán, aislado, da margen de contribución NEGATIVO

Directamente del Anexo 3: margen de contribución unitario del "Proyecto Persán" = **-0,52
€/palet**. Esto es **antes** de sumar campa y leasing — es decir, ni siquiera cubre los costes
variables al precio pactado de 10 €.

Causa: la **materia prima (madera) es 8,27 €, el 82,7% del precio de venta**, y ese coste es
igual al de la fabricación manual (misma calidad de madera exigida, probablemente heredada del
estándar Tetra Pak). El robot ahorra mano de obra (0,42 € vs 2,18 €) pero **no toca el coste de
la madera**, que es la partida dominante. Cuantificado a volumen (90.000 ud/año):

| Concepto | Importe/año |
|---|---|
| Ingresos (90.000 × 10 €) | +900.000 € |
| Costes variables (90.000 × 10,52 €) | -946.800 € |
| **Margen de contribución** | **-46.800 €** |
| Alquiler campa (5.000 €/mes) | -60.000 € |
| Leasing máquina + camión (4.000 €/mes) | -48.000 € |
| **Subtotal antes de circulante** | **-154.800 €** |
| Coste financiero del circulante (cobro a 180 días, 7,5%) | -33.288 € |
| **IMPACTO ANUAL TOTAL ESTIMADO** | **≈ -188.000 €/año** |

Eso equivale a **-67% del resultado neto actual de toda la empresa (279.364 €)**, y a lo largo de
los 5 años del contrato, del orden de **-940.000 €** si nada cambia. (Cálculo de circulante:
cuenta a cobrar media ≈ ingresos × 180/365 = 443.836 €, coste financiero anual = 443.836 × 7,5%
= 33.288 €. Es una aproximación de bolsillo, no un análisis de tesorería completo — conviene
presentarla como orden de magnitud, no como cifra cerrada.)

**Precio de equilibrio** (para que Persán no reste valor):
- Solo variables: 10,52 €/ud
- + campa y leasing: 11,72 €/ud
- + coste de circulante: ≈12,09 €/ud
- Para igualar el margen actual del negocio (2,80 €/ud): ≈14,52 €/ud

Persán ofrece 10 €/ud fijo 5 años → **está entre 2 y 4,5 € por debajo de lo que necesitaría
Alcopalet**, y ese precio queda congelado mientras la madera (82,7% del coste) puede subir. Una
subida del 10% en el precio de la madera empeoraría el resultado en ~121.000 €/año adicionales.

## 4. El robot es un excelente proyecto por sí mismo — separado de Persán

Sustituir manual por robot en términos de mano de obra: 87.120 palets/año que en manual
requerirían 87.120/(500×242 en 8 operarios) ≈ 5,76 operarios equivalentes. Con solo 1 operario en
el robot:

- Ahorro bruto de mano de obra: (5,76 - 1) × 38.000 € ≈ **181.000 €/año**
- Coste del leasing solo de la máquina (40.000 € al 7,5%, 5 años) ≈ 9.600 €/año
- **Ahorro neto ≈ 171.000 €/año, payback sobre la inversión de 40.000 € en menos de 3 meses**

Y si esa capacidad (87.120 palets) se colocara a precio de mercado normal (15,16 €) en vez de al
precio Persán (10 €), el margen sería +4,55 €/ud en vez de -0,52 €/ud: una diferencia de **más de
440.000 €/año** en la misma capacidad física. **La máquina es buena decisión; el precio pactado
con Persán es el problema, no la tecnología.**

## 5. La petición de los padres: ¿hay caja para dividendos?

- EBITDA 2025: 452.674 €, pero una parte importante está "atrapada" en el aumento de existencias
  (376.062 €) — es decir, gran parte del EBITDA no es caja libre todavía.
- Caja operativa aproximada 2025 (EBITDA - Δexistencias - financieros - impuestos) ≈ **-60.000 €**,
  frente a un resultado neto contable positivo de 279.364 €. Son magnitudes distintas: el
  beneficio contable no es lo mismo que la caja disponible para repartir.
- El propio proyecto Persán, en su primer año, exige del orden de **-600.000 €** de caja
  (circulante por cobro a 180 días + campa/leasing + margen negativo) antes de generar ningún
  retorno.
- **Conclusión de esta pieza**: la petición de los padres de "sacar liquidez ahora" choca
  directamente con las necesidades de caja que exige aceptar Persán tal como está pactado. No son
  dos decisiones independientes: cuanto más se comprometa la empresa con Persán en estas
  condiciones, menos margen habrá para dividendos, y viceversa.

## 6. Palet Ojeda (usado): recorrido de crecimiento bloqueado por espacio, no por demanda

- Producción actual: 260.330 ud/año (21.694/mes) contra un techo de 30.000/mes (360.000/año).
- Recorrido disponible: ~99.670 ud/año adicionales, que a su margen actual (0,43 €/ud) aportarían
  ~43.000 €/año — modesto, pero "gratis" en el sentido de que ya hay demanda, solo falta espacio.
- La campa que se alquilaría para el retén de Persán (5.000 €/mes = 60.000 €/año) es precisamente
  el recurso que falta para esta línea. **Punto de creatividad para el informe**: ¿puede la
  misma nave/campa nueva servir simultáneamente al retén de camiones de Persán y ampliar la
  capacidad de Palet Ojeda? Si es así, parte del coste de la campa (que hoy solo carga sobre un
  contrato con margen negativo) se repartiría también sobre una línea que sí genera margen
  positivo, con demanda adicional real (recorrido de +99.670 ud/año). Es una hipótesis a
  contrastar, no un dato del caso, pero es exactamente el tipo de alternativa creativa que pide
  el criterio de evaluación "creatividad para generar soluciones y alternativas".

## Resumen: el verdadero problema no es "máquina sí o no"

Los datos apuntan a que la pregunta que el caso pone en la superficie — "¿compramos el robot?" —
tiene una respuesta financiera bastante clara en sí misma (sí, con payback de meses). El problema
real, mucho menos obvio, es que **el contrato con Persán, tal como está redactado (10 €/ud fijos
a 5 años, retén de 2 camiones, cobro a 180 días), destruye valor incluso con la máquina más
eficiente posible**, porque el coste que domina el precio (la madera, 83%) no lo resuelve ninguna
máquina. Eso separa con nitidez:

- **Problema principal**: si firmar Persán en las condiciones actuales compromete la rentabilidad
  y la liquidez de la empresa durante 5 años, en un momento en que además los fundadores piden
  liquidez para su jubilación.
- **Decisión de la máquina**: en principio favorable de forma casi independiente de Persán.
- **Alternativas reales a explorar**: renegociar precio/condiciones con Persán (cláusula de
  revisión por precio de madera, PV más alto, menor volumen de retén), buscar otro proveedor o
  mezcla de proveedores de madera, dimensionar la inversión de forma distinta (¿hace falta el
  tráiler de 150.000 € desde el día uno?), y la posible sinergia campa/Palet Ojeda del punto 6.
