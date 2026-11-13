# Análisis inicial

Resultado de revisar los archivos reales de octubre y noviembre de 2026. Este documento contiene solo estructura y conteos, sin datos reales.

Los scripts de esa exploración eran temporales y no se conservaron. Aquí está lo necesario para reconstruir la regla.

## Proceso actual

1. Ella descarga los movimientos del banco a su plantilla de extracto.
2. Exporta los movimientos de la cuenta desde el sistema contable de la empresa.
3. Busca los movimientos que coinciden en débito o crédito.
4. Cuando coinciden, copia el `Comprobante` en `DOC CONCILIADO` y pinta la fila de amarillo.

## Archivo de extracto

Plantilla de Excel con un archivo por mes. El nombre cambia con el mes.

**Hojas**

- **Principal:** primera hoja. Conserva un nombre de un mes anterior, así que no sirve para identificarla.
- **Otras dos:** una tiene una tabla dinámica y la otra llega a casi un millón de filas de dimensión. No intervienen en el cruce, pero deben conservarse en la copia.

**Columnas de la hoja principal, con encabezado en la fila 1**

| Columna | Encabezado | Contenido |
|---|---|---|
| A | `FECHA` | Entero con formato `AAAAMMDD` |
| B | `CONCEPTO` | Texto del banco |
| C | `DETALLE` | Texto del banco; lleno en 374 de los 377 conciliados |
| D | `DOC CONCILIADO` | Comprobante, o vacío si no está conciliado |
| E | ` DB ` | Monto; el encabezado tiene espacios |
| F | ` CR ` | Monto; el encabezado tiene espacios |
| G | ` SALDOS ` | Fórmula con el patrón `=+G{anterior}+E{fila}-F{fila}` |

**Particularidades**

- La fila 2 es el saldo inicial, con concepto `SALDO`. No es un movimiento.
- La plantilla tiene fórmulas de saldo hasta cerca de la fila 1120, en filas todavía sin fecha ni montos. Solo son movimientos las filas con fecha y monto.
- Hay un par de celdas sueltas en `H:J`, cerca de la fila 1120, ajenas a la tabla.
- Las filas conciliadas tienen relleno amarillo `FFFF00` en las columnas `A:G`.
- Tiene filtro automático y paneles congelados.

## Archivo de movimientos

Exportación del sistema contable. El nombre del archivo tampoco es uniforme.

**Hojas**

- **Noviembre:** una sola hoja, tal como sale del sistema.
- **Octubre:** dos hojas. Una tiene la exportación original con solo los primeros días (66 movimientos), y la otra es una versión limpia con el mes completo (377). La de referencia es la del mes completo.
- **Decisión del usuario:** la conciliación usa solo la primera hoja. Las demás no entran al proceso.

**Formatos de la hoja**

- **Exportación original:** cuatro filas de título combinadas, encabezado en la fila 5 y datos desde la fila 6. Al final trae una fila de totales y un pie sin `Comprobante`, que se ignoran.
- **Versión limpia:** encabezado en la fila 1 y datos desde la fila 2.
- **Cómo ubicar el encabezado:** buscar la fila que contiene `Comprobante`.

**Columnas**

`Código cuenta contable`, `Cuenta contable`, `Comprobante`, `Secuencia`, `Fecha Elaboración`, `Identificación`, `Sucursal`, `Nombre tercero`, `Descripcion`, `Detalle`, `Centro de Costo`, `Saldo inicial`, `Débito`, `Crédito`, `Saldo Movimiento`, `Saldo total cuenta`.

**Particularidades**

- `Fecha Elaboración` es texto con formato `DD-MM-AAAA`.
- `Comprobante` tiene la forma `CC-9-9`: un prefijo `CC` y dos grupos de dígitos.
- `Comprobante` puede repetirse, aunque casi nunca: hay 376 distintos en 377 filas. La clave de una fila es `Comprobante` + `Secuencia`.

## Datos de octubre

**Extracto**

- 913 movimientos: 65 con `DB` y 848 con `CR`. El análisis inicial contaba 914 porque incluía la fila de saldo inicial, que no es un movimiento.
- 377 conciliados y 536 sin conciliar.
- Las filas sin conciliar son casi todas de cobro de IVA, servicio de pagos, impuesto del gobierno e intereses. También hay un puñado sin clasificar.

**Movimientos**

- 377 filas, todas con `Comprobante`: 35 con débito y 342 con crédito.

**Relación entre ambos**

- Los 377 valores de `DOC CONCILIADO` existen en `Comprobante`.
- La relación es casi siempre de uno a uno: 373 comprobantes se usan una vez y 2 se usan dos veces. Aún no se ha revisado qué pasa en esos 2.
- Se usaron 375 de los 376 comprobantes distintos.
- 360 de 377 coinciden exacto en monto y sentido. De los otros 17, 12 difieren hasta $1, 3 entre $1 y $100, y 2 más de $100.
- La fecha coincide en 371 de 377; las otras 6 difieren en un día.
- En 365 de 377 hay al menos una palabra en común entre el texto del extracto y el de movimientos.

## Regla probada y línea base

**Regla**

1. Candidatos: movimientos del mismo lado (`DB`/`Débito` o `CR`/`Crédito`), con el monto dentro de la tolerancia y hasta 3 días de diferencia.
2. Atender primero las filas del extracto con menos candidatos.
3. Elegir el candidato con la fecha más cercana. En caso de empate, el que comparta más palabras con la fila del extracto.
4. Comparar `CONCEPTO` + `DETALLE` contra `Nombre tercero` + `Descripcion` + `Detalle` + `Identificación`. Las palabras se normalizan a mayúsculas sin tildes, con al menos 3 caracteres alfanuméricos.
5. Cada fila de movimientos se usa una sola vez.

**Resultados contra octubre (377 conciliados)**

| Tolerancia | Usa texto | Acierta | Otro comprobante | Sin propuesta | Propone en fila vacía | Con varios candidatos |
|---|---|---:|---:|---:|---:|---:|
| Exacta | No | 319 | 48 | 10 | 0 | 140 |
| Exacta | Sí | 348 | 19 | 10 | 0 | 140 |
| **$1** | **Sí** | **353** | **20** | **4** | **0** | **154** |
| $100 | Sí | 356 | 20 | 1 | 0 | 154 |

**Línea base:** tolerancia de $1 con texto, 353 de 377 (93,6 %). Se eligió sobre la de $100 para no aceptar diferencias grandes sin revisión.

**Cómo leer estos resultados**

- De los 20 con otro comprobante, 5 tienen el mismo monto y fecha que el correcto. Esos datos no alcanzan para diferenciarlos.
- No se midió cuántos de los 20 errores caen dentro de los 154 casos con varios candidatos. Es lo primero que hay que medir, porque define qué se puede marcar como seguro.
- Tener varios candidatos no significa que la propuesta sea dudosa: el texto puede resolver el empate. Hay que definir cuándo un caso es seguro.

## Prueba en seco de noviembre

Sin escribir nada, y con tolerancia de $1 y texto:

- Extracto: la prueba contó 53 movimientos, de los cuales 30 son comisiones, impuestos o intereses, y dejó 23 por conciliar. La etapa 2 mide 52 movimientos, porque ese conteo incluía la fila de saldo inicial.
- Movimientos: 22 filas. Las 22 encontraron pareja, y 8 tenían varios candidatos.
- La prueba dejó 1 fila del extracto sin pareja. Podría ser la de saldo inicial; se confirma en la etapa 3.
- Los movimientos solo cubren el 3 y 4 de noviembre, y los del extracto van del 1 al 4 de noviembre.

## Prueba de la copia (etapa 1)

**Estado:** cerrada de forma provisional. Falta que ella abra la copia en Excel y confirme que no aparece un aviso de reparación. Esa revisión debe estar lista antes de generar la salida de la etapa 5.

**Método elegido:** edición directa. Un `.xlsx` es un `.zip` con varios archivos XML, y solo se cambian la hoja principal y los estilos. Las demás partes se copian sin modificarlas.

**Estructura de los extractos**

- Tres hojas visibles y una tabla dinámica en una hoja distinta de la principal.
- La tercera hoja interna tiene 1 047 999 filas y 3 144 534 celdas definidas, de las cuales solo 2118 tienen contenido.

**Resultados con 3 filas de prueba**

La edición directa se probó con octubre y noviembre, con el mismo resultado. `openpyxl` solo se probó con octubre.

| Comprobación | `openpyxl` | Edición directa |
|---|---:|---:|
| Partes internas sin cambios | 1 de 21 | 19 de 21 |
| Saldos con valor calculado | 0 de 1115 | 1115 de 1115 |
| Tabla dinámica | Reescrita | Igual |
| Configuración de impresora | Eliminada | Se conserva |
| Tiempo de creación de la copia | 21 s | 1 s |
| Memoria máxima | 1,7 GB | 293 MB |

**Revisiones**

- **Comparador:** sin diferencias fuera de las filas de prueba, en las tres hojas.
- **Originales:** las sumas de comprobación no cambiaron.
- **LibreOffice:** abrió la copia. El usuario comprobó que la tabla dinámica funciona y que los datos se ven iguales al original.
- **Recálculo:** los 1115 saldos coinciden en LibreOffice, pero tardó solo 2 s, así que no está confirmado que los haya recalculado.

## Lectura (etapa 2)

`conciliacion/lectura.py` lee ambos archivos y no reportó filas para revisar en ninguno de los dos meses.

| Conteo | Octubre | Noviembre |
|---|---:|---:|
| Movimientos del extracto | 913 | 52 |
| Con `DB` / con `CR` | 65 / 848 | 3 / 49 |
| Ya conciliados | 377 | 0 |
| Movimientos contables | 377 | 22 |
| Con débito / con crédito | 35 / 342 | 1 / 21 |
| Comprobantes distintos | 376 | 21 |
| Claves comprobante + secuencia | 377 | 22 |

Los movimientos contables son los de la primera hoja. En octubre, los 377 valores de `DOC CONCILIADO` existen en `Comprobante`.

## Cruce (etapa 3)

`conciliacion/cruce.py` usa la regla definida por el usuario: mismo lado siempre, misma fecha y monto con hasta $1.000 de diferencia. Reemplaza la regla y la línea base del análisis inicial, que admitían $1 y 3 días.

- **Segura:** un solo candidato, que no sirve para otra fila, con monto exacto.
- **Dudosa:** varios candidatos, un candidato compartido o una diferencia de monto.
- **Elección:** entre varios candidatos, gana la menor diferencia de monto y, si empatan, la que comparte más palabras. Las filas con coincidencia exacta se atienden primero.

**Medición con octubre, vaciando `DOC CONCILIADO` en memoria**

| Clasificación | Filas | Coincide con ella | Intercambiada con una fila vecina | Movimiento que ella no usó |
|---|---:|---:|---:|---:|
| Segura | 259 | 255 | 3 | 1 |
| Dudosa | 118 | 103 | 15 | 0 |
| Pendiente | 536 | - | - | - |

- **Coincide con ella:** 358 de 377. Las otras 19 tienen propuesta: 18 son intercambios y 1 es la fila 488.
- **Filas que ella dejó vacías:** las 536 quedan pendientes, sin propuesta.
- **Filas de ella con coincidencia exacta:** 354 de 377. Monto exacto en 360 y misma fecha en 371.

**Medición anterior con la regla exacta, sin tolerancia:** 349 coincidencias, 280 seguras (9 con otro comprobante), 87 dudosas y 10 filas de ella sin propuesta.
Análisis hecho con la regla exacta:

- **Filas no exactas:** de las 23, 17 difieren en monto (12 hasta $1, 3 entre $1 y $100, 2 más de $100) y 6 difieren en un día. El script no propone nada en 10 y propone otro comprobante en 13.
- **Seguras con otro comprobante:** las 9 están entre esas 23.
  - En 8, el movimiento propuesto coincide exactamente con la fila, pero ella lo usó en otra fila.
  - De esas 8, 6 forman 3 pares de filas con los movimientos intercambiados entre sí.
  - En 1, el movimiento propuesto no lo usó en ninguna fila, y su comprobante tiene una diferencia de más de $100 y está usado en dos filas.
- **Por resolver en la etapa 4:** ella debe aclarar si esas 9 filas son errores del script o si su conciliación tuvo otra razón. El criterio exige 0 errores marcados como seguros.

## Evaluación (etapa 4)

`conciliacion/evaluacion.py` mide el mes completo y simula exportaciones parciales. Una propuesta cuenta como **inventada** si cae en una fila que ella dejó vacía, o en una fila cuyo movimiento no está en la exportación.

**Exportaciones parciales de octubre, con días completos**

| Días | Movimientos | Inventadas | Seguras que no coinciden con ella |
|---:|---:|---:|---:|
| 3 | 46 | 0 | 0 |
| 7 | 100 | 0 | 3 (filas 202, 227 y 238) |
| 15 | 187 | 0 | 3 (las mismas) |

Como la regla exige la misma fecha, quitar días completos no genera parejas inventadas.

**Movimientos faltantes al azar dentro de un mismo día** (prueba temporal, 20 repeticiones por porcentaje)

| Movimientos quitados | Total quitados | Inventadas seguras | Inventadas dudosas |
|---:|---:|---:|---:|
| 5 % | 392 | 6 | 48 |
| 10 % | 766 | 9 | 94 |
| 30 % | 2278 | 16 | 231 |

- **Inventadas seguras:** siempre son las 4 filas ya conocidas: 202, 227 y 238, que son intercambios, y 488. Cuando falta el movimiento de ella, el script toma el que coincide exactamente.
- **Inventadas dudosas:** cuando falta un movimiento, el script propone uno parecido del mismo día, con diferencia de monto o compartido con otra fila. Queda marcado como dudoso, pero el comprobante propuesto es incorrecto.

## Salida (etapa 5)

`python -m conciliacion` crea la copia marcada (`conciliacion/escritura.py`) y el archivo de revisión (`conciliacion/revision.py`).

**Comprobación con octubre:** se corrió sobre una versión temporal del extracto con `DOC CONCILIADO` vacío, borrada al terminar.

- **Tiempo:** el comando terminó en 0,9 s.
- **Resultado:** 259 seguras, 118 dudosas y 536 pendientes, igual que en la evaluación.
- **Comprobantes escritos:** 377, de los cuales 358 coinciden con los de ella.
- **Colores:** 1813 celdas en amarillo, que son 259 filas × 7 columnas, y 826 en naranja, que son 118 × 7.
- **Partes del archivo:** solo se reescribieron la hoja principal y los estilos; las otras 19 partes quedaron idénticas.
- **Contenido:** sin diferencias fuera de lo escrito en las tres hojas, y se conservaron los 1115 saldos calculados.
- **LibreOffice:** abrió la copia y leyó los 377 comprobantes.
- **Originales:** sin cambios.

## Prueba real con noviembre (etapa 6)

**Estado:** el comando ya corrió; falta que ella revise el resultado. Los archivos están en `~/Escritorio/conciliacion-salida/noviembre/`.

| Resultado | Filas |
|---|---:|
| Seguras (amarillo) | 11 |
| Dudosas (naranja) | 11, con 25 candidatos en total |
| Pendientes del extracto | 30, todas sin candidatos |
| Movimientos sin pareja | 0 de 22 |
| Filas no leídas | 0 |

- **Dudosas:** las 11 tienen varios movimientos con el mismo monto y la misma fecha. Ninguna es dudosa por diferencia de monto.
- **Pendientes:** son 30, igual que las comisiones, impuestos e intereses de la prueba en seco. Los 22 movimientos encontraron pareja, así que la fila sin pareja de la prueba en seco era la de saldo inicial.
- **Copia:** solo se reescribieron la hoja principal y los estilos, y el comando confirmó que los originales no cambiaron.

## Riesgos técnicos

- **Apertura en Excel:** aún sin comprobar. LibreOffice acepta archivos que Excel podría pedir reparar.
- **Rendimiento:** la edición directa crea la copia en 1 s. Falta medir el flujo completo.
- **Montos decimales:** vienen como `float`. Compararlos en centavos o con `Decimal`, no con `==` sobre `float`.
