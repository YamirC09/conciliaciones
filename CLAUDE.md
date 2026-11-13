# CLAUDE.md

Script local en Python que automatiza la conciliación bancaria diaria de una contadora, novia del usuario. Es un proyecto de práctica, sin cobro. Responde siempre en español neutro.

## Objetivo

Cruzar dos archivos de Excel y entregar una copia del extracto con la columna `DOC CONCILIADO` llena y marcada, para que ella solo revise lo dudoso:

- **Extracto:** movimientos que ella descarga del banco y acumula en su plantilla.
- **Movimientos:** exportación del sistema contable de la empresa, con la columna `Comprobante`.

## Antes de empezar

Lee `docs/analisis-inicial.md`. Tiene la estructura de los archivos, las reglas observadas y los resultados contra los que se mide el trabajo.

## Datos reales

- Están en `~/Escritorio/mi-sistema/docs/`, fuera de este repositorio. Si no tienes acceso, pide abrir la sesión con `claude --add-dir ~/Escritorio/mi-sistema/docs`.
- El usuario decidió usarlos para desarrollar y evaluar, porque el script debe funcionar con ellos.
- Los archivos de octubre ya están conciliados y son la referencia. Los de noviembre están sin conciliar.
- Nunca modifiques esos archivos. Son solo de lectura.
- No los copies a este repositorio.
- No imprimas valores reales: ni nombres, identificaciones, comprobantes, montos ni números de cuenta. Para analizar, muestra columnas, conteos, porcentajes y formatos enmascarados.
- Los nombres de los archivos contienen datos de la cuenta. No los escribas en código, pruebas ni documentación. El script los recibe como argumentos.
- Las pruebas unitarias usan solo datos inventados, generados dentro de cada prueba.
- Si aparece un dato sensible en la conversación, avisa y no lo repitas.

## Orden de trabajo

Revisa cada etapa con el usuario antes de pasar a la siguiente.

1. **Prueba de la copia:** comprobar que abrir y guardar el extracto no daña su contenido. Tiene una tabla dinámica, fórmulas y una hoja de casi un millón de filas. Si `openpyxl` altera algo, buscar otra forma de escribir antes de seguir.
2. **Lectura:** leer y normalizar ambos archivos.
3. **Cruce:** implementar las reglas con pruebas unitarias.
4. **Evaluación:** medir contra octubre y comparar con la línea base de `docs/analisis-inicial.md`.
5. **Salida:** generar la copia marcada y la lista de pendientes.
6. **Prueba real:** correr con noviembre sin tocar los originales, para que ella revise el resultado.
7. **Entrega:** definir cómo lo ejecuta ella en su equipo y escribir las instrucciones de uso.

## Decisiones tomadas

El usuario pidió avanzar sin esperar más respuestas. Estas decisiones salen de lo observado en octubre, y se ajustan cuando ella revise resultados reales.

| Tema | Decisión | Por qué |
|---|---|---|
| Archivos originales | Nunca se modifican; siempre se genera una copia | Un error no debe dañar su trabajo contable |
| Ejecución | Local, sin servicios externos ni APIs de IA | Son datos financieros de una empresa y de terceros |
| Sentido del cruce | `DB` con `Débito` y `CR` con `Crédito` | Es lo que ocurre en octubre |
| Comisiones, IVA, impuestos e intereses | Quedan sin comprobante | Casi todas las filas sin conciliar son de estos conceptos |
| Coincidencia | Mismo lado siempre, misma fecha y monto con hasta $1.000 de diferencia; lo demás va a pendientes | Lo definió el usuario. Reemplaza la tolerancia de $1 y los 3 días del análisis inicial |
| Diferencia de monto | Se elige el de monto más cercano y queda seguro aunque tenga diferencia; una coincidencia exacta nunca pierde su movimiento | Lo definió el usuario: es el que más coincide en lado, fecha y monto. En octubre, 9 de 9 únicos con diferencia coincidían |
| Empate de monto y fecha | Si los empatados son del mismo tercero, segura. Si son de terceros distintos, decide el nombre aproximado (abreviado, cortado o con letras distintas) y la fila va en azul (`9BC2E6`). Si el nombre no distingue, naranja | Lo definió el usuario. En octubre el nombre nunca eligió mal |
| Filas azules y naranjas en la copia | Se escribe el comprobante propuesto y la fila `A:G` se pinta | Lo definió el usuario: ella solo confirma o corrige. El color cambia si ella prefiere otro |
| Asignación | Se ordenan todas las parejas por diferencia de monto y luego por parecido del nombre, y se asignan en ese orden | Así el resultado no depende del orden de las filas del extracto |
| Lista de revisión | Archivo de Excel aparte, junto a la copia, no una hoja nueva | La copia solo cambia la hoja principal y los estilos, y su plantilla no acumula hojas en cada ejecución |
| Filas ya conciliadas | No se sobrescriben y sus comprobantes salen de los candidatos | Ella acumula el mes en el mismo archivo; así no se propone dos veces el mismo |
| Hoja de movimientos | Solo la primera; las demás no entran al proceso. `Secuencia` y `Detalle` son opcionales | Lo definió el usuario. El archivo de otro banco no trae esas dos columnas, que solo se muestran en la revisión |
| Escritura de la copia | Edición directa: solo cambian la hoja principal y los estilos. Falta confirmarla en Excel | `openpyxl` quita los saldos calculados y la configuración de impresora, y reescribe la tabla dinámica |
| Definición de “seguro” | El de monto más cercano sin empate con otro tercero | Medido en octubre: 0 inventadas; las seguras distintas a ella son las parejas intercambiadas conocidas y la fila 488 |
| Evaluación | Además del mes completo, simular una exportación parcial con los primeros días de la primera hoja | En octubre casi todos tienen pareja; noviembre llega parcial. El script no debe inventar parejas |
| Nombre de la salida | Si el archivo ya termina en `_conciliado` (aunque esté repetido), no se agrega otra vez; la revisión usa el mismo nombre base | Ella concilia a diario el archivo del día anterior |
| Ubicación de columnas | Por encabezado, no por posición ni por nombre de hoja o archivo | Los nombres cambian de un mes a otro |

## Salida esperada

Una copia del extracto con:

- **Coincidencia segura:** comprobante en `DOC CONCILIADO` y fila `A:G` en amarillo (`FFFF00`), igual que hace ella.
- **Parecida:** empate de monto y fecha con otro tercero, resuelto por el nombre: comprobante propuesto y fila `A:G` en azul (`9BC2E6`), con sus candidatos en la lista de revisión.
- **Caso dudoso:** empate que el nombre no resuelve: comprobante propuesto y fila `A:G` en naranja (`FFC000`), con sus candidatos en la lista de revisión.
- **Lista de revisión:** archivo aparte con los dudosos y sus candidatos (comprobante, fila de movimientos, monto, diferencia y cuál se recomienda), y los pendientes de ambos archivos, con su fila de origen.
- **Copia marcada de movimientos:** `<movimientos>_conciliado.xlsx`. Cada movimiento usado se pinta del mismo color que su fila del extracto, en todas sus columnas con encabezado; no se escribe texto. Los sin pareja quedan sin color. Ella ya marca en amarillo los movimientos que concilia a mano. Falta que confirme si le sirve así.

## Criterio de éxito

- **Evaluación con octubre:** igualar o superar la línea base sin proponer comprobantes en filas que ella dejó vacías.
- **Dudosos:** bajar los 154 casos con varios candidatos, porque revisar tantos reduce el ahorro de tiempo.
- **Errores:** un comprobante incorrecto no debe quedar marcado como seguro. Un caso dudoso de más cuesta unos segundos, pero un error marcado como seguro puede pasar sin revisión.
- **Validación real:** ella confirma que el resultado de noviembre le sirve y le ahorra tiempo.

La evaluación usa un solo mes, así que su resultado puede ser optimista. Cuando haya otro mes conciliado, se repite con ese mes.

## Por validar con ella

No bloquean el desarrollo:

- Si la empresa permite procesar estos datos fuera de sus herramientas habituales.
- Si las comisiones e impuestos siempre quedan sin comprobante, o los registra de otra forma.
- Si en los pares de filas intercambiadas da igual cuál comprobante use cada fila.
- Fila 488 de octubre: usó un comprobante del lado contrario; el script propone otro que coincide en `DB`, monto y fecha.
- Cómo distingue pagos con el mismo monto y fecha, o si da igual cuál comprobante use.
- Si puede instalar Python en su equipo o necesita un ejecutable. Tampoco se conoce su sistema operativo.

## Comandos

En Linux, desde la raíz del proyecto. En Windows, usa `.venv\Scripts\python` en lugar de `.venv/bin/python`.

```bash
# Instalar
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt

# Probar
.venv/bin/python -m pytest tests -q

# Verificar que una copia no dañe el extracto (solo imprime conteos)
# Agrega --libreoffice para abrir la copia también con LibreOffice
.venv/bin/python scripts/verificar_copia.py RUTA_EXTRACTO CARPETA_SALIDA --metodo directo
```

Para ver lo que se lee de los archivos (solo imprime conteos):

```bash
.venv/bin/python -m conciliacion.lectura RUTA_EXTRACTO RUTA_MOVIMIENTOS
```

Para medir el cruce contra un mes ya conciliado, incluidas exportaciones parciales (solo imprime conteos y números de fila):

```bash
.venv/bin/python -m conciliacion.evaluacion RUTA_EXTRACTO RUTA_MOVIMIENTOS --detalle
```

Para conciliar (comando principal):

```bash
.venv/bin/python -m conciliacion RUTA_EXTRACTO RUTA_MOVIMIENTOS CARPETA_SALIDA
```

Crea en la carpeta de salida `<extracto>_conciliado.xlsx` (amarillo: seguro; azul: elegido por el nombre; naranja: dudoso), `<movimientos>_conciliado.xlsx` (mismos colores) y `<extracto>_revision.xlsx` (azules y naranjas con sus candidatos, y pendientes). No sobrescribe archivos existentes y comprueba que los originales no cambien.

Para conciliar con la ventana (la forma pensada para ella):

```bash
.venv/bin/python -m conciliacion.interfaz
```

Guarda los resultados en `Conciliaciones/<fecha y hora>/`, junto al extracto elegido, para que cada ejecución quede aparte. En Linux requiere `sudo apt install python3-tk`; en Windows y macOS Tkinter viene con Python. El ejecutable se genera con PyInstaller en el mismo sistema operativo de ella, que aún no se conoce.

Guarda las copias con datos reales fuera del repositorio. Durante el desarrollo se usa `~/Escritorio/conciliacion-salida/`.

## Cómo trabajar

- Sé directo y explica el porqué de las decisiones técnicas importantes.
- En este proyecto, avanza con supuestos razonables y déjalos documentados. Pregunta solo si una decisión bloquea el trabajo o puede dañar datos.
- El usuario está aprendiendo: explica el enfoque de las pruebas al escribirlas.
- Entrega código completo, con nombres descriptivos y siguiendo PEP 8.
- Escribe pruebas y documentación de lo que se construye.
- No agregues ni cambies nada que no se haya pedido. Si crees que algo debería cambiar, pregunta primero.
- Pregunta antes de borrar o sobrescribir archivos.
- El código debe funcionar en Linux y en Windows. Usa `pathlib` para las rutas.

## Git

- Repositorio inicializado en la rama `main`, sin commits ni remoto.
- Antes del primer commit, pregunta qué identidad usar y configúrala solo en este repositorio. No cambies la configuración global de Git.
- Haz commit o push solo cuando el usuario lo pida.
- `.gitignore` excluye los archivos de Excel y CSV. Antes de cada commit, confirma que no se incluya ningún dato real.
