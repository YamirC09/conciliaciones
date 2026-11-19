# Cómo usar el programa de conciliación

## Instalarlo (solo la primera vez)

1. Entra a la página de versiones (*Releases*) del repositorio y abre la más reciente.
2. Descarga el archivo de tu equipo:
   - **Windows:** `Conciliacion-Windows.zip`
   - **Mac:** `Conciliacion-Mac.zip`
3. Descomprime el archivo y guarda el programa donde quieras, por ejemplo en el Escritorio.
   - En Windows queda un archivo `Conciliacion.exe`.
   - En Mac queda una aplicación `Conciliacion`.

No hay que instalar nada más.

### La primera vez que lo abres

El programa no está firmado por una empresa, así que el sistema muestra una advertencia una sola vez:

- **Windows:** si aparece "Windows protegió su PC", haz clic en **Más información** y luego en **Ejecutar de todas formas**.
- **Mac:** haz clic derecho sobre la aplicación, elige **Abrir** y confirma con **Abrir**. Si no te deja, ve a **Ajustes del Sistema → Privacidad y seguridad** y pulsa **Abrir igualmente**.

Las siguientes veces abre con doble clic, sin advertencias.

## Conciliar

1. Abre el programa.
2. En **Extracto del banco**, pulsa **Elegir…** y busca tu archivo del extracto.
3. En **Movimientos del sistema**, pulsa **Elegir…** y busca la exportación de movimientos.
4. Pulsa **Conciliar** y espera unos segundos.
5. Pulsa **Abrir carpeta de resultados**.

Tus archivos originales nunca se modifican.

## Dónde quedan los resultados

En una carpeta `Conciliaciones`, junto al extracto que elegiste, con una subcarpeta por cada vez que concilias (con la fecha y la hora). Allí hay tres archivos:

| Archivo | Qué es |
|---|---|
| Extracto `..._conciliado.xlsx` | Tu extracto con los comprobantes escritos y las filas pintadas |
| Movimientos `..._conciliado.xlsx` | Tus movimientos, pintados del mismo color que su fila del extracto |
| `..._revision.xlsx` | Lo que debes revisar, con todos los candidatos, y lo que quedó pendiente |

Si concilias un archivo que ya decía `_conciliado`, el nombre no se repite.

## Qué significa cada color

| Color | Significa | Qué hacer |
|---|---|---|
| Amarillo | Coincide en lado, fecha y monto | Nada, igual que cuando lo marcas tú |
| Azul | El mismo comprobante quedó en otra fila con otra persona | Revisar que esté bien repartido |
| Naranja | Había varios movimientos iguales y el nombre no permitió elegir | Confirmar o cambiar el comprobante escrito |
| Sin color | No se encontró pareja | Conciliar a mano |

## Si algo sale mal

- **"Si está abierto en Excel, ciérralo":** cierra el archivo en Excel y vuelve a pulsar **Conciliar**.
- **"Revisa que elegiste cada archivo en su lugar":** probablemente invertiste el extracto y los movimientos.
- **Cualquier otro mensaje:** anota el mensaje y el número de versión que aparece en el título de la ventana, y repórtalo.

## Actualizar

Cuando haya una versión nueva, descárgala igual que la primera vez y reemplaza el programa anterior. El número de versión aparece en el título de la ventana.
