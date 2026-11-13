"""Escribe la copia marcada del extracto sin alterar el resto del archivo.

Un `.xlsx` es un `.zip` con varios archivos XML. Aquí se copian tal cual todas
las partes, salvo dos: la hoja principal, donde se escriben los comprobantes, y
los estilos, donde se agregan los colores. Así la tabla dinámica, las demás
hojas y los saldos ya calculados quedan como estaban.

`openpyxl` no sirve para esto: al guardar, quita los saldos calculados y la
configuración de impresora, y reescribe la tabla dinámica (ver la etapa 1 en
`docs/analisis-inicial.md`).
"""

import re
import zipfile
from dataclasses import dataclass
from pathlib import PurePosixPath
from xml.sax.saxutils import escape

from openpyxl.utils.cell import column_index_from_string, get_column_letter

COLOR_SEGURA = "FFFF00"
COLOR_DUDOSA = "FFC000"
COLOR_REPETIDA = "9BC2E6"  # Azul claro: el texto negro se sigue leyendo.


@dataclass(frozen=True)
class Marca:
    """Texto que se escribe en `DOC CONCILIADO` y color de la fila."""

    texto: str
    color: str


def resolver_destino(base, destino):
    """Convierte el destino de una relación en la ruta dentro del `.xlsx`."""
    if destino.startswith("/"):
        return destino.lstrip("/")
    partes = []
    for segmento in (PurePosixPath(base) / destino).parts:
        if segmento == "..":
            partes.pop()
        else:
            partes.append(segmento)
    return "/".join(partes)


def atributo(etiqueta, nombre):
    """Devuelve el valor de un atributo dentro de una etiqueta XML, o `None`."""
    encontrado = re.search(rf'\s{nombre}="([^"]*)"', etiqueta)
    return encontrado.group(1) if encontrado else None


def ruta_de_hoja(archivo_zip, indice):
    """Devuelve la ruta interna de la hoja en esa posición (desde 0)."""
    relaciones_xml = archivo_zip.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    relaciones = {
        atributo(etiqueta, "Id"): atributo(etiqueta, "Target")
        for etiqueta in re.findall(r"<Relationship\b[^>]*>", relaciones_xml)
    }
    libro_xml = archivo_zip.read("xl/workbook.xml").decode("utf-8")
    hojas = re.findall(r"<sheet\b[^>]*>", libro_xml)
    return resolver_destino("xl", relaciones[atributo(hojas[indice], "r:id")])


def agregar_relleno(estilos_xml, color):
    """Devuelve los estilos con un relleno sólido del color y su índice.

    Reutiliza el relleno si el archivo ya lo tiene, y solo lo agrega si falta.
    """
    seccion = re.search(r"<fills\b[^>]*>(.*?)</fills>", estilos_xml, re.DOTALL)
    rellenos = re.findall(r"<fill>.*?</fill>", seccion.group(1), re.DOTALL)
    for indice, relleno in enumerate(rellenos):
        if 'patternType="solid"' in relleno and re.search(
            rf'<fgColor rgb="(?:FF)?{color}"', relleno
        ):
            return estilos_xml, indice
    nuevo = (
        '<fill><patternFill patternType="solid">'
        f'<fgColor rgb="FF{color}"/><bgColor indexed="64"/>'
        "</patternFill></fill>"
    )
    estilos_xml = estilos_xml.replace("</fills>", f"{nuevo}</fills>", 1)
    estilos_xml = re.sub(
        r'(<fills\b[^>]*?count=")\d+(")',
        lambda m: f"{m.group(1)}{len(rellenos) + 1}{m.group(2)}",
        estilos_xml,
        count=1,
    )
    return estilos_xml, len(rellenos)


def agregar_estilo_con_relleno(estilos_xml, indice_base, indice_relleno):
    """Agrega una copia del estilo indicado con otro relleno y devuelve su índice.

    Solo cambia el relleno: fuente, bordes, alineación y formato de número se
    conservan, para que la celda se vea igual, pero con el color nuevo.
    """
    seccion = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", estilos_xml, re.DOTALL)
    estilos = re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", seccion.group(1), re.DOTALL)
    apertura = re.match(r"<xf\b[^>]*", estilos[indice_base]).group(0).rstrip("/")
    resto = estilos[indice_base][len(apertura):]
    apertura = re.sub(r'\s(?:fillId|applyFill)="[^"]*"', "", apertura)
    nuevo = f'{apertura} fillId="{indice_relleno}" applyFill="1"{resto}'
    estilos_xml = estilos_xml.replace("</cellXfs>", f"{nuevo}</cellXfs>", 1)
    estilos_xml = re.sub(
        r'(<cellXfs\b[^>]*?count=")\d+(")',
        lambda m: f"{m.group(1)}{len(estilos) + 1}{m.group(2)}",
        estilos_xml,
        count=1,
    )
    return estilos_xml, len(estilos)


def estilos_de_columnas(hoja_xml):
    """Devuelve el estilo de cada columna que tenga uno definido."""
    estilos = {}
    for columna in re.findall(r"<col\b[^>]*>", hoja_xml):
        minimo = atributo(columna, "min")
        maximo = atributo(columna, "max")
        estilo = atributo(columna, "style")
        if minimo and maximo and estilo:
            for numero in range(int(minimo), min(int(maximo), 256) + 1):
                estilos[numero] = int(estilo)
    return estilos


class Escritor:
    """Acumula los cambios sobre el XML de la hoja y de los estilos."""

    def __init__(self, hoja_xml, estilos_xml):
        self.hoja_xml = hoja_xml
        self.estilos_xml = estilos_xml
        self.rellenos = {}
        self.estilos_con_relleno = {}
        self.estilos_por_columna = estilos_de_columnas(hoja_xml)

    def estilo_con_color(self, base, color):
        """Devuelve el índice de un estilo igual a `base`, pero con ese color."""
        if color not in self.rellenos:
            self.estilos_xml, self.rellenos[color] = agregar_relleno(
                self.estilos_xml, color
            )
        clave = (base, self.rellenos[color])
        if clave not in self.estilos_con_relleno:
            self.estilos_xml, self.estilos_con_relleno[clave] = (
                agregar_estilo_con_relleno(self.estilos_xml, base, clave[1])
            )
        return self.estilos_con_relleno[clave]

    def estilo_base(self, celda, etiqueta_fila, columna):
        """Devuelve el estilo que tiene, o heredaría, una celda."""
        if celda is not None:
            return int(atributo(re.match(r"<c\b[^>]*", celda).group(0), "s") or 0)
        estilo_fila = atributo(etiqueta_fila, "s")
        if atributo(etiqueta_fila, "customFormat") == "1" and estilo_fila:
            return int(estilo_fila)
        return self.estilos_por_columna.get(columna, 0)

    def marcar_fila(self, fila, marca, columna_documento, columnas_color):
        """Escribe el texto y pinta las celdas de una fila.

        No sobrescribe una celda de documento que ya tenga contenido.
        """
        apertura = re.search(rf'<row\b[^>]*?\sr="{fila}"[^>]*?>', self.hoja_xml)
        if apertura is None or apertura.group(0).endswith("/>"):
            raise ValueError("La fila a marcar no existe o no tiene celdas.")
        fin = self.hoja_xml.index("</row>", apertura.end())
        contenido = self.hoja_xml[apertura.end():fin]

        celdas = {}
        for celda in re.findall(r"<c\b[^>]*?(?:/>|>.*?</c>)", contenido, re.DOTALL):
            referencia = re.search(r'\br="([A-Z]+)\d+"', celda).group(1)
            celdas[column_index_from_string(referencia)] = celda

        etiqueta_fila = apertura.group(0)
        for columna in columnas_color:
            celda = celdas.get(columna)
            nuevo_estilo = self.estilo_con_color(
                self.estilo_base(celda, etiqueta_fila, columna), marca.color
            )
            referencia = f"{get_column_letter(columna)}{fila}"

            if columna == columna_documento:
                if celda is not None and re.search(r"<(?:v|is|f)\b", celda):
                    raise ValueError("La celda de documento ya tiene contenido.")
                celdas[columna] = (
                    f'<c r="{referencia}" s="{nuevo_estilo}" t="inlineStr">'
                    f"<is><t>{escape(marca.texto)}</t></is></c>"
                )
            elif celda is None:
                celdas[columna] = f'<c r="{referencia}" s="{nuevo_estilo}"/>'
            elif re.search(r'<c\b[^>]*?\ss="\d+"', celda):
                celdas[columna] = re.sub(
                    r'(<c\b[^>]*?\ss=")\d+(")',
                    lambda m: f"{m.group(1)}{nuevo_estilo}{m.group(2)}",
                    celda,
                    count=1,
                )
            else:
                celdas[columna] = celda.replace("<c ", f'<c s="{nuevo_estilo}" ', 1)

        contenido_nuevo = "".join(celdas[numero] for numero in sorted(celdas))
        etiqueta_nueva = etiqueta_fila
        rango = re.search(r'\bspans="(\d+):(\d+)"', etiqueta_fila)
        if rango:
            etiqueta_nueva = etiqueta_fila.replace(
                rango.group(0),
                f'spans="{min(int(rango.group(1)), min(celdas))}:'
                f'{max(int(rango.group(2)), max(celdas))}"',
            )
        self.hoja_xml = (
            self.hoja_xml[:apertura.start()]
            + etiqueta_nueva
            + contenido_nuevo
            + self.hoja_xml[fin:]
        )


def escribir_copia(origen, destino, indice_hoja, columna_documento, columnas_color, marcas):
    """Crea la copia del extracto con las marcas indicadas.

    `marcas` relaciona el número de fila con su `Marca`. Las columnas se
    cuentan desde 1 y la hoja desde 0. Con `columna_documento=None` solo se
    pinta, sin escribir texto (así se marca la copia de movimientos). Si una
    marca falla, no se crea el archivo, porque el destino se abre solo después
    de aplicarlas todas.
    """
    with zipfile.ZipFile(origen) as entrada:
        parte_hoja = ruta_de_hoja(entrada, indice_hoja)
        escritor = Escritor(
            entrada.read(parte_hoja).decode("utf-8"),
            entrada.read("xl/styles.xml").decode("utf-8"),
        )
        for fila in sorted(marcas):
            escritor.marcar_fila(fila, marcas[fila], columna_documento, columnas_color)

        reemplazos = {
            parte_hoja: escritor.hoja_xml.encode("utf-8"),
            "xl/styles.xml": escritor.estilos_xml.encode("utf-8"),
        }
        with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as salida:
            for informacion in entrada.infolist():
                contenido = reemplazos.get(
                    informacion.filename, entrada.read(informacion.filename)
                )
                salida.writestr(informacion, contenido, zipfile.ZIP_DEFLATED)
