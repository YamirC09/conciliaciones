"""Comprueba si una copia marcada del extracto conserva intacto su contenido.

Genera una copia de prueba del extracto, con unas pocas marcas en filas vacías
de la plantilla, y la compara con el original. La comparación es por
significado: si dos archivos guardan lo mismo escrito de otra forma, no cuenta
como diferencia.

Solo imprime conteos y nombres técnicos de elementos de Excel. Nunca imprime
valores de celdas, nombres de hojas ni rutas, porque pueden contener datos
reales.

Uso:
    python scripts/verificar_copia.py EXTRACTO CARPETA_SALIDA --metodo directo
    Agrega --libreoffice para recalcular la copia y comparar los saldos.
"""

import argparse
import hashlib
import re
import shutil
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from xml.sax.saxutils import escape

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import PatternFill
from openpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
)

COLOR_CONCILIADO = "FFFF00"
TEXTO_MARCA = "PRUEBA"
CANTIDAD_MARCAS = 3
ENCABEZADO_DOCUMENTO = "DOC CONCILIADO"
ENCABEZADO_FECHA = "FECHA"
ENCABEZADO_SALDOS = "SALDOS"

ATRIBUTOS_FILA = ("ht", "customHeight", "hidden", "outlineLevel", "collapsed")
ELEMENTOS_ESTILO_CELDA = {"numFmts", "fonts", "fills", "borders", "cellXfs"}
TOLERANCIA_NUMERICA = 1e-6


def nombre_local(etiqueta):
    """Quita el espacio de nombres de una etiqueta XML."""
    return etiqueta.rsplit("}", 1)[-1]


def canonizar(elemento, ordenar_hijos=False):
    """Convierte un elemento XML en una estructura comparable.

    Ignora los identificadores de relación (`r:id`) y `xml:space`, porque un
    programa puede reescribirlos sin que cambie el contenido.
    """
    atributos = tuple(
        sorted(
            (nombre_local(clave), valor)
            for clave, valor in elemento.attrib.items()
            if nombre_local(clave) not in {"id", "space"}
        )
    )
    hijos = [canonizar(hijo, ordenar_hijos) for hijo in elemento]
    if ordenar_hijos:
        hijos.sort()
    return (
        nombre_local(elemento.tag),
        atributos,
        (elemento.text or "").strip(),
        tuple(hijos),
    )


def leer_xml(archivo_zip, parte):
    """Devuelve la raíz de una parte XML del archivo."""
    with archivo_zip.open(parte) as contenido:
        return ET.parse(contenido).getroot()


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


def rutas_de_hojas(archivo_zip):
    """Devuelve las rutas internas de las hojas, en el orden del libro."""
    relaciones = {
        relacion.attrib["Id"]: relacion.attrib["Target"]
        for relacion in leer_xml(archivo_zip, "xl/_rels/workbook.xml.rels")
    }
    rutas = []
    for hijo in leer_xml(archivo_zip, "xl/workbook.xml"):
        if nombre_local(hijo.tag) != "sheets":
            continue
        for hoja in hijo:
            identificador = next(
                valor
                for clave, valor in hoja.attrib.items()
                if nombre_local(clave) == "id"
            )
            rutas.append(resolver_destino("xl", relaciones[identificador]))
    return rutas


def leer_textos_compartidos(archivo_zip):
    """Devuelve cada texto compartido en forma comparable."""
    if "xl/sharedStrings.xml" not in archivo_zip.namelist():
        return []
    return [
        canonizar(texto)[3] for texto in leer_xml(archivo_zip, "xl/sharedStrings.xml")
    ]


def leer_estilos(archivo_zip):
    """Traduce cada índice de estilo de celda a sus componentes reales.

    Los programas renumeran los estilos al guardar. Por eso se compara lo que
    significa cada índice, no el número.
    """
    raiz = leer_xml(archivo_zip, "xl/styles.xml")
    secciones = {nombre_local(hijo.tag): hijo for hijo in raiz}
    formatos = {
        formato.attrib["numFmtId"]: formato.attrib.get("formatCode", "")
        for formato in secciones.get("numFmts", [])
    }

    def componentes(seccion):
        return [canonizar(item, ordenar_hijos=True) for item in secciones[seccion]]

    # Un relleno sin `patternType` equivale a "none" según el formato de Excel.
    for relleno in secciones["fills"]:
        for patron in relleno:
            if nombre_local(patron.tag) == "patternFill":
                patron.attrib.setdefault("patternType", "none")

    fuentes = componentes("fonts")
    rellenos = componentes("fills")
    bordes = componentes("borders")
    estilos = []
    for formato_celda in secciones["cellXfs"]:
        extras = {
            nombre_local(hijo.tag): canonizar(hijo) for hijo in formato_celda
        }
        id_formato = formato_celda.attrib.get("numFmtId", "0")
        estilos.append(
            {
                "numFmt": formatos.get(id_formato, id_formato),
                "font": fuentes[int(formato_celda.attrib.get("fontId", 0))],
                "fill": rellenos[int(formato_celda.attrib.get("fillId", 0))],
                "border": bordes[int(formato_celda.attrib.get("borderId", 0))],
                "alignment": extras.get("alignment"),
                "protection": extras.get("protection"),
            }
        )
    return estilos


@dataclass
class Celda:
    """Contenido de una celda con sus referencias ya resueltas."""

    columna: int
    estilo: int = 0
    formula: str | None = None
    valor: object = None


@dataclass
class Fila:
    """Una fila de la hoja, con sus atributos y celdas."""

    numero: int
    atributos: dict
    estilo: int | None
    celdas: dict = field(default_factory=dict)


def normalizar_valor(tipo, valor, celda_xml, textos):
    """Convierte el valor de una celda en una forma comparable."""
    if tipo == "s":
        return ("texto", textos[int(valor)]) if valor is not None else None
    if tipo == "inlineStr":
        for hijo in celda_xml:
            if nombre_local(hijo.tag) == "is":
                return ("texto", canonizar(hijo)[3])
        return None
    if valor is None:
        return None
    if tipo in {"str", "e", "b"}:
        return (tipo, valor)
    try:
        return ("numero", float(valor))
    except ValueError:
        return ("otro", valor)


def leer_celda(celda_xml, textos, formulas_compartidas):
    """Interpreta un elemento `<c>` y expande sus fórmulas compartidas."""
    referencia = celda_xml.attrib["r"]
    tipo = celda_xml.attrib.get("t")
    formula = None
    valor = None
    for hijo in celda_xml:
        etiqueta = nombre_local(hijo.tag)
        if etiqueta == "v":
            valor = hijo.text
        elif etiqueta == "f":
            formula = hijo.text or ""
            if hijo.attrib.get("t") == "shared":
                indice = hijo.attrib["si"]
                if formula:
                    formulas_compartidas[indice] = (formula, referencia)
                else:
                    base, origen = formulas_compartidas[indice]
                    formula = Translator(
                        f"={base}", origin=origen
                    ).translate_formula(referencia)[1:]
    columna, _ = coordinate_from_string(referencia)
    return Celda(
        columna=column_index_from_string(columna),
        estilo=int(celda_xml.attrib.get("s", 0)),
        formula=formula,
        valor=normalizar_valor(tipo, valor, celda_xml, textos),
    )


def recorrer_hoja(archivo_zip, parte, textos, elementos_superiores):
    """Entrega las filas de una hoja una por una, sin cargarla toda.

    La hoja grande del extracto tiene millones de celdas. Leerla por partes
    mantiene baja la memoria. Los demás elementos de la hoja, como el filtro o
    los paneles, se guardan en `elementos_superiores`.
    """
    formulas_compartidas = {}
    profundidad = 0
    datos_hoja = None
    with archivo_zip.open(parte) as contenido:
        for evento, elemento in ET.iterparse(contenido, events=("start", "end")):
            if evento == "start":
                profundidad += 1
                if nombre_local(elemento.tag) == "sheetData":
                    datos_hoja = elemento
                continue
            profundidad -= 1
            etiqueta = nombre_local(elemento.tag)
            if profundidad == 1 and etiqueta != "sheetData":
                elementos_superiores[etiqueta] = canonizar(elemento)
            elif profundidad == 2 and etiqueta == "row":
                celdas = {}
                for celda_xml in elemento:
                    celda = leer_celda(celda_xml, textos, formulas_compartidas)
                    celdas[celda.columna] = celda
                con_estilo = elemento.attrib.get("customFormat") == "1"
                yield Fila(
                    numero=int(elemento.attrib["r"]),
                    atributos={
                        clave: elemento.attrib[clave]
                        for clave in ATRIBUTOS_FILA
                        if clave in elemento.attrib
                    },
                    estilo=int(elemento.attrib.get("s", 0)) if con_estilo else None,
                    celdas=celdas,
                )
                datos_hoja.clear()


def emparejar_filas(filas_a, filas_b):
    """Recorre dos hojas a la vez y empareja sus filas por número."""
    fila_a = next(filas_a, None)
    fila_b = next(filas_b, None)
    while fila_a is not None or fila_b is not None:
        if fila_b is None or (fila_a is not None and fila_a.numero < fila_b.numero):
            yield fila_a, None
            fila_a = next(filas_a, None)
        elif fila_a is None or fila_b.numero < fila_a.numero:
            yield None, fila_b
            fila_b = next(filas_b, None)
        else:
            yield fila_a, fila_b
            fila_a = next(filas_a, None)
            fila_b = next(filas_b, None)


def valores_iguales(valor_a, valor_b):
    """Compara dos valores, con un margen mínimo para los números."""
    if (
        valor_a is not None
        and valor_b is not None
        and valor_a[0] == "numero"
        and valor_b[0] == "numero"
    ):
        return abs(valor_a[1] - valor_b[1]) <= TOLERANCIA_NUMERICA
    return valor_a == valor_b


@dataclass
class ResultadoHoja:
    """Conteos de una hoja comparada contra su copia."""

    filas: int = 0
    celdas: int = 0
    formulas: int = 0
    calculados_original: int = 0
    calculados_copia: int = 0
    marcas_correctas: int = 0
    rellenos_correctos: int = 0
    diferencias: Counter = field(default_factory=Counter)
    elementos_faltantes: list = field(default_factory=list)
    elementos_nuevos: list = field(default_factory=list)
    elementos_cambiados: list = field(default_factory=list)


def comparar_celda(celda_a, celda_b, contexto, resultado):
    """Registra las diferencias de una celda entre el original y la copia."""
    estilos_a, estilos_b, marcas, columna_doc, columnas_marca = contexto
    columna = celda_a.columna if celda_a else celda_b.columna
    existia_en_original = celda_a is not None
    celda_a = celda_a or Celda(columna=columna)
    celda_b = celda_b or Celda(columna=columna)

    resultado.celdas += 1
    if celda_a.formula is not None:
        resultado.formulas += 1
        resultado.calculados_original += celda_a.valor is not None
        resultado.calculados_copia += celda_b.valor is not None

    if celda_a.formula != celda_b.formula:
        resultado.diferencias["formula"] += 1

    es_marca_doc = marcas and columna == columna_doc
    if es_marca_doc:
        if celda_b.valor == ("texto", (("t", (), TEXTO_MARCA, ()),)):
            resultado.marcas_correctas += 1
        else:
            resultado.diferencias["marca_de_texto_incorrecta"] += 1
    elif celda_a.formula is not None:
        if celda_a.valor is not None and celda_b.valor is None:
            resultado.diferencias["valor_calculado_perdido"] += 1
        elif celda_b.valor is not None and not valores_iguales(
            celda_a.valor, celda_b.valor
        ):
            resultado.diferencias["valor_calculado_distinto"] += 1
    elif not valores_iguales(celda_a.valor, celda_b.valor):
        resultado.diferencias["valor"] += 1

    estilo_a = estilos_a[celda_a.estilo]
    estilo_b = estilos_b[celda_b.estilo]
    es_celda_marcada = marcas and columna in columnas_marca
    if es_celda_marcada and COLOR_CONCILIADO in str(estilo_b["fill"]):
        resultado.rellenos_correctos += 1
    if es_celda_marcada and not existia_en_original:
        return  # La celda es nueva: no hay estilo original con cuál comparar.
    for componente, valor_a in estilo_a.items():
        if valor_a == estilo_b[componente]:
            continue
        if es_celda_marcada and componente == "fill":
            continue
        resultado.diferencias[f"estilo_{componente}"] += 1


def comparar_filas(fila_a, fila_b, contexto, resultado):
    """Compara los atributos y todas las celdas de una fila."""
    estilos_a, estilos_b, marcas, columna_doc, columnas_marca = contexto
    resultado.filas += 1
    atributos_a = fila_a.atributos if fila_a else {}
    atributos_b = fila_b.atributos if fila_b else {}
    for clave in ATRIBUTOS_FILA:
        if atributos_a.get(clave) != atributos_b.get(clave):
            resultado.diferencias[f"fila_{clave}"] += 1

    estilo_fila_a = fila_a.estilo if fila_a else None
    estilo_fila_b = fila_b.estilo if fila_b else None
    if (estilo_fila_a is None) != (estilo_fila_b is None) or (
        estilo_fila_a is not None
        and estilos_a[estilo_fila_a] != estilos_b[estilo_fila_b]
    ):
        resultado.diferencias["fila_estilo"] += 1

    celdas_a = fila_a.celdas if fila_a else {}
    celdas_b = fila_b.celdas if fila_b else {}
    numero = fila_a.numero if fila_a else fila_b.numero
    contexto_fila = (
        estilos_a,
        estilos_b,
        numero in marcas,
        columna_doc,
        columnas_marca,
    )
    for columna in sorted(celdas_a.keys() | celdas_b.keys()):
        comparar_celda(
            celdas_a.get(columna), celdas_b.get(columna), contexto_fila, resultado
        )


def comparar_hoja(zip_a, zip_b, parte_a, parte_b, datos, marcas_hoja):
    """Compara una hoja completa del original contra la de la copia."""
    textos_a, textos_b, estilos_a, estilos_b = datos
    superiores_a = {}
    superiores_b = {}
    resultado = ResultadoHoja()
    filas_marca, columna_doc, columnas_marca = marcas_hoja
    contexto = (estilos_a, estilos_b, filas_marca, columna_doc, columnas_marca)
    for fila_a, fila_b in emparejar_filas(
        recorrer_hoja(zip_a, parte_a, textos_a, superiores_a),
        recorrer_hoja(zip_b, parte_b, textos_b, superiores_b),
    ):
        comparar_filas(fila_a, fila_b, contexto, resultado)

    resultado.elementos_faltantes = sorted(superiores_a.keys() - superiores_b.keys())
    resultado.elementos_nuevos = sorted(superiores_b.keys() - superiores_a.keys())
    resultado.elementos_cambiados = sorted(
        etiqueta
        for etiqueta in superiores_a.keys() & superiores_b.keys()
        if superiores_a[etiqueta] != superiores_b[etiqueta]
    )
    return resultado


def comparar_elementos_superiores(zip_a, zip_b, parte_a, parte_b, ignorar=frozenset()):
    """Compara los elementos de primer nivel de una parte XML."""
    hijos_a = Counter()
    hijos_b = Counter()
    for raiz, hijos in ((leer_xml(zip_a, parte_a), hijos_a), (leer_xml(zip_b, parte_b), hijos_b)):
        for hijo in raiz:
            if nombre_local(hijo.tag) not in ignorar:
                hijos[(nombre_local(hijo.tag), canonizar(hijo))] += 1
    faltan = sorted({clave[0] for clave in (hijos_a - hijos_b)})
    sobran = sorted({clave[0] for clave in (hijos_b - hijos_a)})
    return faltan, sobran


def resumen_partes(zip_a, zip_b):
    """Clasifica las partes internas según se conserven o no."""
    nombres_a = set(zip_a.namelist())
    nombres_b = set(zip_b.namelist())
    comunes = nombres_a & nombres_b
    identicas = {
        parte
        for parte in comunes
        if hashlib.sha256(zip_a.read(parte)).digest()
        == hashlib.sha256(zip_b.read(parte)).digest()
    }
    return {
        "identicas": sorted(identicas),
        "cambiadas": sorted(comunes - identicas),
        "faltantes": sorted(nombres_a - nombres_b),
        "nuevas": sorted(nombres_b - nombres_a),
    }


def encontrar_hoja_principal(archivo_zip):
    """Ubica la hoja principal por sus encabezados y las filas a marcar.

    Devuelve el índice de la hoja, las columnas de la marca y las primeras filas
    de plantilla que siguen a la última con fecha.
    """
    textos = leer_textos_compartidos(archivo_zip)
    for indice, parte in enumerate(rutas_de_hojas(archivo_zip)):
        generador = recorrer_hoja(archivo_zip, parte, textos, {})
        primera = next(generador, None)
        if primera is None or primera.numero != 1:
            generador.close()
            continue
        encabezados = {}
        for columna, celda in primera.celdas.items():
            if celda.valor and celda.valor[0] == "texto":
                encabezados[columna] = "".join(
                    nodo[2] for nodo in celda.valor[1]
                ).strip()
        if not {ENCABEZADO_FECHA, ENCABEZADO_DOCUMENTO} <= set(encabezados.values()):
            generador.close()
            continue
        por_nombre = {nombre: columna for columna, nombre in encabezados.items()}
        columna_fecha = por_nombre[ENCABEZADO_FECHA]
        ultima_con_fecha = 1
        filas_vistas = set()
        for fila in generador:
            filas_vistas.add(fila.numero)
            celda_fecha = fila.celdas.get(columna_fecha)
            if celda_fecha is not None and celda_fecha.valor is not None:
                ultima_con_fecha = fila.numero
        filas_marca = [
            numero
            for numero in sorted(filas_vistas)
            if numero > ultima_con_fecha
        ][:CANTIDAD_MARCAS]
        if len(filas_marca) < CANTIDAD_MARCAS:
            raise ValueError("La plantilla no tiene filas vacías para las marcas.")
        columnas_marca = range(columna_fecha, por_nombre[ENCABEZADO_SALDOS] + 1)
        return indice, por_nombre[ENCABEZADO_DOCUMENTO], columnas_marca, filas_marca
    raise ValueError("Ninguna hoja tiene los encabezados del extracto.")


def generar_copia_openpyxl(origen, destino, hoja_principal):
    """Abre el extracto con `openpyxl`, agrega las marcas y lo guarda."""
    indice, columna_doc, columnas_marca, filas_marca = hoja_principal
    libro = load_workbook(origen)
    hoja = libro.worksheets[indice]
    relleno = PatternFill("solid", start_color=COLOR_CONCILIADO, end_color=COLOR_CONCILIADO)
    for fila in filas_marca:
        hoja.cell(fila, columna_doc).value = TEXTO_MARCA
        for columna in columnas_marca:
            hoja.cell(fila, columna).fill = relleno
    libro.save(destino)
    libro.close()


def agregar_relleno_amarillo(estilos_xml):
    """Devuelve los estilos con el relleno amarillo y el índice de ese relleno.

    Reutiliza el amarillo si el archivo ya lo tiene, y solo lo agrega si falta.
    """
    seccion = re.search(r"<fills\b[^>]*>(.*?)</fills>", estilos_xml, re.DOTALL)
    rellenos = re.findall(r"<fill>.*?</fill>", seccion.group(1), re.DOTALL)
    for indice, relleno in enumerate(rellenos):
        if 'patternType="solid"' in relleno and re.search(
            rf'<fgColor rgb="(?:FF)?{COLOR_CONCILIADO}"', relleno
        ):
            return estilos_xml, indice
    nuevo = (
        '<fill><patternFill patternType="solid">'
        f'<fgColor rgb="FF{COLOR_CONCILIADO}"/><bgColor indexed="64"/>'
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
    conservan, para que la celda se vea igual, pero en amarillo.
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
        minimo = re.search(r'\bmin="(\d+)"', columna)
        maximo = re.search(r'\bmax="(\d+)"', columna)
        estilo = re.search(r'\bstyle="(\d+)"', columna)
        if minimo and maximo and estilo:
            for numero in range(
                int(minimo.group(1)), min(int(maximo.group(1)), 256) + 1
            ):
                estilos[numero] = int(estilo.group(1))
    return estilos


def marcar_fila(hoja_xml, fila, marca, estilos_xml, cache_estilos):
    """Escribe el texto y pinta las celdas de una fila dentro del XML de la hoja.

    Devuelve la hoja y los estilos actualizados. No sobrescribe una celda de
    documento que ya tenga contenido.
    """
    columna_doc, columnas_marca, texto, indice_relleno = marca
    apertura = re.search(rf'<row\b[^>]*?\sr="{fila}"[^>]*?>', hoja_xml)
    if apertura is None or apertura.group(0).endswith("/>"):
        raise ValueError("La fila a marcar no existe o no tiene celdas.")
    fin = hoja_xml.index("</row>", apertura.end())
    contenido = hoja_xml[apertura.end():fin]

    celdas = {}
    for celda in re.findall(r"<c\b[^>]*?(?:/>|>.*?</c>)", contenido, re.DOTALL):
        referencia = re.search(r'\br="([A-Z]+)\d+"', celda).group(1)
        celdas[column_index_from_string(referencia)] = celda

    etiqueta_fila = apertura.group(0)
    estilo_fila = re.search(r'\bs="(\d+)"', etiqueta_fila)
    tiene_estilo_fila = 'customFormat="1"' in etiqueta_fila and estilo_fila
    por_columna = estilos_de_columnas(hoja_xml)

    for columna in columnas_marca:
        celda = celdas.get(columna)
        if celda is not None:
            estilo = re.search(r'<c\b[^>]*?\ss="(\d+)"', celda)
            base = int(estilo.group(1)) if estilo else 0
        elif tiene_estilo_fila:
            base = int(estilo_fila.group(1))
        else:
            base = por_columna.get(columna, 0)

        if base not in cache_estilos:
            estilos_xml, cache_estilos[base] = agregar_estilo_con_relleno(
                estilos_xml, base, indice_relleno
            )
        nuevo_estilo = cache_estilos[base]
        referencia = f"{get_column_letter(columna)}{fila}"

        if columna == columna_doc:
            if celda is not None and re.search(r"<(?:v|is|f)\b", celda):
                raise ValueError("La celda de documento ya tiene contenido.")
            celdas[columna] = (
                f'<c r="{referencia}" s="{nuevo_estilo}" t="inlineStr">'
                f"<is><t>{escape(texto)}</t></is></c>"
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
    hoja_xml = (
        hoja_xml[:apertura.start()]
        + etiqueta_nueva
        + contenido_nuevo
        + hoja_xml[fin:]
    )
    return hoja_xml, estilos_xml


def generar_copia_directa(origen, destino, hoja_principal):
    """Crea la copia cambiando solo la hoja principal y los estilos.

    Un `.xlsx` es un `.zip` con varios archivos XML. Aquí se copian tal cual
    todas las partes, salvo esas dos. Así la tabla dinámica, las demás hojas y
    los saldos ya calculados quedan como estaban.
    """
    indice, columna_doc, columnas_marca, filas_marca = hoja_principal
    with zipfile.ZipFile(origen) as entrada:
        parte_hoja = rutas_de_hojas(entrada)[indice]
        hoja_xml = entrada.read(parte_hoja).decode("utf-8")
        estilos_xml = entrada.read("xl/styles.xml").decode("utf-8")

        estilos_xml, indice_relleno = agregar_relleno_amarillo(estilos_xml)
        marca = (columna_doc, columnas_marca, TEXTO_MARCA, indice_relleno)
        cache_estilos = {}
        for fila in filas_marca:
            hoja_xml, estilos_xml = marcar_fila(
                hoja_xml, fila, marca, estilos_xml, cache_estilos
            )

        reemplazos = {
            parte_hoja: hoja_xml.encode("utf-8"),
            "xl/styles.xml": estilos_xml.encode("utf-8"),
        }
        with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as salida:
            for informacion in entrada.infolist():
                contenido = reemplazos.get(
                    informacion.filename, entrada.read(informacion.filename)
                )
                salida.writestr(informacion, contenido, zipfile.ZIP_DEFLATED)


def comprobar_con_libreoffice(copia, original, hoja_principal):
    """Abre la copia en LibreOffice, recalcula y compara los saldos.

    Es una revisión adicional: confirma que otro programa puede leer el archivo
    y que al recalcular las fórmulas dan lo mismo. No reemplaza abrirlo en Excel.
    """
    programa = shutil.which("soffice")
    if programa is None:
        print("LibreOffice no está instalado; se omite esta revisión.")
        return
    indice, columna_doc, _, filas_marca = hoja_principal
    with tempfile.TemporaryDirectory() as temporal:
        perfil = Path(temporal) / "perfil"
        (perfil / "user").mkdir(parents=True)
        # OOXMLRecalcMode=0 obliga a recalcular siempre al abrir.
        (perfil / "user" / "registrymodifications.xcu").write_text(
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
            'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
            '<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop>'
            "</item></oor:items>",
            encoding="utf-8",
        )
        resultado = Path(temporal) / "resultado"
        inicio = time.perf_counter()
        proceso = subprocess.run(
            [
                programa,
                "--headless",
                f"-env:UserInstallation={perfil.as_uri()}",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(resultado),
                str(copia),
            ],
            capture_output=True,
            timeout=1500,
            check=False,
        )
        convertido = resultado / copia.name
        print(f"\n== LibreOffice (recalculando, {time.perf_counter() - inicio:.0f} s)")
        if proceso.returncode != 0 or not convertido.exists():
            print(f"  No pudo abrir o convertir la copia. Código {proceso.returncode}.")
            return

        with zipfile.ZipFile(original) as zip_a, zipfile.ZipFile(convertido) as zip_b:
            rutas_a = rutas_de_hojas(zip_a)
            rutas_b = rutas_de_hojas(zip_b)
            print(f"  Hojas: original {len(rutas_a)}, LibreOffice {len(rutas_b)}")
            textos_a = leer_textos_compartidos(zip_a)
            textos_b = leer_textos_compartidos(zip_b)
            formulas = iguales = distintos = sin_valor = marcas = 0
            for fila_a, fila_b in emparejar_filas(
                recorrer_hoja(zip_a, rutas_a[indice], textos_a, {}),
                recorrer_hoja(zip_b, rutas_b[indice], textos_b, {}),
            ):
                celdas_a = fila_a.celdas if fila_a else {}
                celdas_b = fila_b.celdas if fila_b else {}
                for columna, celda_a in celdas_a.items():
                    if celda_a.formula is None or celda_a.valor is None:
                        continue
                    formulas += 1
                    celda_b = celdas_b.get(columna)
                    if celda_b is None or celda_b.valor is None:
                        sin_valor += 1
                    elif (
                        celda_a.valor[0] == "numero"
                        and celda_b.valor[0] == "numero"
                        and abs(celda_a.valor[1] - celda_b.valor[1]) < 0.005
                    ) or celda_a.valor == celda_b.valor:
                        iguales += 1
                    else:
                        distintos += 1
                numero = fila_b.numero if fila_b else None
                if numero in filas_marca:
                    celda_doc = celdas_b.get(columna_doc)
                    if celda_doc is not None and celda_doc.valor is not None:
                        marcas += 1
            print(f"  Fórmulas recalculadas: {formulas}; iguales al original: {iguales}; distintas: {distintos}; sin valor: {sin_valor}")
            print(f"  Marcas leídas: {marcas} de {len(filas_marca)}")


def huella(ruta):
    """Calcula la suma de comprobación de un archivo."""
    calculo = hashlib.sha256()
    with ruta.open("rb") as contenido:
        for bloque in iter(lambda: contenido.read(1024 * 1024), b""):
            calculo.update(bloque)
    return calculo.hexdigest()


def memoria_maxima_mb():
    """Devuelve la memoria máxima usada, o `None` si el sistema no la informa."""
    try:
        import resource  # Solo existe en Linux y macOS.
    except ImportError:
        return None
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024


def imprimir_lista(titulo, elementos):
    """Imprime una lista de nombres técnicos."""
    print(f"  {titulo} ({len(elementos)}): {', '.join(elementos) if elementos else '-'}")


def comparar_libros(origen, copia, hoja_principal):
    """Compara el original con la copia e imprime el resultado.

    Devuelve la cantidad de diferencias de contenido encontradas en las hojas.
    """
    indice_principal, columna_doc, columnas_marca, filas_marca = hoja_principal
    with zipfile.ZipFile(origen) as zip_a, zipfile.ZipFile(copia) as zip_b:
        print("\n== Partes internas del archivo")
        partes = resumen_partes(zip_a, zip_b)
        print(f"  Idénticas byte a byte: {len(partes['identicas'])}")
        imprimir_lista("Reescritas", partes["cambiadas"])
        imprimir_lista("Eliminadas", partes["faltantes"])
        imprimir_lista("Nuevas", partes["nuevas"])

        print("\n== Libro y estilos, por elemento")
        for parte, ignorar in (
            ("xl/workbook.xml", frozenset()),
            ("xl/styles.xml", ELEMENTOS_ESTILO_CELDA),
        ):
            faltan, sobran = comparar_elementos_superiores(
                zip_a, zip_b, parte, parte, ignorar
            )
            print(f"  {parte}: cambian o faltan {faltan or '-'}; nuevos o distintos {sobran or '-'}")

        print("\n== Tabla dinámica")
        for carpeta in ("xl/pivotTables/", "xl/pivotCache/"):
            xml_a = sorted(
                n for n in zip_a.namelist() if n.startswith(carpeta) and n.endswith(".xml")
            )
            xml_b = sorted(
                n for n in zip_b.namelist() if n.startswith(carpeta) and n.endswith(".xml")
            )
            print(f"  {carpeta}: original {len(xml_a)}, copia {len(xml_b)}")
            for parte_a, parte_b in zip(xml_a, xml_b):
                igual = canonizar(leer_xml(zip_a, parte_a)) == canonizar(
                    leer_xml(zip_b, parte_b)
                )
                print(f"    {PurePosixPath(parte_a).name}: {'igual' if igual else 'DISTINTA'}")

        rutas_a = rutas_de_hojas(zip_a)
        rutas_b = rutas_de_hojas(zip_b)
        if len(rutas_a) != len(rutas_b):
            print(f"\nCANTIDAD DE HOJAS DISTINTA: {len(rutas_a)} y {len(rutas_b)}")
            return 1

        datos = (
            leer_textos_compartidos(zip_a),
            leer_textos_compartidos(zip_b),
            leer_estilos(zip_a),
            leer_estilos(zip_b),
        )
        total_diferencias = 0
        for indice, (parte_a, parte_b) in enumerate(zip(rutas_a, rutas_b)):
            es_principal = indice == indice_principal
            marcas = (
                (set(filas_marca), columna_doc, columnas_marca)
                if es_principal
                else (set(), None, ())
            )
            resultado = comparar_hoja(zip_a, zip_b, parte_a, parte_b, datos, marcas)
            nombre = "principal" if es_principal else "secundaria"
            print(f"\n== Hoja {indice + 1} ({nombre})")
            print(f"  Filas: {resultado.filas}; celdas: {resultado.celdas}; fórmulas: {resultado.formulas}")
            print(f"  Fórmulas con valor calculado: original {resultado.calculados_original}, copia {resultado.calculados_copia}")
            if es_principal:
                print(f"  Marcas de texto correctas: {resultado.marcas_correctas} de {len(filas_marca)}")
                total_rellenos = len(filas_marca) * len(columnas_marca)
                print(f"  Celdas marcadas en amarillo: {resultado.rellenos_correctos} de {total_rellenos}")
            print(f"  Diferencias: {dict(resultado.diferencias) or 'ninguna'}")
            imprimir_lista("Elementos eliminados", resultado.elementos_faltantes)
            imprimir_lista("Elementos nuevos", resultado.elementos_nuevos)
            imprimir_lista("Elementos cambiados", resultado.elementos_cambiados)
            total_diferencias += sum(resultado.diferencias.values())
    return total_diferencias


def analizar_argumentos():
    """Define los argumentos de la línea de comandos."""
    analizador = argparse.ArgumentParser(
        description="Genera una copia de prueba del extracto y la compara con el original."
    )
    analizador.add_argument("extracto", type=Path, help="Archivo de extracto original")
    analizador.add_argument("salida", type=Path, help="Carpeta para la copia de prueba")
    analizador.add_argument("--metodo", choices=["openpyxl", "directo"], default="directo")
    analizador.add_argument(
        "--libreoffice",
        action="store_true",
        help="Recalcula la copia con LibreOffice y compara los saldos",
    )
    return analizador.parse_args()


def main():
    """Genera la copia, la compara y confirma que el original sigue igual."""
    argumentos = analizar_argumentos()
    origen = argumentos.extracto.resolve()
    salida = argumentos.salida.resolve()
    if salida == origen.parent:
        print("La carpeta de salida no puede ser la de los originales.")
        return 2
    salida.mkdir(parents=True, exist_ok=True)
    copia = salida / f"copia_{argumentos.metodo}.xlsx"
    if copia.exists():
        print("La copia de prueba ya existe; bórrala o usa otra carpeta.")
        return 2

    huella_inicial = huella(origen)
    with zipfile.ZipFile(origen) as archivo_zip:
        hoja_principal = encontrar_hoja_principal(archivo_zip)
    print(f"Hoja principal: posición {hoja_principal[0] + 1}. Marcas en {CANTIDAD_MARCAS} filas de plantilla.")

    inicio = time.perf_counter()
    if argumentos.metodo == "openpyxl":
        generar_copia_openpyxl(origen, copia, hoja_principal)
    else:
        generar_copia_directa(origen, copia, hoja_principal)
    duracion = time.perf_counter() - inicio
    memoria = memoria_maxima_mb()
    print(f"Copia generada en {duracion:.0f} s; tamaño {copia.stat().st_size / 1e6:.1f} MB.")
    if memoria is not None:
        print(f"Memoria máxima hasta aquí: {memoria:.0f} MB.")

    inicio = time.perf_counter()
    diferencias = comparar_libros(origen, copia, hoja_principal)
    print(f"\nComparación en {time.perf_counter() - inicio:.0f} s. Diferencias en hojas: {diferencias}.")

    if argumentos.libreoffice:
        comprobar_con_libreoffice(copia, origen, hoja_principal)

    original_intacto = huella(origen) == huella_inicial
    print(f"Original sin cambios: {'sí' if original_intacto else 'NO'}.")
    return 0 if original_intacto else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as error:  # Sin el mensaje: podría incluir rutas o datos.
        linea = error.__traceback__
        while linea.tb_next is not None:
            linea = linea.tb_next
        print(f"Error {type(error).__name__} en la línea {linea.tb_lineno}.")
        sys.exit(1)
