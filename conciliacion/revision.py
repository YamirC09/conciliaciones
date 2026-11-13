"""Crea el archivo de revisión: lo que ella debe mirar después del cruce.

Es un Excel nuevo, aparte de la copia del extracto. Tiene una hoja por tema:

- Resumen: conteos de la ejecución.
- Para revisar: cada fila azul o naranja con todos sus candidatos; el
  recomendado es el que quedó escrito en la copia y va del mismo color.
- Pendientes extracto: filas sin comprobante y por qué.
- Movimientos sin pareja: movimientos contables que no se usaron.
- Filas no leídas: filas que parecían movimientos pero no se pudieron leer.

Este archivo contiene datos reales, porque es para ella. Se guarda fuera del
repositorio, en la carpeta de salida.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from conciliacion.cruce import Estado, lado_y_monto
from conciliacion.escritura import COLOR_DUDOSA, COLOR_REPETIDA

FORMATO_MONTO = "#,##0.00"
POR_REVISAR = {
    Estado.REPETIDA: ("Azul", COLOR_REPETIDA),
    Estado.DUDOSA: ("Naranja", COLOR_DUDOSA),
}
FORMATO_FECHA = "dd/mm/yyyy"
ANCHO_MAXIMO = 45
NOMBRES_LADO = {"debito": "DB", "credito": "CR"}

ENCABEZADOS_DUDOSAS = [
    "Fila extracto",
    "Color",
    "Fecha",
    "Lado",
    "Monto extracto",
    "Concepto",
    "Detalle extracto",
    "Candidato",
    "Fila movimientos",
    "Comprobante",
    "Secuencia",
    "Monto movimiento",
    "Diferencia",
    "Tercero",
    "Descripción",
    "Usado en fila del extracto",
    "Motivo",
    "Mismo comprobante en filas",
]
ENCABEZADOS_PENDIENTES = [
    "Fila extracto",
    "Fecha",
    "Lado",
    "Monto",
    "Concepto",
    "Detalle",
    "Motivo",
    "Filas de movimientos candidatas",
]
ENCABEZADOS_SIN_PAREJA = [
    "Fila movimientos",
    "Comprobante",
    "Secuencia",
    "Fecha",
    "Lado",
    "Monto",
    "Tercero",
    "Descripción",
    "Detalle",
]
ENCABEZADOS_NO_LEIDAS = ["Archivo", "Fila", "Motivo"]


def a_pesos(centavos):
    """Convierte centavos en un número con decimales para Excel."""
    return centavos / 100


def lado_y_pesos(fila):
    """Devuelve el lado (DB o CR) y el monto en pesos de una fila del extracto."""
    datos = lado_y_monto(fila)
    if datos is None:
        return "DB y CR", None
    lado, monto = datos
    return NOMBRES_LADO[lado], a_pesos(monto)


def lado_de_movimiento(movimiento):
    """Devuelve el lado y el monto en pesos de un movimiento contable."""
    if movimiento.debito and movimiento.credito:
        return "Débito y crédito", None
    if movimiento.debito:
        return "Débito", a_pesos(movimiento.debito)
    return "Crédito", a_pesos(movimiento.credito)


def texto_de_motivos(propuesta):
    """Une los motivos de una propuesta en un solo texto."""
    return "; ".join(motivo.value for motivo in propuesta.motivos)


def filas_de_dudosas(resultado):
    """Arma una fila por cada candidato de cada fila azul o naranja.

    El recomendado va primero, porque es el comprobante que quedó escrito.
    """
    asignado_a = {
        p.movimiento.fila: p.fila.fila for p in resultado.propuestas if p.movimiento
    }
    filas = []
    for propuesta in resultado.propuestas:
        if propuesta.estado not in POR_REVISAR:
            continue
        color = POR_REVISAR[propuesta.estado][0]
        fila = propuesta.fila
        lado, monto = lado_y_pesos(fila)
        nombre_lado = lado_y_monto(fila)[0]
        otros = [m for m in propuesta.candidatos if m.fila != propuesta.movimiento.fila]
        for orden, movimiento in enumerate([propuesta.movimiento, *otros]):
            monto_movimiento = a_pesos(getattr(movimiento, nombre_lado))
            usado_en = asignado_a.get(movimiento.fila)
            filas.append(
                [
                    fila.fila,
                    color,
                    fila.fecha,
                    lado,
                    monto,
                    fila.concepto,
                    fila.detalle,
                    "Recomendado (escrito)" if orden == 0 else "Alternativa",
                    movimiento.fila,
                    movimiento.comprobante,
                    movimiento.secuencia,
                    monto_movimiento,
                    round(monto_movimiento - monto, 2),
                    movimiento.tercero,
                    movimiento.descripcion,
                    usado_en if usado_en not in (None, fila.fila) else None,
                    texto_de_motivos(propuesta),
                    ", ".join(map(str, propuesta.otras_filas)) or None,
                ]
            )
    return filas


def filas_de_pendientes(resultado):
    """Arma una fila por cada fila del extracto que quedó sin comprobante."""
    filas = []
    for propuesta in resultado.propuestas:
        if propuesta.estado is not Estado.PENDIENTE:
            continue
        fila = propuesta.fila
        lado, monto = lado_y_pesos(fila)
        candidatas = ", ".join(str(m.fila) for m in propuesta.candidatos)
        filas.append(
            [
                fila.fila,
                fila.fecha,
                lado,
                monto,
                fila.concepto,
                fila.detalle,
                texto_de_motivos(propuesta),
                candidatas or None,
            ]
        )
    return filas


def filas_sin_pareja(resultado):
    """Arma una fila por cada movimiento contable que no se usó."""
    filas = []
    for movimiento in resultado.movimientos_sin_pareja:
        lado, monto = lado_de_movimiento(movimiento)
        filas.append(
            [
                movimiento.fila,
                movimiento.comprobante,
                movimiento.secuencia,
                movimiento.fecha,
                lado,
                monto,
                movimiento.tercero,
                movimiento.descripcion,
                movimiento.detalle,
            ]
        )
    return filas


def agregar_hoja(libro, titulo, encabezados, filas):
    """Agrega una hoja con encabezado en negrita, filtro y paneles congelados."""
    hoja = libro.create_sheet(titulo)
    hoja.append(encabezados)
    for celda in hoja[1]:
        celda.font = Font(bold=True)
    for fila in filas:
        hoja.append(fila)

    for numero, encabezado in enumerate(encabezados, start=1):
        letra = get_column_letter(numero)
        largo = max(
            [len(encabezado)] + [len(str(f[numero - 1] or "")) for f in filas]
        )
        hoja.column_dimensions[letra].width = min(largo + 2, ANCHO_MAXIMO)
        for celda in hoja[letra][1:]:
            if encabezado.startswith(("Monto", "Diferencia")):
                celda.number_format = FORMATO_MONTO
            elif encabezado == "Fecha":
                celda.number_format = FORMATO_FECHA

    hoja.freeze_panes = "A2"
    if filas:
        hoja.auto_filter.ref = hoja.dimensions
    return hoja


def resaltar_recomendados(hoja):
    """Pinta las filas recomendadas del mismo color que en la copia."""
    rellenos = {
        nombre: PatternFill("solid", start_color=color, end_color=color)
        for nombre, color in POR_REVISAR.values()
    }
    columna_candidato = ENCABEZADOS_DUDOSAS.index("Candidato")
    columna_color = ENCABEZADOS_DUDOSAS.index("Color")
    for fila in hoja.iter_rows(min_row=2):
        if str(fila[columna_candidato].value).startswith("Recomendado"):
            for celda in fila:
                celda.fill = rellenos[fila[columna_color].value]


def escribir_revision(destino, resultado, advertencias_extracto, advertencias_movimientos):
    """Crea el archivo de revisión y devuelve cuántas filas tiene cada hoja."""
    dudosas = filas_de_dudosas(resultado)
    pendientes = filas_de_pendientes(resultado)
    sin_pareja = filas_sin_pareja(resultado)
    no_leidas = [["Extracto", fila, motivo] for fila, motivo in advertencias_extracto]
    no_leidas += [
        ["Movimientos", fila, motivo] for fila, motivo in advertencias_movimientos
    ]

    estados = [propuesta.estado for propuesta in resultado.propuestas]
    resumen = [
        ["Filas que ya estaban conciliadas (no se tocaron)", resultado.filas_ya_conciliadas],
        ["Seguras (amarillo)", estados.count(Estado.SEGURA)],
        ["Comprobante repetido con otro tercero (azul, revisar)", estados.count(Estado.REPETIDA)],
        ["Dudosas (naranja, revisar)", estados.count(Estado.DUDOSA)],
        ["Pendientes del extracto (sin comprobante)", estados.count(Estado.PENDIENTE)],
        ["Movimientos sin pareja", len(sin_pareja)],
        ["Filas no leídas", len(no_leidas)],
        [
            "Comprobantes ya escritos que no están en movimientos",
            resultado.documentos_sin_movimiento,
        ],
    ]

    libro = Workbook()
    libro.remove(libro.active)
    agregar_hoja(libro, "Resumen", ["Concepto", "Cantidad"], resumen)
    resaltar_recomendados(agregar_hoja(libro, "Para revisar", ENCABEZADOS_DUDOSAS, dudosas))
    agregar_hoja(libro, "Pendientes extracto", ENCABEZADOS_PENDIENTES, pendientes)
    agregar_hoja(libro, "Movimientos sin pareja", ENCABEZADOS_SIN_PAREJA, sin_pareja)
    agregar_hoja(libro, "Filas no leídas", ENCABEZADOS_NO_LEIDAS, no_leidas)
    libro.save(destino)
    return {
        "dudosas": len(dudosas),
        "pendientes": len(pendientes),
        "sin_pareja": len(sin_pareja),
        "no_leidas": len(no_leidas),
    }
