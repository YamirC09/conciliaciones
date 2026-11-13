"""Pruebas de la lectura, con archivos de Excel inventados dentro de cada prueba.

Enfoque: las conversiones (montos y fechas) se prueban directo, con varios
valores. La lectura se prueba creando un Excel pequeño en una carpeta temporal,
y cada prueba comprueba una sola regla.
"""

from datetime import date, datetime

import pytest
from openpyxl import Workbook

from conciliacion.lectura import (
    ArchivoInvalido,
    a_centavos,
    a_texto,
    fecha_de_extracto,
    fecha_de_movimientos,
    leer_extracto,
    leer_movimientos,
)

ENCABEZADOS_EXTRACTO = ["FECHA", "CONCEPTO", "DETALLE", "DOC CONCILIADO", " DB ", " CR ", " SALDOS "]
ENCABEZADOS_MOVIMIENTOS = [
    "Comprobante",
    "Secuencia",
    "Fecha Elaboración",
    "Identificación",
    "Nombre tercero",
    "Descripcion",
    "Detalle",
    "Débito",
    "Crédito",
]


def crear_libro(ruta, *hojas):
    """Crea un Excel con una lista de filas por cada hoja."""
    libro = Workbook()
    libro.remove(libro.active)
    for indice, filas in enumerate(hojas):
        hoja = libro.create_sheet(f"Hoja{indice}")
        for fila in filas:
            hoja.append(fila)
    libro.save(ruta)
    return ruta


def fila_movimiento(comprobante="CC-1-1", secuencia=1, fecha="15-10-2026", debito=None, credito=1000.5):
    """Arma una fila de movimientos inventada."""
    return [comprobante, secuencia, fecha, 900, "TERCERO FICTICIO", "PAGO", "DETALLE", debito, credito]


def leer_filas_extracto(tmp_path, *filas):
    """Crea un extracto con las filas indicadas y lo lee."""
    ruta = crear_libro(tmp_path / "extracto.xlsx", [ENCABEZADOS_EXTRACTO, *filas])
    return leer_extracto(ruta)


@pytest.mark.parametrize(
    ("valor", "esperado"),
    [
        (None, 0),
        ("", 0),
        (100, 10000),
        (1000.5, 100050),
        (0.1 + 0.2, 30),  # En float da 0.30000000000000004.
        (1234.565, 123457),  # El medio centavo redondea hacia arriba.
        ("250.75", 25075),
        (-5.5, -550),
    ],
)
def test_montos_quedan_en_centavos_enteros(valor, esperado):
    assert a_centavos(valor) == esperado


@pytest.mark.parametrize("valor", ["abc", "1.234,56", True])
def test_monto_ambiguo_o_invalido_se_rechaza(valor):
    with pytest.raises(ValueError):
        a_centavos(valor)


@pytest.mark.parametrize(
    "valor", [20261015, 20261015.0, "20261015", datetime(2026, 10, 15), date(2026, 10, 15)]
)
def test_fecha_de_extracto_en_sus_formatos(valor):
    assert fecha_de_extracto(valor) == date(2026, 10, 15)


@pytest.mark.parametrize("valor", ["15-10-2026", " 15-10-2026 ", datetime(2026, 10, 15)])
def test_fecha_de_movimientos_en_sus_formatos(valor):
    assert fecha_de_movimientos(valor) == date(2026, 10, 15)


@pytest.mark.parametrize("valor", [20261345, "31-02-2026", "2026-10-15", "ayer"])
def test_fecha_invalida_se_rechaza(valor):
    funcion = fecha_de_extracto if isinstance(valor, int) else fecha_de_movimientos
    with pytest.raises(ValueError):
        funcion(valor)


def test_codigo_guardado_como_numero_no_conserva_decimal():
    assert a_texto(900123.0) == "900123"


def test_extracto_ubica_columnas_por_encabezado_aunque_cambien_de_lugar(tmp_path):
    ruta = crear_libro(
        tmp_path / "extracto.xlsx",
        [
            [" CR ", "DOC CONCILIADO", "fecha", "Concepto", "Detalle", " DB "],
            [500, "CC-1-1", 20261015, "PAGO", "CLIENTE", None],
        ],
    )
    fila = leer_extracto(ruta).movimientos[0]
    assert fila.fecha == date(2026, 10, 15)
    assert fila.credito == 50000
    assert fila.debito == 0
    assert fila.documento == "CC-1-1"


def test_extracto_encuentra_la_hoja_principal_aunque_no_sea_la_primera(tmp_path):
    ruta = crear_libro(
        tmp_path / "extracto.xlsx",
        [["OTRA", "TABLA"], [1, 2]],
        [ENCABEZADOS_EXTRACTO, [20261015, "PAGO", "X", None, 100, None]],
    )
    lectura = leer_extracto(ruta)
    assert lectura.posicion_hoja == 2
    assert len(lectura.movimientos) == 1


def test_extracto_excluye_saldo_inicial_y_filas_de_plantilla(tmp_path):
    lectura = leer_filas_extracto(
        tmp_path,
        [20261001, "SALDO", None, None, 5000, None, 5000],
        [20261002, "PAGO", "X", None, None, 100, "=G2-F3"],
        [None, None, None, None, None, None, "=G3-F4"],
    )
    assert [fila.concepto for fila in lectura.movimientos] == ["PAGO"]
    assert lectura.ignoradas == {"saldo inicial": 1, "sin fecha ni monto": 1}
    assert lectura.advertencias == []


def test_extracto_conserva_numero_de_fila_de_origen(tmp_path):
    lectura = leer_filas_extracto(
        tmp_path,
        [20261001, "PAGO", None, None, None, 100],
        [],  # Una fila vacía no debe correr la numeración.
        [20261002, "PAGO", None, None, None, 200],
    )
    assert [fila.fila for fila in lectura.movimientos] == [2, 4]


def test_extracto_distingue_conciliados(tmp_path):
    lectura = leer_filas_extracto(
        tmp_path,
        [20261001, "PAGO", None, "CC-1-1", None, 100],
        [20261002, "PAGO", None, None, None, 200],
    )
    assert [fila.conciliado for fila in lectura.movimientos] == [True, False]


@pytest.mark.parametrize(
    ("fila", "motivo"),
    [
        ([20261001, "PAGO", None, None, None, None], "fecha sin monto"),
        ([None, "PAGO", None, None, None, 100], "monto sin fecha"),
        ([20269999, "PAGO", None, None, None, 100], "fecha no reconocida"),
        ([20261001, "PAGO", None, None, "abc", None], "monto no reconocido"),
    ],
)
def test_extracto_avisa_filas_dudosas_sin_incluirlas(tmp_path, fila, motivo):
    lectura = leer_filas_extracto(tmp_path, fila)
    assert lectura.movimientos == []
    assert lectura.advertencias == [(2, motivo)]


def test_extracto_avisa_monto_en_ambos_lados_y_lo_conserva(tmp_path):
    lectura = leer_filas_extracto(tmp_path, [20261001, "PAGO", None, None, 100, 100])
    assert len(lectura.movimientos) == 1
    assert lectura.advertencias == [(2, "monto en DB y en CR")]


def test_extracto_sin_encabezados_es_archivo_invalido(tmp_path):
    ruta = crear_libro(tmp_path / "otro.xlsx", [["A", "B"], [1, 2]])
    with pytest.raises(ArchivoInvalido):
        leer_extracto(ruta)


def test_movimientos_exportacion_con_titulos_y_pie(tmp_path):
    ruta = crear_libro(
        tmp_path / "movimientos.xlsx",
        [
            ["TITULO INVENTADO"],
            ["SUBTITULO"],
            [],
            [],
            ENCABEZADOS_MOVIMIENTOS,
            fila_movimiento(),
            [None, None, None, None, None, None, "Totales", 0, 1000.5],
            ["Impreso por sistema ficticio"],
        ],
    )
    lectura = leer_movimientos(ruta)
    assert len(lectura.movimientos) == 1
    assert lectura.movimientos[0].fila == 6
    assert lectura.movimientos[0].credito == 100050
    assert lectura.movimientos[0].fecha == date(2026, 10, 15)


def test_movimientos_version_limpia_con_encabezado_en_fila_uno(tmp_path):
    ruta = crear_libro(tmp_path / "movimientos.xlsx", [ENCABEZADOS_MOVIMIENTOS, fila_movimiento()])
    assert leer_movimientos(ruta).movimientos[0].fila == 2


def test_movimientos_usa_la_primera_hoja_aunque_otra_tenga_mas_filas(tmp_path):
    ruta = crear_libro(
        tmp_path / "movimientos.xlsx",
        [ENCABEZADOS_MOVIMIENTOS, fila_movimiento()],
        [ENCABEZADOS_MOVIMIENTOS, fila_movimiento(), fila_movimiento("CC-1-2")],
    )
    lectura = leer_movimientos(ruta)
    assert lectura.posicion_hoja == 1
    assert len(lectura.movimientos) == 1


def test_movimientos_no_busca_en_otra_hoja_si_la_primera_no_sirve(tmp_path):
    ruta = crear_libro(
        tmp_path / "movimientos.xlsx",
        [["A", "B"], [1, 2]],
        [ENCABEZADOS_MOVIMIENTOS, fila_movimiento()],
    )
    with pytest.raises(ArchivoInvalido):
        leer_movimientos(ruta)


def test_movimientos_avisa_fecha_invalida_sin_incluirla(tmp_path):
    ruta = crear_libro(
        tmp_path / "movimientos.xlsx",
        [ENCABEZADOS_MOVIMIENTOS, fila_movimiento(fecha="sin fecha")],
    )
    lectura = leer_movimientos(ruta)
    assert lectura.movimientos == []
    assert lectura.advertencias == [(2, "fecha no reconocida")]


def test_movimientos_sin_secuencia_ni_detalle_se_leen(tmp_path):
    # Otro banco exporta sin esas dos columnas: solo se muestran en la revisión.
    encabezados = [e for e in ENCABEZADOS_MOVIMIENTOS if e not in ("Secuencia", "Detalle")]
    fila = ["CC-1-1", "15-10-2026", 900, "TERCERO FICTICIO", "PAGO", None, 1000.5]
    lectura = leer_movimientos(crear_libro(tmp_path / "m.xlsx", [encabezados, fila]))
    movimiento = lectura.movimientos[0]
    assert (movimiento.secuencia, movimiento.detalle) == ("", "")
    assert movimiento.tercero == "TERCERO FICTICIO"
    assert movimiento.credito == 100050


@pytest.mark.parametrize("faltante", ["Comprobante", "Fecha Elaboración", "Nombre tercero", "Débito"])
def test_movimientos_sin_una_columna_obligatoria_es_archivo_invalido(tmp_path, faltante):
    encabezados = [e for e in ENCABEZADOS_MOVIMIENTOS if e != faltante]
    ruta = crear_libro(tmp_path / "m.xlsx", [encabezados, fila_movimiento()[:-1]])
    with pytest.raises(ArchivoInvalido):
        leer_movimientos(ruta)


def test_movimientos_sin_encabezados_es_archivo_invalido(tmp_path):
    ruta = crear_libro(tmp_path / "otro.xlsx", [["A", "B"], [1, 2]])
    with pytest.raises(ArchivoInvalido):
        leer_movimientos(ruta)
