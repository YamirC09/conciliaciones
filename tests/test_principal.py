"""Prueba de punta a punta del comando principal, con archivos inventados.

Enfoque: se arman un extracto y unos movimientos donde ya sabemos qué debe
pasar con cada fila (segura, dudosa, pendiente, ya conciliada). Se corre el
comando completo y se abren los archivos generados para comprobar qué quedó
escrito y de qué color, como lo haría ella.
"""

import pytest
from openpyxl import Workbook, load_workbook

from conciliacion.__main__ import SalidaExistente, conciliar, huella, nombre_base

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
AMARILLO = "FFFFFF00"
AZUL = "FF9BC2E6"
NARANJA = "FFFFC000"


def crear_archivos(carpeta):
    """Crea un extracto y unos movimientos inventados.

    - Fila 3: ya conciliada por ella.
    - Fila 4: una sola pareja exacta: segura.
    - Fila 5 y 6: mismo monto y fecha, dos terceros que el nombre distingue: seguras.
    - Fila 7: $500 de diferencia, pero es el más cercano: segura.
    - Fila 8: comisión sin movimiento: pendiente.
    - Fila 9 y 10: mismo monto y fecha, y el nombre no distingue: dudosas.
    - Fila 11 y 12: un mismo comprobante con dos líneas de terceros distintos: azules.
    """
    extracto = Workbook()
    hoja = extracto.active
    hoja.append(ENCABEZADOS_EXTRACTO)
    hoja.append([20261001, "SALDO", None, None, 1000, None, 1000])
    filas = [
        [20261001, "PAGO", "CLIENTE UNO", "CC-1-1", None, 100],
        [20261002, "PAGO", "CLIENTE DOS", None, None, 200],
        [20261003, "PAGO", "CLIENTE TRES", None, None, 300],
        [20261003, "PAGO", "CLIENTE CUATRO", None, None, 300],
        [20261004, "PAGO", "CLIENTE CINCO", None, None, 1000],
        [20261005, "COMISION", None, None, 15, None],
        [20261006, "TRANSFERENCIA", "REF 111", None, None, 400],
        [20261006, "TRANSFERENCIA", "REF 222", None, None, 400],
        [20261007, "PAGO", "PERSONA SEIS", None, None, 500],
        [20261007, "PAGO", "PERSONA SIETE", None, None, 5000],
    ]
    for numero, fila in enumerate(filas, start=3):
        hoja.append([*fila, f"=G{numero - 1}+E{numero}-F{numero}"])
    extracto.create_sheet("Otra").append(["DATO", 1])

    movimientos = Workbook()
    hoja = movimientos.active
    hoja.append(ENCABEZADOS_MOVIMIENTOS)
    for datos in [
        ["CC-1-1", 1, "01-10-2026", 1, "CLIENTE UNO", "PAGO", "", None, 100],
        ["CC-1-2", 1, "02-10-2026", 2, "CLIENTE DOS", "PAGO", "", None, 200],
        ["CC-1-3", 1, "03-10-2026", 3, "CLIENTE TRES", "PAGO", "", None, 300],
        ["CC-1-4", 1, "03-10-2026", 4, "CLIENTE CUATRO", "PAGO", "", None, 300],
        ["CC-1-5", 1, "04-10-2026", 5, "CLIENTE CINCO", "PAGO", "", None, 1500],
        ["CC-1-6", 1, "09-10-2026", 6, "SIN PAREJA", "PAGO", "", None, 999],
        ["CC-1-7", 1, "06-10-2026", 7, "PROVEEDOR UNO", "PAGO", "", None, 400],
        ["CC-1-8", 1, "06-10-2026", 8, "PROVEEDOR DOS", "PAGO", "", None, 400],
        ["CC-1-9", 1, "07-10-2026", 9, "PERSONA SEIS", "PAGO", "", None, 500],
        ["CC-1-9", 2, "07-10-2026", 10, "PERSONA SIETE", "PAGO", "", None, 5000],
    ]:
        hoja.append(datos)

    ruta_extracto = carpeta / "extracto inventado.xlsx"
    ruta_movimientos = carpeta / "movimientos inventados.xlsx"
    extracto.save(ruta_extracto)
    movimientos.save(ruta_movimientos)
    return ruta_extracto, ruta_movimientos


@pytest.fixture
def ejecucion(tmp_path):
    ruta_extracto, ruta_movimientos = crear_archivos(tmp_path)
    huellas = (huella(ruta_extracto), huella(ruta_movimientos))
    resumen = conciliar(ruta_extracto, ruta_movimientos, tmp_path / "salida")
    return ruta_extracto, ruta_movimientos, huellas, resumen


def test_resumen_cuenta_cada_tipo_de_fila(ejecucion):
    *_, resumen = ejecucion
    assert resumen.ya_conciliadas == 1
    assert resumen.seguras == 4
    assert resumen.repetidas == 2
    assert resumen.dudosas == 2
    assert resumen.pendientes == 1
    assert resumen.sin_pareja == 1


def test_copia_tiene_comprobantes_y_colores(ejecucion):
    *_, resumen = ejecucion
    hoja = load_workbook(resumen.copia).worksheets[0]
    assert hoja["D3"].value == "CC-1-1"  # Ya estaba: no se toca.
    assert hoja["D3"].fill.fgColor.rgb != AMARILLO
    assert hoja["D4"].value == "CC-1-2"
    assert hoja["D5"].value == "CC-1-3"  # Elegidos por el nombre.
    assert hoja["D6"].value == "CC-1-4"
    assert hoja["D7"].value == "CC-1-5"
    assert hoja["D8"].value is None
    assert {hoja["D9"].value, hoja["D10"].value} == {"CC-1-7", "CC-1-8"}
    assert hoja["D11"].value == hoja["D12"].value == "CC-1-9"
    colores = {
        **dict.fromkeys((4, 5, 6, 7), AMARILLO),
        **dict.fromkeys((9, 10), NARANJA),
        **dict.fromkeys((11, 12), AZUL),
    }
    for columna in "ABCDEFG":
        for fila, color in colores.items():
            assert hoja[f"{columna}{fila}"].fill.fgColor.rgb == color
        assert hoja[f"{columna}8"].fill.fgColor.rgb not in {AMARILLO, AZUL, NARANJA}


def test_copia_conserva_formulas_y_otras_hojas(ejecucion):
    *_, resumen = ejecucion
    libro = load_workbook(resumen.copia)
    assert libro.worksheets[0]["G4"].value == "=G3+E4-F4"
    assert libro.sheetnames[1] == "Otra"


def test_revision_lista_candidatos_y_pendientes(ejecucion):
    *_, resumen = ejecucion
    libro = load_workbook(resumen.revision)
    assert libro.sheetnames == [
        "Resumen",
        "Para revisar",
        "Pendientes extracto",
        "Movimientos sin pareja",
        "Filas no leídas",
    ]
    por_revisar = list(libro["Para revisar"].iter_rows(min_row=2, values_only=True))
    # Filas 9 y 10 tienen dos candidatos cada una; la 11 y la 12, uno.
    assert len(por_revisar) == 6
    recomendados = {fila[0]: fila for fila in por_revisar if fila[7].startswith("Recomendado")}
    assert sorted(recomendados) == [9, 10, 11, 12]
    assert recomendados[11][1] == "Azul"
    assert "otra fila con otro tercero" in recomendados[11][16]
    assert recomendados[11][17] == "12"
    assert recomendados[9][1] == "Naranja"
    assert "no permite elegir" in recomendados[9][16]

    pendientes = list(libro["Pendientes extracto"].iter_rows(min_row=2, values_only=True))
    assert [fila[0] for fila in pendientes] == [8]
    sin_pareja = list(libro["Movimientos sin pareja"].iter_rows(min_row=2, values_only=True))
    assert [fila[1] for fila in sin_pareja] == ["CC-1-6"]


def test_originales_no_cambian(ejecucion):
    ruta_extracto, ruta_movimientos, huellas, _ = ejecucion
    assert (huella(ruta_extracto), huella(ruta_movimientos)) == huellas


def test_no_sobrescribe_una_salida_existente(ejecucion, tmp_path):
    ruta_extracto, ruta_movimientos, _, resumen = ejecucion
    contenido = resumen.copia.read_bytes()
    with pytest.raises(SalidaExistente):
        conciliar(ruta_extracto, ruta_movimientos, tmp_path / "salida")
    assert resumen.copia.read_bytes() == contenido


def test_copia_de_movimientos_pinta_cada_movimiento_usado_del_color_de_su_fila(ejecucion):
    _, ruta_movimientos, _, resumen = ejecucion
    copia = load_workbook(resumen.copia_movimientos).worksheets[0]
    original = load_workbook(ruta_movimientos).worksheets[0]
    # Fila de movimientos: color esperado (la 2 es del comprobante que ella ya
    # había puesto y la 7 no tiene pareja: quedan sin color).
    colores = {
        **dict.fromkeys((3, 4, 5, 6), AMARILLO),
        **dict.fromkeys((8, 9), NARANJA),
        **dict.fromkeys((10, 11), AZUL),
    }
    for fila in range(1, 12):
        for columna in "ABCDEFGHI":
            celda = copia[f"{columna}{fila}"]
            assert celda.value == original[f"{columna}{fila}"].value  # No escribe texto.
            if fila in colores:
                assert celda.fill.fgColor.rgb == colores[fila]
            else:
                assert celda.fill.fgColor.rgb not in {AMARILLO, AZUL, NARANJA}


def test_no_deja_archivos_temporales(ejecucion, tmp_path):
    archivos = sorted(ruta.name for ruta in (tmp_path / "salida").iterdir())
    assert archivos == [
        "extracto inventado_conciliado.xlsx",
        "extracto inventado_revision.xlsx",
        "movimientos inventados_conciliado.xlsx",
    ]


@pytest.mark.parametrize(
    ("archivo", "esperado"),
    [
        ("extracto.xlsx", "extracto"),
        ("extracto_conciliado.xlsx", "extracto"),
        ("extracto_conciliado_conciliado.xlsx", "extracto"),
        ("extracto_CONCILIADO.xlsx", "extracto"),
        ("extracto conciliado.xlsx", "extracto conciliado"),  # Sin guion bajo: es su nombre.
        ("_conciliado.xlsx", "_conciliado"),  # No deja el nombre vacío.
    ],
)
def test_nombre_base_quita_el_sufijo_solo_del_final(archivo, esperado):
    assert nombre_base(archivo) == esperado


def test_extracto_ya_conciliado_no_repite_el_sufijo(tmp_path):
    # Ella concilia el archivo que salió el día anterior.
    ruta_extracto, ruta_movimientos = crear_archivos(tmp_path)
    de_ayer = ruta_extracto.rename(tmp_path / "extracto inventado_conciliado.xlsx")
    resumen = conciliar(de_ayer, ruta_movimientos, tmp_path / "hoy")
    assert resumen.copia.name == "extracto inventado_conciliado.xlsx"
    assert resumen.revision.name == "extracto inventado_revision.xlsx"
