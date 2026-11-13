"""Pruebas de la escritura de la copia, con un `.xlsx` mínimo escrito a mano.

Enfoque: el archivo de prueba trae saldos ya calculados y una segunda hoja,
para comprobar que solo cambian las filas marcadas. Se reutiliza el extracto
mínimo de las pruebas de la etapa 1.
"""

import zipfile

import pytest

from conciliacion.escritura import COLOR_DUDOSA, COLOR_SEGURA, Marca, escribir_copia
from scripts.verificar_copia import leer_estilos
from tests.test_verificar_copia import HOJA_PRINCIPAL, crear_extracto

COLUMNAS = range(1, 8)


def leer_parte(ruta, parte):
    with zipfile.ZipFile(ruta) as archivo:
        return archivo.read(parte).decode()


def estilo_de(hoja, referencia):
    return int(hoja.split(f'r="{referencia}" s="')[1].split('"')[0])


@pytest.fixture
def original(tmp_path):
    ruta = tmp_path / "original.xlsx"
    crear_extracto(ruta)
    return ruta


def test_seguras_en_amarillo_y_dudosas_en_naranja(original, tmp_path):
    copia = tmp_path / "copia.xlsx"
    marcas = {3: Marca("CC-1-1", COLOR_SEGURA), 4: Marca("CC-1-2", COLOR_DUDOSA)}
    escribir_copia(original, copia, 0, 4, COLUMNAS, marcas)

    hoja = leer_parte(copia, "xl/worksheets/sheet1.xml")
    with zipfile.ZipFile(copia) as archivo:
        estilos = leer_estilos(archivo)
    assert "<t>CC-1-1</t>" in hoja and "<t>CC-1-2</t>" in hoja
    for columna in "ABCDEFG":
        assert COLOR_SEGURA in str(estilos[estilo_de(hoja, f"{columna}3")]["fill"])
        assert COLOR_DUDOSA in str(estilos[estilo_de(hoja, f"{columna}4")]["fill"])


def test_filas_sin_marca_y_otras_hojas_no_cambian(original, tmp_path):
    copia = tmp_path / "copia.xlsx"
    escribir_copia(original, copia, 0, 4, COLUMNAS, {3: Marca("CC-1-1", COLOR_SEGURA)})
    hoja = leer_parte(copia, "xl/worksheets/sheet1.xml")
    for fila in ("1", "2", "4", "5"):
        inicio = f'<row r="{fila}"'
        original_fila = HOJA_PRINCIPAL.split(inicio)[1].split("</row>")[0]
        assert original_fila == hoja.split(inicio)[1].split("</row>")[0]
    assert leer_parte(original, "xl/worksheets/sheet2.xml") == leer_parte(
        copia, "xl/worksheets/sheet2.xml"
    )


def test_conserva_saldos_calculados(original, tmp_path):
    copia = tmp_path / "copia.xlsx"
    escribir_copia(original, copia, 0, 4, COLUMNAS, {3: Marca("CC-1-1", COLOR_SEGURA)})
    assert leer_parte(copia, "xl/worksheets/sheet1.xml").count("<v>100.5</v>") == 5


def test_texto_con_caracteres_especiales_se_escapa(original, tmp_path):
    copia = tmp_path / "copia.xlsx"
    escribir_copia(original, copia, 0, 4, COLUMNAS, {3: Marca("A&B<1>", COLOR_SEGURA)})
    assert "<t>A&amp;B&lt;1&gt;</t>" in leer_parte(copia, "xl/worksheets/sheet1.xml")


def test_un_color_se_agrega_una_sola_vez(original, tmp_path):
    copia = tmp_path / "copia.xlsx"
    marcas = {fila: Marca("CC", COLOR_SEGURA) for fila in (3, 4, 5)}
    escribir_copia(original, copia, 0, 4, COLUMNAS, marcas)
    assert leer_parte(copia, "xl/styles.xml").count(COLOR_SEGURA) == 1


def test_no_sobrescribe_documento_y_no_crea_archivo(original, tmp_path):
    hoja = HOJA_PRINCIPAL.replace(
        '<row r="3" spans="1:7">',
        '<row r="3" spans="1:7"><c r="D3" t="inlineStr"><is><t>CC-9-9</t></is></c>',
    )
    crear_extracto(original, hoja)
    copia = tmp_path / "copia.xlsx"
    with pytest.raises(ValueError):
        escribir_copia(original, copia, 0, 4, COLUMNAS, {3: Marca("CC-1-1", COLOR_SEGURA)})
    assert not copia.exists()


def test_fila_inexistente_es_error(original, tmp_path):
    with pytest.raises(ValueError):
        escribir_copia(
            original, tmp_path / "copia.xlsx", 0, 4, COLUMNAS, {99: Marca("CC", COLOR_SEGURA)}
        )
