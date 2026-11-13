"""Pruebas de la copia directa y del comparador, con datos inventados.

Enfoque: cada prueba arma un `.xlsx` mínimo escrito a mano, para controlar
exactamente su contenido, incluidos los saldos ya calculados. `openpyxl` no
sirve para armarlo porque no guarda esos valores.

Las pruebas de control modifican la copia a propósito. Un comparador que nunca
encuentra diferencias también pasaría las demás pruebas, así que hay que
demostrar que sí detecta un daño.
"""

import zipfile

import pytest

from scripts.verificar_copia import (
    comparar_libros,
    encontrar_hoja_principal,
    generar_copia_directa,
    leer_estilos,
)

TIPOS = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/></Types>'
)
NS_HOJA = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
RELS_LIBRO = (
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="worksheet" Target="worksheets/sheet2.xml"/>'
    "</Relationships>"
)
LIBRO = (
    f'<?xml version="1.0"?><workbook {NS_HOJA} {NS_R}><sheets>'
    '<sheet name="Otra" sheetId="1" r:id="rId1"/>'
    '<sheet name="Resumen" sheetId="2" r:id="rId2"/>'
    "</sheets></workbook>"
)
ESTILOS = (
    f'<?xml version="1.0"?><styleSheet {NS_HOJA}>'
    '<fonts count="2"><font><sz val="11"/></font><font><b/><sz val="11"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border/></borders>'
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
    '<xf numFmtId="4" fontId="1" fillId="0" borderId="0" applyFont="1">'
    '<alignment horizontal="right"/></xf></cellXfs>'
    "</styleSheet>"
)
ENCABEZADOS = "".join(
    f'<c r="{columna}1" t="inlineStr"><is><t>{texto}</t></is></c>'
    for columna, texto in zip(
        "ABCDEFG",
        ["FECHA", "CONCEPTO", "DETALLE", "DOC CONCILIADO", " DB ", " CR ", " SALDOS "],
    )
)
HOJA_PRINCIPAL = (
    f'<?xml version="1.0"?><worksheet {NS_HOJA}><sheetData>'
    f'<row r="1" spans="1:7">{ENCABEZADOS}</row>'
    '<row r="2" spans="1:7"><c r="A2"><v>20260101</v></c>'
    '<c r="B2" t="inlineStr"><is><t>PAGO INVENTADO</t></is></c>'
    '<c r="E2" s="1"><v>100.5</v></c><c r="G2" s="1"><f>E2</f><v>100.5</v></c></row>'
    '<row r="3" spans="1:7"><c r="G3" s="1"><f>G2+E3-F3</f><v>100.5</v></c></row>'
    '<row r="4" spans="1:7"><c r="G4" s="1"><f>G3+E4-F4</f><v>100.5</v></c></row>'
    '<row r="5" spans="1:7"><c r="G5" s="1"><f>G4+E5-F5</f><v>100.5</v></c></row>'
    "</sheetData></worksheet>"
)
HOJA_RESUMEN = (
    f'<?xml version="1.0"?><worksheet {NS_HOJA}><sheetData>'
    '<row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>'
)


def crear_extracto(ruta, hoja=HOJA_PRINCIPAL):
    """Escribe un extracto mínimo, con la hoja principal en la primera posición."""
    with zipfile.ZipFile(ruta, "w") as archivo:
        archivo.writestr("[Content_Types].xml", TIPOS)
        archivo.writestr("xl/workbook.xml", LIBRO)
        archivo.writestr("xl/_rels/workbook.xml.rels", RELS_LIBRO)
        archivo.writestr("xl/styles.xml", ESTILOS)
        archivo.writestr("xl/worksheets/sheet1.xml", hoja)
        archivo.writestr("xl/worksheets/sheet2.xml", HOJA_RESUMEN)


def alterar_parte(ruta, parte, viejo, nuevo):
    """Cambia un texto dentro de una parte de la copia, para simular un daño."""
    with zipfile.ZipFile(ruta) as archivo:
        contenido = {nombre: archivo.read(nombre) for nombre in archivo.namelist()}
    texto = contenido[parte].decode()
    assert viejo in texto
    contenido[parte] = texto.replace(viejo, nuevo, 1).encode()
    with zipfile.ZipFile(ruta, "w") as archivo:
        for nombre, datos in contenido.items():
            archivo.writestr(nombre, datos)


@pytest.fixture
def extracto(tmp_path):
    """Entrega el original, la copia marcada y los datos de la hoja principal."""
    original = tmp_path / "original.xlsx"
    copia = tmp_path / "copia.xlsx"
    crear_extracto(original)
    with zipfile.ZipFile(original) as archivo:
        hoja_principal = encontrar_hoja_principal(archivo)
    generar_copia_directa(original, copia, hoja_principal)
    return original, copia, hoja_principal


def leer_parte(ruta, parte):
    with zipfile.ZipFile(ruta) as archivo:
        return archivo.read(parte)


def test_ubica_hoja_y_filas_vacias_por_encabezado(extracto):
    _, _, (indice, columna_doc, columnas, filas) = extracto
    assert indice == 0
    assert columna_doc == 4
    assert list(columnas) == [1, 2, 3, 4, 5, 6, 7]
    assert filas == [3, 4, 5]  # La última con fecha es la 2.


def test_solo_cambian_la_hoja_principal_y_los_estilos(extracto):
    original, copia, _ = extracto
    assert leer_parte(original, "xl/worksheets/sheet2.xml") == leer_parte(
        copia, "xl/worksheets/sheet2.xml"
    )
    assert leer_parte(original, "xl/workbook.xml") == leer_parte(copia, "xl/workbook.xml")
    assert leer_parte(original, "xl/worksheets/sheet1.xml") != leer_parte(
        copia, "xl/worksheets/sheet1.xml"
    )


def test_copia_correcta_no_tiene_diferencias(extracto, capsys):
    original, copia, hoja_principal = extracto
    assert comparar_libros(original, copia, hoja_principal) == 0
    salida = capsys.readouterr().out
    assert "Marcas de texto correctas: 3 de 3" in salida
    assert "Celdas marcadas en amarillo: 21 de 21" in salida


def test_conserva_saldos_calculados(extracto):
    _, copia, _ = extracto
    hoja = leer_parte(copia, "xl/worksheets/sheet1.xml").decode()
    assert hoja.count("<v>100.5</v>") == 5


def test_estilo_amarillo_conserva_fuente_y_formato(extracto):
    _, copia, _ = extracto
    hoja = leer_parte(copia, "xl/worksheets/sheet1.xml").decode()
    with zipfile.ZipFile(copia) as archivo:
        estilos = leer_estilos(archivo)
    indice = int(hoja.split('r="G3" s="')[1].split('"')[0])
    assert "FFFF00" in str(estilos[indice]["fill"])
    assert estilos[indice]["font"] == estilos[1]["font"]
    assert estilos[indice]["numFmt"] == estilos[1]["numFmt"]
    assert estilos[indice]["alignment"] == estilos[1]["alignment"]


def test_celdas_nuevas_quedan_en_orden_de_columna(extracto):
    _, copia, _ = extracto
    hoja = leer_parte(copia, "xl/worksheets/sheet1.xml").decode()
    fila = hoja.split('<row r="3"')[1].split("</row>")[0]
    posiciones = [fila.index(f'r="{letra}3"') for letra in "ABCDEFG"]
    assert posiciones == sorted(posiciones)


def test_no_sobrescribe_documento_con_contenido(tmp_path):
    original = tmp_path / "original.xlsx"
    crear_extracto(original)
    with zipfile.ZipFile(original) as archivo:
        hoja_principal = encontrar_hoja_principal(archivo)
    hoja = HOJA_PRINCIPAL.replace(
        '<row r="3" spans="1:7">',
        '<row r="3" spans="1:7"><c r="D3" t="inlineStr"><is><t>CC-1-1</t></is></c>',
    )
    crear_extracto(original, hoja)
    with pytest.raises(ValueError):
        generar_copia_directa(original, tmp_path / "copia.xlsx", hoja_principal)


@pytest.mark.parametrize(
    ("parte", "viejo", "nuevo", "diferencia"),
    [
        ("xl/worksheets/sheet1.xml", "<f>E2</f><v>100.5</v>", "<f>E2</f>", "valor_calculado_perdido"),
        ("xl/worksheets/sheet1.xml", "<f>E2</f>", "<f>F2</f>", "formula"),
        ("xl/worksheets/sheet1.xml", '<c r="E2" s="1"><v>100.5</v>', '<c r="E2" s="1"><v>999</v>', "valor"),
        ("xl/worksheets/sheet2.xml", "<v>7</v>", "<v>8</v>", "valor"),
    ],
)
def test_control_el_comparador_detecta_danos(extracto, capsys, parte, viejo, nuevo, diferencia):
    original, copia, hoja_principal = extracto
    alterar_parte(copia, parte, viejo, nuevo)
    assert comparar_libros(original, copia, hoja_principal) > 0
    assert diferencia in capsys.readouterr().out


def test_control_relleno_sin_tipo_equivale_a_ninguno(tmp_path):
    ruta = tmp_path / "estilos.xlsx"
    crear_extracto(ruta)
    alterar_parte(ruta, "xl/styles.xml", '<patternFill patternType="none"/>', "<patternFill/>")
    otra = tmp_path / "otra.xlsx"
    crear_extracto(otra)
    with zipfile.ZipFile(ruta) as a, zipfile.ZipFile(otra) as b:
        assert leer_estilos(a)[0]["fill"] == leer_estilos(b)[0]["fill"]
