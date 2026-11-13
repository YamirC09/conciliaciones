"""Pruebas de la lógica de la ventana, sin abrir ventanas.

Enfoque: la ventana solo muestra lo que calculan estas funciones. Si ellas
están bien, la ventana sugiere la carpeta correcta, avisa lo que falta y
traduce los errores a mensajes claros. Ningún mensaje de un error inesperado
debe mostrar su texto interno, porque podría tener rutas o datos.
"""

import zipfile
from datetime import datetime
from pathlib import Path

from conciliacion.__main__ import Resumen, SalidaExistente
from conciliacion.escritura import COLOR_DUDOSA, COLOR_REPETIDA, COLOR_SEGURA
from conciliacion.interfaz import (
    carpeta_de_salida,
    lineas_de_resumen,
    mensaje_de_error,
    validar_seleccion,
)
from conciliacion.lectura import ArchivoInvalido


def test_carpeta_nueva_junto_al_extracto_con_fecha_y_hora():
    carpeta = carpeta_de_salida(Path("banco") / "extracto.xlsx", datetime(2026, 11, 11, 9, 5, 7))
    assert carpeta == Path("banco") / "Conciliaciones" / "2026-11-11 09-05-07"


def test_dos_ejecuciones_seguidas_usan_carpetas_distintas():
    primera = carpeta_de_salida("extracto.xlsx", datetime(2026, 11, 11, 9, 5, 7))
    segunda = carpeta_de_salida("extracto.xlsx", datetime(2026, 11, 11, 9, 5, 8))
    assert primera != segunda


def test_avisa_que_falta_cada_archivo(tmp_path):
    archivo = tmp_path / "a.xlsx"
    archivo.touch()
    assert "extracto" in validar_seleccion("", str(archivo))
    assert "movimientos" in validar_seleccion(str(archivo), "")


def test_avisa_si_elige_el_mismo_archivo_dos_veces(tmp_path):
    archivo = tmp_path / "a.xlsx"
    archivo.touch()
    assert "mismo archivo" in validar_seleccion(str(archivo), str(archivo))


def test_avisa_si_un_archivo_no_existe(tmp_path):
    existe = tmp_path / "a.xlsx"
    existe.touch()
    assert "No se encontró" in validar_seleccion(str(existe), str(tmp_path / "b.xlsx"))


def test_seleccion_completa_no_tiene_problemas(tmp_path):
    extracto, movimientos = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    extracto.touch()
    movimientos.touch()
    assert validar_seleccion(str(extracto), str(movimientos)) is None


def test_mensajes_de_errores_conocidos():
    assert "primero el extracto" in mensaje_de_error(ArchivoInvalido("Sin encabezados."))
    assert "Ya existe" in mensaje_de_error(SalidaExistente("Ya existe."))
    assert "ciérralo" in mensaje_de_error(PermissionError("ruta secreta"))
    assert ".xlsx" in mensaje_de_error(zipfile.BadZipFile("x"))


def test_error_inesperado_no_muestra_su_texto_interno():
    mensaje = mensaje_de_error(ValueError("dato real 123"))
    assert "dato real" not in mensaje
    assert "ValueError" in mensaje


def test_resumen_usa_los_colores_de_la_copia():
    resumen = Resumen(1, 2, 3, 4, 5, 6, 7, Path("c"), Path("r"), Path("m"))
    colores = {texto: color for texto, _, color in lineas_de_resumen(resumen)}
    assert colores["Seguras"] == COLOR_SEGURA
    assert colores["Dudosas (revisar)"] == COLOR_DUDOSA
    assert COLOR_REPETIDA in colores.values()
    cantidades = {texto: cantidad for texto, cantidad, _ in lineas_de_resumen(resumen)}
    assert cantidades["Seguras"] == 2
    assert cantidades["Ya estaban conciliadas (no se tocaron)"] == 1
