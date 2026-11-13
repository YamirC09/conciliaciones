"""Pruebas de la evaluación, con casos inventados de respuesta conocida.

Enfoque: cada prueba arma un mes mínimo donde ya sabemos qué debe contar la
evaluación (un acierto, un intercambio, una pareja inventada...). Si la
evaluación cuenta mal, mediríamos mal el script y tomaríamos malas decisiones.
"""

from datetime import date, timedelta

from conciliacion.cruce import Estado
from conciliacion.evaluacion import (
    COINCIDE,
    INTERCAMBIO,
    INVENTADA_EN_VACIA,
    INVENTADA_SIN_MOVIMIENTO,
    OTRO,
    PENDIENTE_CORRECTA,
    evaluar,
    recortar_movimientos,
)
from conciliacion.lectura import MovimientoContable, MovimientoExtracto

FECHA = date(2026, 10, 15)


def fila(numero, documento="", credito=100000, fecha=FECHA):
    return MovimientoExtracto(numero, fecha, "PAGO", "", documento, 0, credito)


def movimiento(numero, comprobante, credito=100000, fecha=FECHA):
    return MovimientoContable(
        numero, comprobante, "1", fecha, "", "", "", "", 0, credito
    )


def test_acierto_seguro():
    evaluacion = evaluar([fila(2, "CC-1-1")], [movimiento(5, "CC-1-1")])
    assert evaluacion.total(COINCIDE, Estado.SEGURA) == 1
    assert evaluacion.filas_para_mirar == []


def test_fila_vacia_sin_propuesta_es_pendiente_correcta():
    evaluacion = evaluar([fila(2)], [])
    assert evaluacion.total(PENDIENTE_CORRECTA) == 1


def test_propuesta_en_fila_que_ella_dejo_vacia_es_inventada():
    evaluacion = evaluar([fila(2)], [movimiento(5, "CC-1-1")])
    assert evaluacion.total(INVENTADA_EN_VACIA) == 1
    assert evaluacion.filas_para_mirar[0][:2] == (2, INVENTADA_EN_VACIA)


def test_propuesta_cuando_su_movimiento_no_esta_es_inventada():
    # Ella usó CC-1-1, que no llegó; el script no debe tomar CC-1-2.
    evaluacion = evaluar([fila(2, "CC-1-1")], [movimiento(6, "CC-1-2", credito=100500)])
    assert evaluacion.total(INVENTADA_SIN_MOVIMIENTO) == 1


def test_pendiente_cuando_su_movimiento_no_esta_es_correcta():
    evaluacion = evaluar([fila(2, "CC-1-1")], [])
    assert evaluacion.total(PENDIENTE_CORRECTA) == 1


def test_intercambio_entre_dos_filas():
    # Ella cruzó los comprobantes; el script asigna cada uno a su monto exacto.
    filas = [fila(2, "CC-1-2", credito=100000), fila(3, "CC-1-1", credito=100100)]
    movimientos = [movimiento(5, "CC-1-1", credito=100000), movimiento(6, "CC-1-2", credito=100100)]
    evaluacion = evaluar(filas, movimientos)
    assert evaluacion.total(INTERCAMBIO) == 2


def test_movimiento_que_ella_no_uso():
    filas = [fila(2, "CC-1-1", credito=100000)]
    movimientos = [movimiento(5, "CC-1-1", credito=100900), movimiento(6, "CC-1-9", credito=100000)]
    evaluacion = evaluar(filas, movimientos)
    assert evaluacion.total(OTRO) == 1


def test_recorte_conserva_solo_los_primeros_dias():
    movimientos = [movimiento(n, f"CC-1-{n}", fecha=FECHA + timedelta(days=n)) for n in range(5)]
    assert [m.fila for m in recortar_movimientos(movimientos, 3)] == [0, 1, 2]
    assert recortar_movimientos([], 3) == []
