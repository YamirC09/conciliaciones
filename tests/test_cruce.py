"""Pruebas de las reglas del cruce, con movimientos inventados.

Enfoque: cada regla se prueba en su borde, con un caso que apenas la cumple y
otro que apenas la incumple. Las pruebas de "no inventar" y de "no marcar como
seguro" son las más importantes, porque protegen del error más costoso.
"""

from datetime import date, timedelta

import pytest

from conciliacion.cruce import Estado, Motivo, cruzar, palabras
from conciliacion.lectura import MovimientoContable, MovimientoExtracto

FECHA = date(2026, 10, 15)


def fila(numero=2, credito=100000, debito=0, fecha=FECHA, documento="", concepto="PAGO", detalle=""):
    """Crea una fila de extracto inventada. Los montos van en centavos."""
    return MovimientoExtracto(numero, fecha, concepto, detalle, documento, debito, credito)


def movimiento(numero=2, credito=100000, debito=0, fecha=FECHA, comprobante=None, tercero="", secuencia="1"):
    """Crea un movimiento contable inventado. Los montos van en centavos."""
    return MovimientoContable(
        fila=numero,
        comprobante=comprobante or f"CC-1-{numero}",
        secuencia=secuencia,
        fecha=fecha,
        identificacion="",
        tercero=tercero,
        descripcion="",
        detalle="",
        debito=debito,
        credito=credito,
    )


def unica(filas, movimientos):
    """Cruza y devuelve la única propuesta."""
    propuestas = cruzar(filas, movimientos).propuestas
    assert len(propuestas) == 1
    return propuestas[0]


def test_coincidencia_unica_y_exacta_es_segura():
    propuesta = unica([fila()], [movimiento()])
    assert propuesta.estado is Estado.SEGURA
    assert propuesta.movimiento.comprobante == "CC-1-2"
    assert propuesta.motivos == ()


def test_debito_cruza_con_debito():
    propuesta = unica([fila(debito=100000, credito=0)], [movimiento(debito=100000, credito=0)])
    assert propuesta.estado is Estado.SEGURA


@pytest.mark.parametrize(
    ("datos_fila", "datos_movimiento"),
    [
        ({"debito": 100000, "credito": 0}, {"debito": 0, "credito": 100000}),
        ({"debito": 0, "credito": 100000}, {"debito": 100000, "credito": 0}),
    ],
)
def test_lado_contrario_no_es_candidato(datos_fila, datos_movimiento):
    propuesta = unica([fila(**datos_fila)], [movimiento(**datos_movimiento)])
    assert propuesta.estado is Estado.PENDIENTE
    assert propuesta.movimiento is None


@pytest.mark.parametrize("monto_movimiento", [200000, 50000])
def test_diferencia_de_hasta_mil_pesos_es_segura_si_es_el_mas_cercano(monto_movimiento):
    # La fila es de $1.000: $2.000 difiere justo $1.000; $500 difiere $500.
    propuesta = unica([fila(credito=100000)], [movimiento(credito=monto_movimiento)])
    assert propuesta.movimiento is not None
    assert propuesta.estado is Estado.SEGURA
    assert propuesta.motivos == ()


@pytest.mark.parametrize("monto_movimiento", [200001, 99_999_999])
def test_diferencia_mayor_a_mil_pesos_queda_pendiente(monto_movimiento):
    propuesta = unica([fila(credito=100000)], [movimiento(credito=monto_movimiento)])
    assert propuesta.estado is Estado.PENDIENTE
    assert propuesta.movimiento is None
    assert propuesta.motivos == (Motivo.SIN_CANDIDATOS,)


def test_entre_varios_con_diferencia_el_mas_cercano_es_seguro():
    propuesta = unica(
        [fila(credito=100000)],
        [movimiento(2, credito=100800, tercero="UNO"), movimiento(3, credito=100300, tercero="DOS")],
    )
    assert propuesta.movimiento.fila == 3
    assert propuesta.estado is Estado.SEGURA


def test_lado_contrario_no_entra_aunque_el_monto_este_en_tolerancia():
    propuesta = unica([fila(credito=100000, debito=0)], [movimiento(debito=100500, credito=0)])
    assert propuesta.estado is Estado.PENDIENTE


def test_prefiere_monto_exacto_sobre_uno_con_diferencia():
    propuesta = unica([fila(credito=100000)], [movimiento(2, credito=100500), movimiento(3, credito=100000)])
    assert propuesta.movimiento.fila == 3


def test_fila_con_diferencia_no_le_quita_el_movimiento_a_una_exacta():
    # La fila 2 va primero en el extracto, pero la 3 coincide exacto.
    filas = [fila(2, credito=100500), fila(3, credito=100000)]
    resultado = cruzar(filas, [movimiento(9, credito=100000)])
    por_fila = {p.fila.fila: p for p in resultado.propuestas}
    assert por_fila[3].movimiento.fila == 9
    assert por_fila[2].estado is Estado.PENDIENTE
    assert por_fila[2].motivos == (Motivo.CANDIDATOS_YA_USADOS,)


@pytest.mark.parametrize("dias", [-1, 1])
def test_un_dia_de_diferencia_queda_pendiente(dias):
    propuesta = unica([fila()], [movimiento(fecha=FECHA + timedelta(days=dias))])
    assert propuesta.estado is Estado.PENDIENTE
    assert propuesta.movimiento is None


def test_no_inventa_pareja_sin_movimientos():
    propuesta = unica([fila()], [])
    assert propuesta.estado is Estado.PENDIENTE
    assert propuesta.movimiento is None


def test_empate_de_monto_sin_nombre_que_distinga_es_dudosa():
    propuesta = unica(
        [fila(detalle="REF 123")],
        [movimiento(2, tercero="PROVEEDOR UNO"), movimiento(3, tercero="PROVEEDOR DOS")],
    )
    assert propuesta.estado is Estado.DUDOSA
    assert propuesta.motivos == (Motivo.NOMBRE_NO_DISTINGUE,)
    assert len(propuesta.candidatos) == 2


def test_empate_de_monto_con_el_mismo_tercero_es_segura():
    # Dos pagos iguales al mismo tercero el mismo día: da igual cuál se use.
    propuesta = unica(
        [fila(detalle="REF 123")],
        [movimiento(2, tercero="CLIENTE EJEMPLO"), movimiento(3, tercero="CLIENTE EJEMPLO")],
    )
    assert propuesta.estado is Estado.SEGURA


def test_movimiento_de_otra_fecha_no_cuenta_como_candidato():
    propuesta = unica(
        [fila()],
        [movimiento(2, fecha=FECHA + timedelta(days=1)), movimiento(3, fecha=FECHA)],
    )
    assert propuesta.movimiento.fila == 3
    assert propuesta.estado is Estado.SEGURA
    assert len(propuesta.candidatos) == 1


def test_empate_de_monto_se_resuelve_por_nombre_y_queda_amarillo():
    propuesta = unica(
        [fila(detalle="Cliente Ejemplo")],
        [movimiento(2, tercero="OTRO NOMBRE"), movimiento(3, tercero="CLIENTE EJEMPLO")],
    )
    assert propuesta.movimiento.fila == 3
    assert propuesta.estado is Estado.SEGURA
    assert propuesta.motivos == ()


@pytest.mark.parametrize(
    "detalle_banco",
    ["PAGO INVERS ALFAOMEGA", "PAGO INVERSIONES ALFAOMEG", "INVERCIONES ALFA OMEGA"],
)
def test_nombre_se_compara_de_forma_aproximada(detalle_banco):
    # Abreviado, cortado o con una letra distinta: igual debe reconocerlo.
    propuesta = unica(
        [fila(detalle=detalle_banco)],
        [movimiento(2, tercero="CLIENTE EJEMPLO"), movimiento(3, tercero="INVERSIONES ALFAOMEGA")],
    )
    assert propuesta.movimiento.fila == 3
    assert propuesta.estado is Estado.SEGURA


def test_dos_filas_iguales_en_monto_se_reparten_por_nombre_sin_importar_el_orden():
    # La fila 2 aparece primero, pero su nombre corresponde al movimiento 9.
    filas = [fila(2, detalle="CLIENTE BETA"), fila(3, detalle="CLIENTE ALFA")]
    movimientos = [movimiento(8, tercero="CLIENTE ALFA"), movimiento(9, tercero="CLIENTE BETA")]
    por_fila = {p.fila.fila: p for p in cruzar(filas, movimientos).propuestas}
    assert por_fila[2].movimiento.fila == 9
    assert por_fila[3].movimiento.fila == 8
    assert {p.estado for p in por_fila.values()} == {Estado.SEGURA}


def test_movimiento_se_usa_una_sola_vez_y_compartido_sin_nombre_es_dudosa():
    resultado = cruzar([fila(2, detalle="REF 111"), fila(3, detalle="REF 222")], [movimiento(2)])
    estados = sorted(p.estado.value for p in resultado.propuestas)
    assert estados == ["dudosa", "pendiente"]
    usados = [p.movimiento for p in resultado.propuestas if p.movimiento]
    assert len(usados) == 1


def test_movimiento_compartido_se_da_a_la_fila_cuyo_nombre_coincide():
    filas = [fila(2, detalle="CLIENTE BETA"), fila(3, detalle="CLIENTE ALFA")]
    resultado = cruzar(filas, [movimiento(8, tercero="CLIENTE ALFA")])
    por_fila = {p.fila.fila: p for p in resultado.propuestas}
    assert por_fila[3].estado is Estado.SEGURA
    assert por_fila[2].estado is Estado.PENDIENTE


def lineas_de_un_comprobante(*terceros):
    """Crea un comprobante con una línea por tercero, cada una de distinto monto."""
    return [
        movimiento(10 + i, credito=100000 * (i + 1), comprobante="CC-1-50", tercero=t, secuencia=str(i + 1))
        for i, t in enumerate(terceros)
    ]


def test_comprobante_en_dos_filas_con_terceros_distintos_queda_azul():
    filas = [fila(2, credito=100000), fila(3, credito=200000)]
    resultado = cruzar(filas, lineas_de_un_comprobante("PERSONA UNO", "PERSONA DOS"))
    por_fila = {p.fila.fila: p for p in resultado.propuestas}
    assert {p.estado for p in por_fila.values()} == {Estado.REPETIDA}
    assert por_fila[2].motivos == (Motivo.COMPROBANTE_REPETIDO,)
    assert por_fila[2].otras_filas == (3,)
    assert por_fila[3].otras_filas == (2,)


def test_comprobante_en_dos_filas_con_el_mismo_tercero_queda_amarillo():
    filas = [fila(2, credito=100000), fila(3, credito=200000)]
    resultado = cruzar(filas, lineas_de_un_comprobante("PERSONA UNO", "PERSONA UNO"))
    assert {p.estado for p in resultado.propuestas} == {Estado.SEGURA}


def test_comprobante_que_ella_ya_puso_en_otra_fila_tambien_cuenta():
    # Ella ya usó la primera línea en la fila 2; el script propone la segunda.
    filas = [fila(2, credito=100000, documento="CC-1-50"), fila(3, credito=200000)]
    resultado = cruzar(filas, lineas_de_un_comprobante("PERSONA UNO", "PERSONA DOS"))
    propuesta = resultado.propuestas[0]
    assert propuesta.fila.fila == 3
    assert propuesta.estado is Estado.REPETIDA
    assert propuesta.otras_filas == (2,)


def test_comprobante_repetido_en_fila_dudosa_sigue_naranja():
    # La fila 3 empata en monto con otra línea y el nombre no distingue.
    filas = [fila(2, credito=100000, detalle="REF 1"), fila(3, credito=200000, detalle="REF 2")]
    movimientos = lineas_de_un_comprobante("PERSONA UNO", "PERSONA DOS")
    movimientos.append(movimiento(20, credito=200000, tercero="PERSONA TRES"))
    por_fila = {p.fila.fila: p for p in cruzar(filas, movimientos).propuestas}
    assert por_fila[3].estado is Estado.DUDOSA


def test_filas_conciliadas_no_reciben_propuesta():
    resultado = cruzar([fila(documento="CC-9-9")], [movimiento(comprobante="CC-1-1")])
    assert resultado.propuestas == []
    assert resultado.filas_ya_conciliadas == 1


def test_comprobante_ya_usado_sale_de_los_candidatos():
    resultado = cruzar(
        [fila(2, documento="CC-1-1"), fila(3)],
        [movimiento(5, comprobante="CC-1-1")],
    )
    assert resultado.propuestas[0].estado is Estado.PENDIENTE
    assert resultado.movimientos_sin_pareja == []


def test_comprobante_repetido_solo_retira_la_fila_que_corresponde():
    conciliada = fila(2, credito=100000, documento="CC-1-1")
    pendiente = fila(3, credito=50000)
    movimientos = [
        movimiento(5, credito=100000, comprobante="CC-1-1", secuencia="1"),
        movimiento(6, credito=50000, comprobante="CC-1-1", secuencia="2"),
    ]
    propuesta = cruzar([conciliada, pendiente], movimientos).propuestas[0]
    assert propuesta.estado is Estado.SEGURA
    assert propuesta.movimiento.secuencia == "2"


def test_documento_inexistente_se_cuenta():
    resultado = cruzar([fila(documento="CC-9-9")], [])
    assert resultado.documentos_sin_movimiento == 1


def test_movimiento_sin_pareja_se_reporta():
    resultado = cruzar([fila(credito=100000)], [movimiento(2, credito=100000), movimiento(3, credito=999900)])
    assert [m.fila for m in resultado.movimientos_sin_pareja] == [3]


def test_fila_con_monto_en_ambos_lados_queda_pendiente():
    propuesta = unica([fila(debito=100, credito=100)], [movimiento(credito=100)])
    assert propuesta.estado is Estado.PENDIENTE
    assert propuesta.motivos == (Motivo.MONTO_EN_AMBOS_LADOS,)


def test_palabras_se_normalizan():
    assert palabras("Pagó a José, NIT 12", "de la") == {"PAGO", "JOSE", "NIT"}


def test_propuestas_conservan_el_orden_del_extracto():
    resultado = cruzar(
        [fila(2, credito=100), fila(3, credito=200)],
        [movimiento(9, credito=200), movimiento(8, credito=100)],
    )
    assert [p.fila.fila for p in resultado.propuestas] == [2, 3]
