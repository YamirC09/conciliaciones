"""Mide el cruce contra un mes que ella ya concilió.

Oculta en memoria lo que ella escribió en `DOC CONCILIADO`, deja que el script
proponga y compara fila por fila. No escribe en ningún archivo.

También simula exportaciones parciales: deja todo el extracto, pero solo los
movimientos de los primeros días, como pasa a mitad de mes. Ahí, una fila cuyo
movimiento todavía no llegó debe quedar pendiente; si recibe una propuesta, el
script inventó una pareja.

Solo imprime conteos y números de fila, nunca montos ni comprobantes.

Uso:
    python -m conciliacion.evaluacion EXTRACTO MOVIMIENTOS
    python -m conciliacion.evaluacion EXTRACTO MOVIMIENTOS --dias 3 7 --detalle
"""

import argparse
import sys
from collections import Counter
from dataclasses import dataclass, field, replace
from datetime import timedelta
from pathlib import Path

from conciliacion.cruce import Estado, cruzar
from conciliacion.lectura import ArchivoInvalido, leer_extracto, leer_movimientos

DIAS_PARCIALES = (3, 7, 15)

COINCIDE = "coincide con ella"
INTERCAMBIO = "intercambiada con otra fila"
OTRO = "movimiento que ella no usó"
SIN_PROPUESTA = "sin propuesta"
INVENTADA_EN_VACIA = "inventada en fila que ella dejó vacía"
INVENTADA_SIN_MOVIMIENTO = "inventada: su movimiento no estaba"
PENDIENTE_CORRECTA = "pendiente correcta"

INVENTADAS = (INVENTADA_EN_VACIA, INVENTADA_SIN_MOVIMIENTO)
ORDEN = (
    COINCIDE,
    INTERCAMBIO,
    OTRO,
    SIN_PROPUESTA,
    PENDIENTE_CORRECTA,
    INVENTADA_EN_VACIA,
    INVENTADA_SIN_MOVIMIENTO,
)


@dataclass
class Evaluacion:
    """Conteos por resultado y estado, y las filas que hay que mirar."""

    conteos: Counter = field(default_factory=Counter)
    filas_para_mirar: list = field(default_factory=list)

    def total(self, resultado, estado=None):
        """Suma las filas de un resultado, en un estado o en todos."""
        if estado is not None:
            return self.conteos[(resultado, estado)]
        return sum(n for (r, _), n in self.conteos.items() if r == resultado)


def clasificar(propuesta, documento, disponibles, usados_por_ella):
    """Compara una propuesta con lo que ella escribió en esa fila."""
    movimiento = propuesta.movimiento
    if not documento:
        return INVENTADA_EN_VACIA if movimiento else PENDIENTE_CORRECTA
    if documento not in disponibles:
        return INVENTADA_SIN_MOVIMIENTO if movimiento else PENDIENTE_CORRECTA
    if movimiento is None:
        return SIN_PROPUESTA
    if movimiento.comprobante == documento:
        return COINCIDE
    if movimiento.comprobante in usados_por_ella:
        return INTERCAMBIO
    return OTRO


def evaluar(filas_extracto, movimientos):
    """Cruza sin los comprobantes de ella y compara con lo que escribió."""
    referencia = {fila.fila: fila.documento for fila in filas_extracto}
    usados_por_ella = {documento for documento in referencia.values() if documento}
    disponibles = {movimiento.comprobante for movimiento in movimientos}
    ocultas = [replace(fila, documento="") for fila in filas_extracto]

    evaluacion = Evaluacion()
    for propuesta in cruzar(ocultas, movimientos).propuestas:
        numero = propuesta.fila.fila
        resultado = clasificar(
            propuesta, referencia[numero], disponibles, usados_por_ella
        )
        evaluacion.conteos[(resultado, propuesta.estado)] += 1
        es_segura_distinta = (
            propuesta.estado is Estado.SEGURA and resultado != COINCIDE
        )
        if resultado in INVENTADAS or resultado == OTRO or es_segura_distinta:
            evaluacion.filas_para_mirar.append(
                (
                    numero,
                    resultado,
                    propuesta.movimiento.fila if propuesta.movimiento else None,
                    propuesta.estado,
                )
            )
    return evaluacion


def recortar_movimientos(movimientos, dias):
    """Deja solo los movimientos de los primeros días del archivo."""
    if not movimientos:
        return []
    corte = min(m.fecha for m in movimientos) + timedelta(days=dias - 1)
    return [m for m in movimientos if m.fecha <= corte]


def imprimir_evaluacion(titulo, evaluacion, movimientos_usados, detalle):
    """Imprime la tabla de resultados por estado."""
    print(f"\n== {titulo} ({movimientos_usados} movimientos)")
    estados = (Estado.SEGURA, Estado.REPETIDA, Estado.DUDOSA, Estado.PENDIENTE)
    print(f"  {'Resultado':<40}" + "".join(f"{e.value:>22}" for e in estados))
    for resultado in ORDEN:
        valores = [evaluacion.total(resultado, estado) for estado in estados]
        if any(valores):
            print(f"  {resultado:<40}" + "".join(f"{v:>22}" for v in valores))

    inventadas = sum(evaluacion.total(r) for r in INVENTADAS)
    seguras_distintas = sum(
        evaluacion.total(r, Estado.SEGURA) for r in ORDEN if r != COINCIDE
    )
    print(f"  Inventadas: {inventadas}. Seguras que no coinciden con ella: {seguras_distintas}.")
    if detalle and evaluacion.filas_para_mirar:
        print("  Filas para mirar (extracto | resultado | movimiento propuesto | estado):")
        for numero, resultado, movimiento, estado in evaluacion.filas_para_mirar:
            print(f"    {numero} | {resultado} | {movimiento or '-'} | {estado.value}")


def main():
    """Evalúa el mes completo y las exportaciones parciales."""
    analizador = argparse.ArgumentParser(
        description="Mide el cruce contra un mes ya conciliado. Solo imprime conteos."
    )
    analizador.add_argument("extracto", type=Path, help="Extracto ya conciliado")
    analizador.add_argument("movimientos", type=Path, help="Movimientos del mismo mes")
    analizador.add_argument(
        "--dias",
        type=int,
        nargs="*",
        default=list(DIAS_PARCIALES),
        help="Días para simular exportaciones parciales",
    )
    analizador.add_argument(
        "--detalle",
        action="store_true",
        help="Muestra los números de fila que hay que mirar",
    )
    argumentos = analizador.parse_args()

    filas = leer_extracto(argumentos.extracto).movimientos
    movimientos = leer_movimientos(argumentos.movimientos).movimientos
    conciliadas = sum(1 for fila in filas if fila.conciliado)
    print(f"Filas del extracto: {len(filas)}; conciliadas por ella: {conciliadas}.")

    imprimir_evaluacion(
        "Mes completo", evaluar(filas, movimientos), len(movimientos), argumentos.detalle
    )
    for dias in argumentos.dias:
        parciales = recortar_movimientos(movimientos, dias)
        imprimir_evaluacion(
            f"Exportación parcial: primeros {dias} días",
            evaluar(filas, parciales),
            len(parciales),
            argumentos.detalle,
        )
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ArchivoInvalido as error:
        print(f"Archivo no válido: {error}")
        sys.exit(1)
    except Exception as error:  # Sin el mensaje: podría incluir rutas o datos.
        print(f"Error inesperado: {type(error).__name__}.")
        sys.exit(1)
