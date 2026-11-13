"""Busca la pareja de cada fila del extracto entre los movimientos contables.

Este módulo no lee ni escribe archivos: recibe movimientos ya normalizados y
devuelve una propuesta por cada fila sin conciliar.

Un movimiento es candidato solo si está en el mismo lado (`DB` con débito, `CR`
con crédito), tiene la misma fecha y su monto difiere hasta $1.000. Si ninguno
cumple, la fila queda pendiente: el script nunca inventa una pareja.

Entre los candidatos se elige el de monto más cercano. Si hay empate en el
monto, decide el nombre: se compara, de forma aproximada, el texto del banco con
el tercero del movimiento, porque el banco los escribe distinto (abreviados,
cortados o en otro orden).

Clasificación:

- Segura (amarillo): el elegido es el de monto más cercano. Vale aunque el
  monto tenga diferencia, y aunque haya otro igual de otro tercero si el
  nombre permitió elegir.
- Repetida (azul): el comprobante elegido también está en otra fila del
  extracto (puesto por ella o por el script) con otro tercero. Pasa cuando un
  comprobante tiene varias líneas, una por persona.
- Dudosa (naranja): había empate de monto y el nombre no permitió distinguir.
  Tiene prioridad sobre el azul.
- Pendiente: sin candidatos.
"""

import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field, replace
from difflib import SequenceMatcher
from enum import Enum

TOLERANCIA_CENTAVOS = 100_000  # Hasta $1.000 de diferencia en el monto.
LARGO_MINIMO_PALABRA = 3
LARGO_MINIMO_ABREVIATURA = 4  # "INVERS" cuenta como "INVERSIONES".
PARECIDO_MINIMO_PALABRA = 0.8  # Tolera una o dos letras distintas.


class Estado(Enum):
    """Clasificación de una fila del extracto."""

    SEGURA = "segura"
    REPETIDA = "comprobante repetido"
    DUDOSA = "dudosa"
    PENDIENTE = "pendiente"


class Motivo(Enum):
    """Explicación de por qué una fila no quedó como segura."""

    COMPROBANTE_REPETIDO = "el comprobante también está en otra fila con otro tercero"
    NOMBRE_NO_DISTINGUE = "mismo monto y fecha que otro de distinto tercero; el nombre no permite elegir"
    SIN_CANDIDATOS = "sin candidatos"
    CANDIDATOS_YA_USADOS = "candidatos ya usados"
    MONTO_EN_AMBOS_LADOS = "monto en DB y en CR"


@dataclass(frozen=True)
class Propuesta:
    """Resultado del cruce para una fila del extracto."""

    fila: object
    estado: Estado
    movimiento: object = None
    candidatos: tuple = ()
    motivos: tuple = ()
    otras_filas: tuple = ()  # Filas del extracto con el mismo comprobante.


@dataclass
class ResultadoCruce:
    """Propuestas por fila y movimientos que no encontraron pareja."""

    propuestas: list = field(default_factory=list)
    movimientos_sin_pareja: list = field(default_factory=list)
    filas_ya_conciliadas: int = 0
    documentos_sin_movimiento: int = 0


def palabras(*textos):
    """Devuelve las palabras normalizadas: mayúsculas, sin tildes, 3 o más caracteres."""
    unido = " ".join(texto or "" for texto in textos)
    sin_tildes = unicodedata.normalize("NFKD", unido)
    limpio = "".join(c for c in sin_tildes if not unicodedata.combining(c)).upper()
    return frozenset(
        palabra
        for palabra in re.findall(r"[A-Z0-9]+", limpio)
        if len(palabra) >= LARGO_MINIMO_PALABRA
    )


def parecido_entre_palabras(una, otra):
    """Mide de 0 a 1 cuánto se parecen dos palabras.

    Una abreviatura (una palabra que empieza como la otra) cuenta casi como
    igual; si no, se mide cuántas letras comparten en orden.
    """
    if una == otra:
        return 1.0
    cortas = min(len(una), len(otra)) >= LARGO_MINIMO_ABREVIATURA
    if cortas and (una.startswith(otra) or otra.startswith(una)):
        return 0.9
    return SequenceMatcher(None, una, otra).ratio()


def texto_del_banco(fila):
    """Palabras del extracto donde aparece el nombre."""
    return palabras(fila.concepto, fila.detalle)


def nombre_del_tercero(movimiento):
    """Palabras del movimiento que identifican al tercero."""
    return palabras(movimiento.tercero, movimiento.identificacion)


def similitud_de_nombre(fila, movimiento):
    """Mide de 0 a 1 qué parte del nombre del tercero aparece en el banco.

    Cada palabra del tercero se compara con la más parecida del banco, y solo
    cuenta si se parece lo suficiente. No se exige coincidencia exacta.
    """
    banco = texto_del_banco(fila)
    tercero = nombre_del_tercero(movimiento)
    if not banco or not tercero:
        return 0.0
    puntajes = [max(parecido_entre_palabras(t, b) for b in banco) for t in tercero]
    return sum(p for p in puntajes if p >= PARECIDO_MINIMO_PALABRA) / len(tercero)


def lado_y_monto(fila):
    """Devuelve el lado y el monto de una fila, o `None` si tiene ambos lados."""
    if fila.debito and fila.credito:
        return None
    return ("debito", fila.debito) if fila.debito else ("credito", fila.credito)


def diferencia_de_monto(fila, movimiento):
    """Devuelve la diferencia en centavos, o `None` si están en lados distintos."""
    datos = lado_y_monto(fila)
    if datos is None:
        return None
    lado, monto = datos
    monto_movimiento = getattr(movimiento, lado)
    if not monto_movimiento:
        return None
    return abs(monto - monto_movimiento)


def es_candidato(fila, movimiento):
    """Indica si coinciden el lado y la fecha, y el monto dentro de la tolerancia."""
    diferencia = diferencia_de_monto(fila, movimiento)
    return (
        diferencia is not None
        and diferencia <= TOLERANCIA_CENTAVOS
        and fila.fecha == movimiento.fecha
    )


def quitar_movimientos_ya_usados(conciliadas, movimientos):
    """Quita los movimientos que corresponden a filas que ella ya concilió.

    Un comprobante puede repetirse en varias filas de movimientos, y
    `DOC CONCILIADO` no indica la secuencia. Por eso cada fila conciliada retira
    una sola fila: la de su mismo lado con el monto más parecido.

    Devuelve los movimientos disponibles, cuántos documentos no se encontraron
    y qué movimiento se retiró por cada fila conciliada.
    """
    disponibles = list(movimientos)
    sin_movimiento = 0
    retirados = {}
    for fila in conciliadas:
        del_documento = [m for m in disponibles if m.comprobante == fila.documento]
        if not del_documento:
            sin_movimiento += 1
            continue

        def orden(movimiento, fila=fila):
            diferencia = diferencia_de_monto(fila, movimiento)
            return (diferencia is None, diferencia or 0, movimiento.fila)

        retirados[fila.fila] = min(del_documento, key=orden)
        disponibles.remove(retirados[fila.fila])
    return disponibles, sin_movimiento, retirados


def asignar(candidatos_por_fila, filas_por_numero):
    """Reparte los movimientos entre las filas, de la pareja más parecida a la menos.

    Se ordenan todas las parejas posibles por diferencia de monto y, si empatan,
    por parecido del nombre. Se recorren en ese orden y se acepta cada pareja
    cuya fila y cuyo movimiento sigan libres. Así una fila con diferencia no le
    quita el movimiento a una exacta, y en un empate de monto cada fila se
    queda con el tercero que más se le parece.
    """
    parejas = sorted(
        (
            diferencia_de_monto(filas_por_numero[numero], movimiento),
            -similitud_de_nombre(filas_por_numero[numero], movimiento),
            numero,
            movimiento.fila,
            movimiento,
        )
        for numero, candidatos in candidatos_por_fila.items()
        for movimiento in candidatos
    )
    asignados = {}
    usados = set()
    for _, _, numero, numero_movimiento, movimiento in parejas:
        if numero in asignados or numero_movimiento in usados:
            continue
        asignados[numero] = movimiento
        usados.add(numero_movimiento)
    return asignados


def clasificar(fila, elegido, candidatos, filas_que_lo_comparten):
    """Decide el color de una fila con propuesta.

    Busca rivales: otro movimiento de la fila con la misma diferencia de monto,
    u otra fila con la misma diferencia frente al elegido. Un rival del mismo
    tercero (o con el mismo texto del banco) no cuenta, porque da igual cuál se
    use. Sin rivales, o si el nombre del elegido se parece más que el de todos
    los rivales, es segura; si no, dudosa.
    """
    diferencia = diferencia_de_monto(fila, elegido)
    otros_movimientos = [
        m
        for m in candidatos
        if m.fila != elegido.fila
        and diferencia_de_monto(fila, m) == diferencia
        and nombre_del_tercero(m) != nombre_del_tercero(elegido)
    ]
    otras_filas = [
        f
        for f in filas_que_lo_comparten
        if f.fila != fila.fila
        and diferencia_de_monto(f, elegido) == diferencia
        and texto_del_banco(f) != texto_del_banco(fila)
    ]
    if not otros_movimientos and not otras_filas:
        return Estado.SEGURA, ()

    propia = similitud_de_nombre(fila, elegido)
    rivales = [similitud_de_nombre(fila, m) for m in otros_movimientos]
    rivales += [similitud_de_nombre(f, elegido) for f in otras_filas]
    if propia > max(rivales):
        return Estado.SEGURA, ()
    return Estado.DUDOSA, (Motivo.NOMBRE_NO_DISTINGUE,)


def marcar_comprobantes_repetidos(propuestas, retirados):
    """Pasa a azul las filas cuyo comprobante está en otra fila con otro tercero.

    Se cuentan las filas que ella ya concilió (con el movimiento que se les
    retiró) y las propuestas del script. Ella no se toca: solo cambian las
    propuestas seguras del script. Las dudosas siguen en naranja.
    """
    filas_por_comprobante = defaultdict(list)
    for numero, movimiento in retirados.items():
        filas_por_comprobante[movimiento.comprobante].append((numero, movimiento))
    for propuesta in propuestas:
        if propuesta.movimiento is not None:
            filas_por_comprobante[propuesta.movimiento.comprobante].append(
                (propuesta.fila.fila, propuesta.movimiento)
            )

    resultado = []
    for propuesta in propuestas:
        if propuesta.estado is Estado.SEGURA:
            propio = propuesta.movimiento
            grupo = filas_por_comprobante[propio.comprobante]
            otras = tuple(sorted(n for n, _ in grupo if n != propuesta.fila.fila))
            terceros = {nombre_del_tercero(m) for _, m in grupo}
            if otras and len(terceros) > 1:
                propuesta = replace(
                    propuesta,
                    estado=Estado.REPETIDA,
                    motivos=(Motivo.COMPROBANTE_REPETIDO,),
                    otras_filas=otras,
                )
        resultado.append(propuesta)
    return resultado


def cruzar(filas_extracto, movimientos):
    """Propone un comprobante para cada fila del extracto sin conciliar."""
    conciliadas = [fila for fila in filas_extracto if fila.conciliado]
    sin_conciliar = [fila for fila in filas_extracto if not fila.conciliado]
    disponibles, sin_movimiento, retirados = quitar_movimientos_ya_usados(
        conciliadas, movimientos
    )

    filas_por_numero = {fila.fila: fila for fila in sin_conciliar}
    candidatos_por_fila = {
        fila.fila: tuple(m for m in disponibles if es_candidato(fila, m))
        for fila in sin_conciliar
    }
    filas_por_movimiento = defaultdict(list)
    for numero, candidatos in candidatos_por_fila.items():
        for movimiento in candidatos:
            filas_por_movimiento[movimiento.fila].append(filas_por_numero[numero])

    asignados = asignar(candidatos_por_fila, filas_por_numero)

    propuestas = []
    for fila in sin_conciliar:
        candidatos = candidatos_por_fila[fila.fila]
        elegido = asignados.get(fila.fila)
        if lado_y_monto(fila) is None:
            motivos = (Motivo.MONTO_EN_AMBOS_LADOS,)
        elif not candidatos:
            motivos = (Motivo.SIN_CANDIDATOS,)
        elif elegido is None:
            motivos = (Motivo.CANDIDATOS_YA_USADOS,)
        else:
            estado, motivos = clasificar(
                fila, elegido, candidatos, filas_por_movimiento[elegido.fila]
            )
            propuestas.append(Propuesta(fila, estado, elegido, candidatos, motivos))
            continue
        propuestas.append(
            Propuesta(fila, Estado.PENDIENTE, candidatos=candidatos, motivos=motivos)
        )

    usados = {movimiento.fila for movimiento in asignados.values()}
    return ResultadoCruce(
        propuestas=marcar_comprobantes_repetidos(propuestas, retirados),
        movimientos_sin_pareja=[m for m in disponibles if m.fila not in usados],
        filas_ya_conciliadas=len(conciliadas),
        documentos_sin_movimiento=sin_movimiento,
    )
