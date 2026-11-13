"""Lee el extracto y los movimientos, y los deja en un mismo formato.

Este módulo solo lee: nunca escribe en los archivos. Las columnas se ubican por
su encabezado, porque los nombres de archivos y hojas cambian de un mes a otro.

Una fila que parece un movimiento pero no se puede interpretar no se descarta
en silencio: queda en `advertencias` con su número de fila, para que se revise.

Los mensajes nunca incluyen valores de celdas, porque pueden ser datos reales.

Para ver un resumen con conteos:
    python -m conciliacion.lectura EXTRACTO MOVIMIENTOS
"""

import argparse
import re
import sys
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook

COLUMNAS_EXTRACTO = {
    "fecha": "FECHA",
    "concepto": "CONCEPTO",
    "detalle": "DETALLE",
    "documento": "DOC CONCILIADO",
    "debito": "DB",
    "credito": "CR",
}
COLUMNAS_MOVIMIENTOS = {
    "comprobante": "COMPROBANTE",
    "secuencia": "SECUENCIA",
    "fecha": "FECHA ELABORACION",
    "identificacion": "IDENTIFICACION",
    "tercero": "NOMBRE TERCERO",
    "descripcion": "DESCRIPCION",
    "detalle": "DETALLE",
    "debito": "DEBITO",
    "credito": "CREDITO",
}
# Solo se muestran en el archivo de revisión; si faltan, quedan vacías.
OPCIONALES_MOVIMIENTOS = frozenset({"secuencia", "detalle"})
CONCEPTO_SALDO_INICIAL = "SALDO"
FILAS_MAXIMAS_ENCABEZADO = 20
PATRON_FECHA_MOVIMIENTOS = re.compile(r"(\d{2})-(\d{2})-(\d{4})")
PATRON_FECHA_EXTRACTO = re.compile(r"(\d{4})(\d{2})(\d{2})")


class ArchivoInvalido(Exception):
    """El archivo no tiene la estructura esperada."""


@dataclass(frozen=True)
class MovimientoExtracto:
    """Una fila de movimiento del extracto bancario."""

    fila: int
    fecha: date
    concepto: str
    detalle: str
    documento: str
    debito: int
    credito: int

    @property
    def conciliado(self):
        """Indica si ella ya le asignó un comprobante."""
        return bool(self.documento)


@dataclass(frozen=True)
class MovimientoContable:
    """Una fila de movimiento del sistema contable."""

    fila: int
    comprobante: str
    secuencia: str
    fecha: date
    identificacion: str
    tercero: str
    descripcion: str
    detalle: str
    debito: int
    credito: int


@dataclass
class Lectura:
    """Resultado de leer una hoja.

    `columnas` relaciona cada encabezado normalizado con su número de columna,
    contando desde 1. La escritura lo usa para saber dónde escribir y pintar.
    """

    posicion_hoja: int
    fila_encabezado: int = 0
    columnas: dict = field(default_factory=dict)
    movimientos: list = field(default_factory=list)
    ignoradas: Counter = field(default_factory=Counter)
    advertencias: list = field(default_factory=list)

    def advertir(self, fila, motivo):
        """Registra una fila que hay que revisar."""
        self.advertencias.append((fila, motivo))


def normalizar_encabezado(valor):
    """Quita tildes y espacios sobrantes, y pasa el encabezado a mayúsculas."""
    if valor is None:
        return ""
    sin_tildes = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(c for c in sin_tildes if not unicodedata.combining(c))
    return " ".join(texto.split()).upper()


def a_texto(valor):
    """Convierte una celda en texto sin espacios sobrantes.

    Excel puede guardar un código como número: 123.0 debe quedar como "123".
    """
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def a_centavos(valor):
    """Convierte un monto en centavos enteros. Una celda vacía vale 0.

    Los decimales de `float` son inexactos: 0.1 + 0.2 no es exactamente 0.3.
    Con enteros, dos montos iguales siempre se comparan como iguales.
    """
    if valor is None or valor == "":
        return 0
    if isinstance(valor, bool):
        raise ValueError("Monto no reconocido")
    try:
        monto = Decimal(str(valor).strip())
    except InvalidOperation as error:
        raise ValueError("Monto no reconocido") from error
    if not monto.is_finite():
        raise ValueError("Monto no reconocido")
    return int((monto * 100).quantize(Decimal(1), rounding=ROUND_HALF_UP))


def fecha_de_extracto(valor):
    """Interpreta una fecha del extracto, que viene como entero `AAAAMMDD`."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    coincidencia = PATRON_FECHA_EXTRACTO.fullmatch(a_texto(valor))
    if coincidencia is None:
        raise ValueError("Fecha no reconocida")
    anio, mes, dia = coincidencia.groups()
    return date(int(anio), int(mes), int(dia))


def fecha_de_movimientos(valor):
    """Interpreta una fecha de movimientos, que viene como texto `DD-MM-AAAA`."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    coincidencia = PATRON_FECHA_MOVIMIENTOS.fullmatch(a_texto(valor))
    if coincidencia is None:
        raise ValueError("Fecha no reconocida")
    dia, mes, anio = coincidencia.groups()
    return date(int(anio), int(mes), int(dia))


def ubicar_encabezado(hoja, columnas, opcionales=frozenset()):
    """Busca la fila con todos los encabezados requeridos.

    Devuelve el número de esa fila, la posición de cada columna y la de todos
    los encabezados de la fila, o `None` si la hoja no los tiene en sus
    primeras filas. Las posiciones cuentan desde 0. Una columna opcional que
    no está queda con posición `None`.
    """
    requeridos = {
        encabezado
        for nombre, encabezado in columnas.items()
        if nombre not in opcionales
    }
    filas = hoja.iter_rows(max_row=FILAS_MAXIMAS_ENCABEZADO, values_only=True)
    for numero, celdas in enumerate(filas, start=1):
        posiciones = {}
        for posicion, valor in enumerate(celdas):
            posiciones.setdefault(normalizar_encabezado(valor), posicion)
        if requeridos <= posiciones.keys():
            ubicacion = {
                nombre: posiciones.get(encabezado)
                for nombre, encabezado in columnas.items()
            }
            return numero, ubicacion, posiciones
    return None


def nueva_lectura(posicion_hoja, encontrado):
    """Crea la lectura de una hoja con los datos de su encabezado."""
    fila_encabezado, _, posiciones = encontrado
    return Lectura(
        posicion_hoja=posicion_hoja,
        fila_encabezado=fila_encabezado,
        columnas={
            nombre: posicion + 1 for nombre, posicion in posiciones.items() if nombre
        },
    )


def filas_de_datos(hoja, fila_encabezado, ubicacion):
    """Entrega cada fila posterior al encabezado como un diccionario por columna."""
    filas = hoja.iter_rows(min_row=fila_encabezado + 1, values_only=True)
    for numero, celdas in enumerate(filas, start=fila_encabezado + 1):
        yield numero, {
            nombre: (
                celdas[posicion]
                if posicion is not None and posicion < len(celdas)
                else None
            )
            for nombre, posicion in ubicacion.items()
        }


def interpretar_fila_extracto(numero, celdas, lectura):
    """Clasifica una fila del extracto y agrega el movimiento si lo es."""
    concepto = a_texto(celdas["concepto"])
    if normalizar_encabezado(concepto) == CONCEPTO_SALDO_INICIAL:
        lectura.ignoradas["saldo inicial"] += 1
        return
    try:
        debito = a_centavos(celdas["debito"])
        credito = a_centavos(celdas["credito"])
    except ValueError:
        lectura.advertir(numero, "monto no reconocido")
        return

    tiene_fecha = celdas["fecha"] not in (None, "")
    tiene_monto = bool(debito or credito)
    if not tiene_fecha and not tiene_monto:
        lectura.ignoradas["sin fecha ni monto"] += 1
        return
    if not tiene_fecha:
        lectura.advertir(numero, "monto sin fecha")
        return
    if not tiene_monto:
        lectura.advertir(numero, "fecha sin monto")
        return
    try:
        fecha = fecha_de_extracto(celdas["fecha"])
    except ValueError:
        lectura.advertir(numero, "fecha no reconocida")
        return
    if debito and credito:
        lectura.advertir(numero, "monto en DB y en CR")

    lectura.movimientos.append(
        MovimientoExtracto(
            fila=numero,
            fecha=fecha,
            concepto=concepto,
            detalle=a_texto(celdas["detalle"]),
            documento=a_texto(celdas["documento"]),
            debito=debito,
            credito=credito,
        )
    )


def interpretar_fila_movimientos(numero, celdas, lectura):
    """Clasifica una fila de movimientos y agrega el movimiento si lo es.

    Las filas sin comprobante son vacías, totales o pie de la exportación.
    """
    comprobante = a_texto(celdas["comprobante"])
    if not comprobante:
        lectura.ignoradas["sin comprobante"] += 1
        return
    try:
        debito = a_centavos(celdas["debito"])
        credito = a_centavos(celdas["credito"])
    except ValueError:
        lectura.advertir(numero, "monto no reconocido")
        return
    try:
        fecha = fecha_de_movimientos(celdas["fecha"])
    except ValueError:
        lectura.advertir(numero, "fecha no reconocida")
        return
    if not debito and not credito:
        lectura.advertir(numero, "sin monto")
    elif debito and credito:
        lectura.advertir(numero, "monto en débito y en crédito")

    lectura.movimientos.append(
        MovimientoContable(
            fila=numero,
            comprobante=comprobante,
            secuencia=a_texto(celdas["secuencia"]),
            fecha=fecha,
            identificacion=a_texto(celdas["identificacion"]),
            tercero=a_texto(celdas["tercero"]),
            descripcion=a_texto(celdas["descripcion"]),
            detalle=a_texto(celdas["detalle"]),
            debito=debito,
            credito=credito,
        )
    )


def abrir_libro(ruta):
    """Abre un libro en modo de solo lectura, sin cargarlo completo en memoria.

    `data_only` lee los valores calculados en lugar de las fórmulas.
    """
    return load_workbook(ruta, read_only=True, data_only=True)


def preparar_hoja(hoja):
    """Ignora la dimensión declarada, porque puede no coincidir con los datos."""
    hoja.reset_dimensions()
    return hoja


def leer_extracto(ruta):
    """Lee la hoja principal del extracto, ubicada por sus encabezados."""
    libro = abrir_libro(ruta)
    try:
        for posicion, hoja in enumerate(libro.worksheets, start=1):
            preparar_hoja(hoja)
            encontrado = ubicar_encabezado(hoja, COLUMNAS_EXTRACTO)
            if encontrado is None:
                continue
            lectura = nueva_lectura(posicion, encontrado)
            for numero, celdas in filas_de_datos(hoja, *encontrado[:2]):
                interpretar_fila_extracto(numero, celdas, lectura)
            return lectura
    finally:
        libro.close()
    raise ArchivoInvalido("Ninguna hoja tiene los encabezados del extracto.")


def leer_movimientos(ruta):
    """Lee los movimientos de la primera hoja; las demás no entran al proceso."""
    libro = abrir_libro(ruta)
    try:
        hoja = preparar_hoja(libro.worksheets[0])
        encontrado = ubicar_encabezado(
            hoja, COLUMNAS_MOVIMIENTOS, OPCIONALES_MOVIMIENTOS
        )
        if encontrado is None:
            raise ArchivoInvalido(
                "La primera hoja no tiene los encabezados de movimientos."
            )
        lectura = nueva_lectura(1, encontrado)
        for numero, celdas in filas_de_datos(hoja, *encontrado[:2]):
            interpretar_fila_movimientos(numero, celdas, lectura)
        return lectura
    finally:
        libro.close()


def imprimir_ignoradas_y_advertencias(lectura):
    """Imprime las filas ignoradas y las que hay que revisar."""
    print(f"  Filas ignoradas: {dict(lectura.ignoradas) or 'ninguna'}")
    motivos = Counter(motivo for _, motivo in lectura.advertencias)
    print(f"  Filas para revisar: {dict(motivos) or 'ninguna'}")
    if lectura.advertencias:
        filas = [fila for fila, _ in lectura.advertencias[:20]]
        print(f"  Números de fila (hasta 20): {filas}")


def imprimir_resumen(extracto, contable):
    """Imprime solo conteos, sin datos reales."""
    filas = extracto.movimientos
    conciliados = [fila for fila in filas if fila.conciliado]
    print("== Extracto")
    print(f"  Hoja principal: posición {extracto.posicion_hoja}")
    print(f"  Movimientos: {len(filas)}")
    print(f"  Con DB: {sum(1 for f in filas if f.debito)}; con CR: {sum(1 for f in filas if f.credito)}")
    print(f"  Conciliados: {len(conciliados)}; sin conciliar: {len(filas) - len(conciliados)}")
    if filas:
        print(f"  Fechas: del {min(f.fecha for f in filas)} al {max(f.fecha for f in filas)}")
    imprimir_ignoradas_y_advertencias(extracto)

    print("\n== Movimientos")
    print(f"  Hoja usada: posición {contable.posicion_hoja}")
    movimientos = contable.movimientos
    print(f"  Movimientos: {len(movimientos)}")
    print(f"  Con débito: {sum(1 for m in movimientos if m.debito)}; con crédito: {sum(1 for m in movimientos if m.credito)}")
    comprobantes = {m.comprobante for m in movimientos}
    claves = {(m.comprobante, m.secuencia) for m in movimientos}
    print(f"  Comprobantes distintos: {len(comprobantes)}; claves comprobante + secuencia: {len(claves)}")
    if movimientos:
        print(f"  Fechas: del {min(m.fecha for m in movimientos)} al {max(m.fecha for m in movimientos)}")
    imprimir_ignoradas_y_advertencias(contable)

    if conciliados:
        print("\n== Relación entre ambos")
        documentos = Counter(fila.documento for fila in conciliados)
        encontrados = sum(1 for fila in conciliados if fila.documento in comprobantes)
        print(f"  DOC CONCILIADO que existen en Comprobante: {encontrados} de {len(conciliados)}")
        print(f"  Comprobantes usados: una vez {sum(1 for n in documentos.values() if n == 1)}; más de una vez {sum(1 for n in documentos.values() if n > 1)}")


def main():
    """Lee los dos archivos e imprime el resumen."""
    analizador = argparse.ArgumentParser(
        description="Lee el extracto y los movimientos, y muestra un resumen con conteos."
    )
    analizador.add_argument("extracto", type=Path, help="Archivo de extracto")
    analizador.add_argument("movimientos", type=Path, help="Archivo de movimientos")
    argumentos = analizador.parse_args()

    extracto = leer_extracto(argumentos.extracto)
    contable = leer_movimientos(argumentos.movimientos)
    imprimir_resumen(extracto, contable)
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
