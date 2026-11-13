"""Comando principal: concilia el extracto con los movimientos.

Uso:
    python -m conciliacion EXTRACTO MOVIMIENTOS CARPETA_SALIDA

Crea tres archivos en la carpeta de salida:

- `<extracto>_conciliado.xlsx`: copia del extracto con los comprobantes
  escritos. Amarillo: seguro. Azul: el comprobante está en otra fila con
  otro tercero. Naranja: dudoso, hay que revisarlo.
- `<movimientos>_conciliado.xlsx`: copia de movimientos; cada movimiento
  usado se pinta del mismo color que su fila del extracto.
- `<extracto>_revision.xlsx`: azules y naranjas con sus candidatos, y pendientes.

Los archivos originales nunca se modifican. Si alguno de los archivos de
salida ya existe, no se sobrescribe y el comando se detiene.

En pantalla solo se muestran conteos, sin datos reales.
"""

import argparse
import hashlib
import sys
from dataclasses import dataclass
from pathlib import Path

from conciliacion.cruce import Estado, cruzar
from conciliacion.escritura import (
    COLOR_DUDOSA,
    COLOR_REPETIDA,
    COLOR_SEGURA,
    Marca,
    escribir_copia,
)
from conciliacion.lectura import (
    COLUMNAS_EXTRACTO,
    ArchivoInvalido,
    leer_extracto,
    leer_movimientos,
)
from conciliacion.revision import escribir_revision

ENCABEZADO_SALDOS = "SALDOS"
SUFIJO_COPIA = "_conciliado.xlsx"
SUFIJO_REVISION = "_revision.xlsx"
SUFIJO_TEMPORAL = ".parcial"


class SalidaExistente(Exception):
    """Un archivo de salida ya existe y no se debe sobrescribir."""


@dataclass
class Resumen:
    """Conteos de una ejecución."""

    ya_conciliadas: int
    seguras: int
    repetidas: int
    dudosas: int
    pendientes: int
    sin_pareja: int
    no_leidas: int
    copia: Path
    revision: Path
    copia_movimientos: Path


def huella(ruta):
    """Calcula la suma de comprobación de un archivo."""
    calculo = hashlib.sha256()
    with Path(ruta).open("rb") as contenido:
        for bloque in iter(lambda: contenido.read(1024 * 1024), b""):
            calculo.update(bloque)
    return calculo.hexdigest()


def nombre_base(ruta):
    """Devuelve el nombre del archivo sin extensión y sin `_conciliado` al final.

    Ella concilia a diario el archivo que salió el día anterior. Sin esto, el
    nombre acumularía `_conciliado_conciliado...`. Se quita aunque esté
    repetido o con otras mayúsculas.
    """
    nombre = Path(ruta).stem
    marca = SUFIJO_COPIA.removesuffix(".xlsx")
    while nombre.lower().endswith(marca) and len(nombre) > len(marca):
        nombre = nombre[: -len(marca)]
    return nombre


def rutas_de_salida(ruta_extracto, ruta_movimientos, carpeta_salida, entradas):
    """Calcula las rutas de salida y comprueba que se puedan crear.

    Devuelve la copia del extracto, la revisión y la copia de movimientos.
    """
    carpeta = Path(carpeta_salida)
    nombre = nombre_base(ruta_extracto)
    rutas = (
        carpeta / f"{nombre}{SUFIJO_COPIA}",
        carpeta / f"{nombre}{SUFIJO_REVISION}",
        carpeta / f"{nombre_base(ruta_movimientos)}{SUFIJO_COPIA}",
    )
    if len(set(rutas)) < len(rutas):
        raise SalidaExistente("El extracto y los movimientos tienen el mismo nombre.")
    for ruta in rutas:
        if ruta.exists():
            raise SalidaExistente(
                "Ya existe uno de los archivos de salida. "
                "Bórralo o elige otra carpeta de salida."
            )
        if ruta.resolve() in entradas:
            raise SalidaExistente("La salida no puede ser uno de los archivos originales.")
    return rutas


def columnas_a_pintar(columnas):
    """Devuelve las columnas de la tabla del extracto, de `FECHA` a `SALDOS`."""
    numeros = [columnas[nombre] for nombre in COLUMNAS_EXTRACTO.values()]
    if ENCABEZADO_SALDOS in columnas:
        numeros.append(columnas[ENCABEZADO_SALDOS])
    return range(min(numeros), max(numeros) + 1)


COLORES = {
    Estado.SEGURA: COLOR_SEGURA,
    Estado.REPETIDA: COLOR_REPETIDA,
    Estado.DUDOSA: COLOR_DUDOSA,
}


def marcas_de(resultado):
    """Relaciona cada fila del extracto con el comprobante y el color que se escriben."""
    return {
        propuesta.fila.fila: Marca(
            propuesta.movimiento.comprobante, COLORES[propuesta.estado]
        )
        for propuesta in resultado.propuestas
        if propuesta.movimiento is not None
    }


def marcas_de_movimientos(resultado):
    """Relaciona cada movimiento usado con el color de su fila del extracto.

    En la copia de movimientos solo se pinta; no se escribe texto.
    """
    return {
        propuesta.movimiento.fila: Marca("", COLORES[propuesta.estado])
        for propuesta in resultado.propuestas
        if propuesta.movimiento is not None
    }


def columnas_de_movimientos(columnas):
    """Devuelve todas las columnas con encabezado de la hoja de movimientos."""
    return range(min(columnas.values()), max(columnas.values()) + 1)


def crear_sin_dejar_restos(destino, crear):
    """Crea un archivo con un nombre temporal y lo renombra al terminar.

    Si algo falla a mitad de camino, no queda un archivo incompleto que se
    pueda confundir con un resultado válido.
    """
    temporal = destino.with_name(destino.name + SUFIJO_TEMPORAL)
    try:
        resultado = crear(temporal)
        temporal.replace(destino)
        return resultado
    finally:
        temporal.unlink(missing_ok=True)


def conciliar(ruta_extracto, ruta_movimientos, carpeta_salida):
    """Ejecuta la lectura, el cruce y la escritura de ambos archivos."""
    ruta_extracto = Path(ruta_extracto)
    ruta_movimientos = Path(ruta_movimientos)
    carpeta_salida = Path(carpeta_salida)
    entradas = {ruta_extracto.resolve(), ruta_movimientos.resolve()}
    huellas = {ruta: huella(ruta) for ruta in entradas}

    carpeta_salida.mkdir(parents=True, exist_ok=True)
    copia, revision, copia_movimientos = rutas_de_salida(
        ruta_extracto, ruta_movimientos, carpeta_salida, entradas
    )

    extracto = leer_extracto(ruta_extracto)
    movimientos = leer_movimientos(ruta_movimientos)
    resultado = cruzar(extracto.movimientos, movimientos.movimientos)

    columna_documento = extracto.columnas[COLUMNAS_EXTRACTO["documento"]]
    crear_sin_dejar_restos(
        copia,
        lambda destino: escribir_copia(
            ruta_extracto,
            destino,
            extracto.posicion_hoja - 1,
            columna_documento,
            columnas_a_pintar(extracto.columnas),
            marcas_de(resultado),
        ),
    )
    crear_sin_dejar_restos(
        copia_movimientos,
        lambda destino: escribir_copia(
            ruta_movimientos,
            destino,
            movimientos.posicion_hoja - 1,
            None,
            columnas_de_movimientos(movimientos.columnas),
            marcas_de_movimientos(resultado),
        ),
    )
    conteos = crear_sin_dejar_restos(
        revision,
        lambda destino: escribir_revision(
            destino, resultado, extracto.advertencias, movimientos.advertencias
        ),
    )

    if any(huella(ruta) != valor for ruta, valor in huellas.items()):
        raise RuntimeError("Un archivo original cambió durante la ejecución.")

    estados = [propuesta.estado for propuesta in resultado.propuestas]
    return Resumen(
        ya_conciliadas=resultado.filas_ya_conciliadas,
        seguras=estados.count(Estado.SEGURA),
        repetidas=estados.count(Estado.REPETIDA),
        dudosas=estados.count(Estado.DUDOSA),
        pendientes=estados.count(Estado.PENDIENTE),
        sin_pareja=conteos["sin_pareja"],
        no_leidas=conteos["no_leidas"],
        copia=copia,
        revision=revision,
        copia_movimientos=copia_movimientos,
    )


def imprimir_resumen(resumen):
    """Imprime los conteos de la ejecución, sin datos reales."""
    print("Conciliación terminada. Los originales no cambiaron.")
    print(f"  Filas que ya estaban conciliadas: {resumen.ya_conciliadas}")
    print(f"  Seguras (amarillo): {resumen.seguras}")
    print(f"  Comprobante repetido con otro tercero (azul, revisar): {resumen.repetidas}")
    print(f"  Dudosas (naranja, revisar): {resumen.dudosas}")
    print(f"  Pendientes del extracto: {resumen.pendientes}")
    print(f"  Movimientos sin pareja: {resumen.sin_pareja}")
    print(f"  Filas no leídas: {resumen.no_leidas}")
    print(
        "Archivos creados en la carpeta de salida: copia del extracto y copia de "
        f"movimientos (terminadas en {SUFIJO_COPIA}) y revisión ({SUFIJO_REVISION})."
    )


def main():
    """Lee los argumentos, concilia e imprime el resumen."""
    analizador = argparse.ArgumentParser(
        prog="python -m conciliacion",
        description="Concilia el extracto con los movimientos contables.",
    )
    analizador.add_argument("extracto", type=Path, help="Archivo de extracto")
    analizador.add_argument("movimientos", type=Path, help="Archivo de movimientos")
    analizador.add_argument("salida", type=Path, help="Carpeta donde se crean los resultados")
    argumentos = analizador.parse_args()

    for ruta in (argumentos.extracto, argumentos.movimientos):
        if not ruta.is_file():
            print("No se encontró uno de los archivos. Revisa las rutas.")
            return 1
    imprimir_resumen(conciliar(argumentos.extracto, argumentos.movimientos, argumentos.salida))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (ArchivoInvalido, SalidaExistente) as error:
        print(f"No se pudo conciliar: {error}")
        sys.exit(1)
    except Exception as error:  # Sin el mensaje: podría incluir rutas o datos.
        print(f"Error inesperado: {type(error).__name__}. No se creó ningún archivo.")
        sys.exit(1)
