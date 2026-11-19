"""Ventana para conciliar sin usar la terminal.

Uso:
    python -m conciliacion.interfaz

Ella elige el extracto y los movimientos, y pulsa "Conciliar". Los resultados
se guardan en una carpeta nueva, junto al extracto, dentro de `Conciliaciones`,
con la fecha y la hora en el nombre. Así cada ejecución queda aparte y nunca se
sobrescribe un resultado anterior.

La lógica que no depende de la ventana (carpeta sugerida, mensajes de error,
resumen) está en funciones aparte para poder probarla sin abrir ventanas.
"""

import os
import queue
import subprocess
import sys
import threading
import zipfile
from datetime import datetime
from pathlib import Path

from conciliacion import __version__
from conciliacion.__main__ import SalidaExistente, conciliar
from conciliacion.escritura import COLOR_DUDOSA, COLOR_REPETIDA, COLOR_SEGURA
from conciliacion.lectura import ArchivoInvalido

CARPETA_RESULTADOS = "Conciliaciones"
FORMATO_CARPETA = "%Y-%m-%d %H-%M-%S"
TIPOS_EXCEL = [("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
ESPERA_MS = 200


def carpeta_de_salida(ruta_extracto, ahora=None):
    """Sugiere una carpeta nueva para los resultados, junto al extracto."""
    ahora = ahora or datetime.now()
    base = Path(ruta_extracto).parent / CARPETA_RESULTADOS
    return base / ahora.strftime(FORMATO_CARPETA)


def validar_seleccion(ruta_extracto, ruta_movimientos):
    """Devuelve un mensaje si falta algo por elegir, o `None` si está completo."""
    if not ruta_extracto:
        return "Elige el archivo del extracto."
    if not ruta_movimientos:
        return "Elige el archivo de movimientos."
    if Path(ruta_extracto).resolve() == Path(ruta_movimientos).resolve():
        return "Elegiste el mismo archivo dos veces. Elige el extracto y los movimientos."
    for ruta in (ruta_extracto, ruta_movimientos):
        if not Path(ruta).is_file():
            return f"No se encontró el archivo {Path(ruta).name}."
    return None


def mensaje_de_error(error):
    """Traduce un error a un mensaje que ella pueda entender y resolver.

    No incluye el texto interno de errores inesperados, porque podría tener
    rutas o datos.
    """
    if isinstance(error, ArchivoInvalido):
        return (
            f"{error}\n\nRevisa que elegiste cada archivo en su lugar: "
            "primero el extracto y después los movimientos."
        )
    if isinstance(error, SalidaExistente):
        return f"{error}"
    if isinstance(error, PermissionError):
        return (
            "No se pudo abrir uno de los archivos. Si está abierto en Excel, "
            "ciérralo y vuelve a intentar."
        )
    if isinstance(error, zipfile.BadZipFile):
        return "Uno de los archivos no es un Excel válido (.xlsx). Ábrelo y guárdalo como .xlsx."
    return (
        f"Ocurrió un error inesperado ({type(error).__name__}). "
        "No se modificó ningún archivo original."
    )


def lineas_de_resumen(resumen):
    """Arma las líneas del resumen: texto, cantidad y color de fondo (o `None`)."""
    return [
        ("Seguras", resumen.seguras, COLOR_SEGURA),
        ("Comprobante repetido con otro tercero (revisar)", resumen.repetidas, COLOR_REPETIDA),
        ("Dudosas (revisar)", resumen.dudosas, COLOR_DUDOSA),
        ("Pendientes del extracto", resumen.pendientes, None),
        ("Movimientos sin pareja", resumen.sin_pareja, None),
        ("Filas que no se pudieron leer", resumen.no_leidas, None),
        ("Ya estaban conciliadas (no se tocaron)", resumen.ya_conciliadas, None),
    ]


def abrir_carpeta(carpeta):
    """Abre la carpeta en el explorador de archivos del sistema."""
    if sys.platform.startswith("win"):
        os.startfile(carpeta)  # noqa: S606 - solo existe en Windows.
    elif sys.platform == "darwin":
        subprocess.run(["open", str(carpeta)], check=False)
    else:
        subprocess.run(["xdg-open", str(carpeta)], check=False)


class Ventana:
    """Ventana principal: elegir archivos, conciliar y ver el resumen."""

    def __init__(self, raiz):
        import tkinter as tk
        from tkinter import ttk

        self.tk = tk
        self.raiz = raiz
        self.resultados = queue.Queue()
        self.carpeta_resultado = None

        raiz.title(f"Conciliación bancaria {__version__}")
        raiz.minsize(560, 0)
        marco = ttk.Frame(raiz, padding=16)
        marco.grid(sticky="nsew")
        raiz.columnconfigure(0, weight=1)
        marco.columnconfigure(1, weight=1)

        self.extracto = tk.StringVar()
        self.movimientos = tk.StringVar()
        filas = [
            ("1. Extracto del banco", self.extracto),
            ("2. Movimientos del sistema", self.movimientos),
        ]
        for numero, (texto, variable) in enumerate(filas):
            ttk.Label(marco, text=texto).grid(row=numero, column=0, sticky="w", pady=4)
            ttk.Entry(marco, textvariable=variable, state="readonly").grid(
                row=numero, column=1, sticky="ew", padx=8
            )
            ttk.Button(
                marco,
                text="Elegir…",
                command=lambda v=variable, t=texto: self.elegir_archivo(v, t),
            ).grid(row=numero, column=2)

        self.boton_conciliar = ttk.Button(marco, text="Conciliar", command=self.conciliar)
        self.boton_conciliar.grid(row=2, column=0, columnspan=3, pady=(12, 4), ipady=6, sticky="ew")

        self.estado = ttk.Label(marco, text="Elige los dos archivos y pulsa Conciliar.")
        self.estado.grid(row=3, column=0, columnspan=3, sticky="w", pady=4)

        self.resumen = ttk.Frame(marco)
        self.resumen.grid(row=4, column=0, columnspan=3, sticky="ew")
        self.resumen.columnconfigure(0, weight=1)

        self.boton_abrir = ttk.Button(
            marco, text="Abrir carpeta de resultados", command=self.abrir_resultados
        )

    def elegir_archivo(self, variable, titulo):
        """Muestra el diálogo para elegir un archivo de Excel."""
        from tkinter import filedialog

        inicial = Path(self.extracto.get() or self.movimientos.get() or Path.home())
        ruta = filedialog.askopenfilename(
            title=titulo,
            filetypes=TIPOS_EXCEL,
            initialdir=inicial.parent if inicial.is_file() else inicial,
        )
        if ruta:
            variable.set(ruta)

    def conciliar(self):
        """Valida la selección y concilia en segundo plano."""
        from tkinter import messagebox

        problema = validar_seleccion(self.extracto.get(), self.movimientos.get())
        if problema:
            messagebox.showwarning("Falta algo", problema)
            return

        for elemento in self.resumen.winfo_children():
            elemento.destroy()
        self.boton_abrir.grid_remove()
        self.boton_conciliar.state(["disabled"])
        self.estado.config(text="Conciliando… puede tardar unos segundos.")

        extracto = Path(self.extracto.get())
        movimientos = Path(self.movimientos.get())
        carpeta = carpeta_de_salida(extracto)
        threading.Thread(
            target=self.trabajar, args=(extracto, movimientos, carpeta), daemon=True
        ).start()
        self.raiz.after(ESPERA_MS, self.revisar_resultado)

    def trabajar(self, extracto, movimientos, carpeta):
        """Corre la conciliación fuera de la ventana para que no se congele."""
        try:
            self.resultados.put(("ok", conciliar(extracto, movimientos, carpeta)))
        except Exception as error:  # Se muestra como mensaje, sin detalles internos.
            self.resultados.put(("error", error))

    def revisar_resultado(self):
        """Espera el resultado sin bloquear la ventana."""
        try:
            tipo, valor = self.resultados.get_nowait()
        except queue.Empty:
            self.raiz.after(ESPERA_MS, self.revisar_resultado)
            return
        self.boton_conciliar.state(["!disabled"])
        if tipo == "ok":
            self.mostrar_resumen(valor)
        else:
            from tkinter import messagebox

            self.estado.config(text="No se pudo conciliar.")
            messagebox.showerror("No se pudo conciliar", mensaje_de_error(valor))

    def mostrar_resumen(self, resumen):
        """Muestra los conteos con el color de cada tipo de fila."""
        from tkinter import ttk

        self.estado.config(text="Listo. Los archivos originales no cambiaron.")
        for numero, (texto, cantidad, color) in enumerate(lineas_de_resumen(resumen)):
            fondo = f"#{color}" if color else self.raiz.cget("background")
            self.tk.Label(
                self.resumen, text=f"  {texto}", anchor="w", background=fondo
            ).grid(row=numero, column=0, sticky="ew", pady=1)
            self.tk.Label(
                self.resumen, text=f"{cantidad}  ", anchor="e", background=fondo
            ).grid(row=numero, column=1, sticky="ew", pady=1)
        self.carpeta_resultado = resumen.copia.parent
        ttk.Label(
            self.resumen,
            text="Revisa en la copia las filas azules y naranjas; el detalle está "
            "en el archivo terminado en _revision.",
            wraplength=520,
        ).grid(row=99, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.boton_abrir.grid(row=5, column=0, columnspan=3, pady=(8, 0), sticky="ew")

    def abrir_resultados(self):
        """Abre la carpeta donde quedaron los resultados."""
        if self.carpeta_resultado:
            abrir_carpeta(self.carpeta_resultado)


def main():
    """Abre la ventana."""
    import tkinter as tk

    raiz = tk.Tk()
    Ventana(raiz)
    raiz.mainloop()


if __name__ == "__main__":
    main()
