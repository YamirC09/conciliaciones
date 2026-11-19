"""Punto de entrada del ejecutable: abre la ventana de conciliación.

PyInstaller necesita un archivo que se pueda correr directamente. Este solo
llama a la ventana; toda la lógica está en el paquete `conciliacion`.
"""

from conciliacion.interfaz import main

if __name__ == "__main__":
    main()
