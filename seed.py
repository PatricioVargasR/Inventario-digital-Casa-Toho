"""
    Carga de datos para el inventario desde un libro de Excel.

    Uso:
        python seed.py                       # carga datos/Resumen_de_productos.xlsx
        python seed.py otro-libro.xlsx       # carga otro libro (o un respaldo descargado de la app)

    Es seguro ejecutarlo varias veces: lo que ya existe se omite y nunca se borra nada.
"""

import sys
from pathlib import Path

from database import Base, LocalSession, engine
from excel import importar

RUTA_POR_DEFECTO = Path(__file__).with_name('datos') / 'Resumen_de_productos.xlsx'


def etiqueta(contador):
    return ', '.join(f'{n} {nombre}' for nombre, n in contador.items()) or 'nada'


if __name__ == '__main__':
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else RUTA_POR_DEFECTO
    if not ruta.exists():
        sys.exit(f'No encuentro el archivo: {ruta}')

    Base.metadata.create_all(bind=engine)
    with LocalSession() as db:
        resultado = importar(db, ruta)

    print(f'{ruta.name}')
    print(f'  Creados:  {etiqueta(resultado["creados"])}')
    print(f'  Omitidos (ya existían): {etiqueta(resultado["omitidos"])}')
    for aviso in resultado['avisos']:
        print(f'  Aviso: {aviso}')
