"""Funciones pequeñas de texto compartidas por la app, el importador y el exportador."""

import re
import unicodedata


def limpiar(texto):
    """Quita espacios al inicio/final y colapsa los espacios repetidos: ' Cebolla  blanca ' -> 'Cebolla blanca'."""
    return re.sub(r'\s+', ' ', str(texto if texto is not None else '')).strip()


def clave(texto):
    """Para comparar nombres: sin acentos, sin mayúsculas y sin espacios de más.
    'Panadería la perla ' y 'panaderia  La Perla' dan la misma clave."""
    sin_acentos = ''.join(
        c for c in unicodedata.normalize('NFD', limpiar(texto)) if not unicodedata.combining(c)
    )
    return sin_acentos.casefold()


def numero(n):
    """12.0 -> '12'; 2.5 -> '2.5'; None -> ''."""
    if n is None or n == '':
        return ''
    return str(int(n)) if float(n).is_integer() else str(n)


def moneda(n):
    """405 -> '$405.00'; None -> ''."""
    return '' if n is None else f'${n:,.2f}'
