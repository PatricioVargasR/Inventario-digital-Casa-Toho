"""
    Lectura y escritura del libro de Excel del inventario.

    El libro tiene el mismo formato que Resumen_de_productos.xlsx:

        Proveedores                  catálogo de proveedores (una columna)
        Categorías                   catálogo de categorías (una columna)
        Cotejo Insumo                lista de productos por área, con su estado
        Insumos original (Resumen)   LA PLANTILLA que se quiere obtener

    importar()      -> carga ese libro en la base de datos (se puede repetir sin duplicar)
    libro_resumen() -> Excel con solo la hoja de resumen (el resultado del inventario)
    libro_respaldo()-> Excel con las 4 hojas, que importar() puede volver a leer
"""

import io
import re
from collections import Counter

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from sqlalchemy import select
from sqlalchemy.orm import joinedload, selectinload

import models
from utiles import clave, limpiar

# Unidades de medida de la plantilla original (la lista desplegable de la columna D)
UNIDADES_BASE = ['kg', 'gr', 'lt', 'ml', 'pza', 'gal']

HOJA_PROVEEDORES = 'Proveedores'
HOJA_CATEGORIAS = 'Categorías'
HOJA_COTEJO = 'Cotejo Insumo'
HOJA_RESUMEN = 'Insumos original (Resumen)'


# ==================================================================== IMPORTAR
def asegurar_catalogos(db):
    """Crea las unidades de medida que falten (no toca las existentes)."""
    existentes = {clave(u.nombre_unidad) for u in db.scalars(select(models.Unidades))}
    for nombre in UNIDADES_BASE:
        if clave(nombre) not in existentes:
            db.add(models.Unidades(nombre_unidad=nombre))
    db.commit()


def _hoja(libro, palabra):
    """La hoja cuyo nombre contiene la palabra (sin importar acentos ni mayúsculas)."""
    return next((ws for ws in libro.worksheets if palabra in clave(ws.title)), None)


def _fila_de_datos(ws, encabezado):
    """Número de la primera fila de datos: la que sigue al encabezado indicado en la columna A."""
    for fila in ws.iter_rows(min_row=1, max_row=15, max_col=1):
        if clave(fila[0].value) == encabezado:
            return fila[0].row + 1
    return None


def _precio(valor):
    """Número mayor que 0, o None. En la plantilla original 0.0 significa 'sin capturar'."""
    if valor is None or limpiar(valor) == '':
        return None
    try:
        numero = float(str(valor).replace('$', '').replace(',', '.').strip())
    except ValueError:
        return None
    return numero if numero > 0 else None


def _asegurar(db, modelo, campo, nombre, cache, creados, etiqueta):
    """Busca por nombre (sin acentos ni mayúsculas) y, si no existe, lo crea."""
    nombre = limpiar(nombre)
    if not nombre:
        return None
    llave = clave(nombre)
    if llave not in cache:
        cache[llave] = modelo(**{campo: nombre})
        db.add(cache[llave])
        creados[etiqueta] += 1
    return cache[llave]


def importar(db, origen):
    """
    Carga un libro de Excel en la base de datos. Es aditivo: lo que ya existe se omite y
    nunca se borra nada, así que se puede ejecutar varias veces.
    Devuelve {'creados': Counter, 'omitidos': Counter, 'avisos': [texto]}.
    """
    libro = load_workbook(origen, data_only=True)
    creados, omitidos, avisos = Counter(), Counter(), []

    asegurar_catalogos(db)
    categorias = {clave(c.nombre_categoria): c for c in db.scalars(select(models.Categorias))}
    proveedores = {clave(p.nombre_proveedor): p for p in db.scalars(select(models.Proveedores))}
    unidades = {clave(u.nombre_unidad): u for u in db.scalars(select(models.Unidades))}

    # --- catálogos (una columna, con encabezado en la fila 1)
    for palabra, modelo, campo, cache, etiqueta in (
        ('categoria', models.Categorias, 'nombre_categoria', categorias, 'categorías'),
        ('proveedor', models.Proveedores, 'nombre_proveedor', proveedores, 'proveedores'),
    ):
        ws = _hoja(libro, palabra)
        if ws is None:
            avisos.append(f'No encontré la hoja de {etiqueta}; se omite.')
            continue
        for fila in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            _asegurar(db, modelo, campo, fila[0], cache, creados, etiqueta)

    # --- resumen: un insumo por fila
    insumos = {clave(i.nombre_insumo): i for i in db.scalars(select(models.Insumos))}
    ws = _hoja(libro, 'resumen')
    inicio = _fila_de_datos(ws, 'nombre de insumo') if ws else None
    if inicio is None:
        avisos.append('No encontré la hoja de resumen (columna "Nombre de Insumo"); no se cargaron insumos.')
    else:
        for fila in ws.iter_rows(min_row=inicio, max_col=6, values_only=True):
            fila = list(fila) + [None] * 6
            nombre = limpiar(fila[0])
            if not nombre:
                continue
            if clave(nombre) in insumos:
                omitidos['insumos'] += 1
                continue
            insumo = models.Insumos(
                nombre_insumo=nombre,
                categoria=_asegurar(db, models.Categorias, 'nombre_categoria', fila[1], categorias, creados, 'categorías'),
                nombre_presentacion=limpiar(fila[2]) or None,
                unidad=_asegurar(db, models.Unidades, 'nombre_unidad', fila[3], unidades, creados, 'unidades'),
                proveedor=_asegurar(db, models.Proveedores, 'nombre_proveedor', fila[4], proveedores, creados, 'proveedores'),
                precio_sin_impuesto=_precio(fila[5]),
            )
            db.add(insumo)
            insumos[clave(nombre)] = insumo
            creados['insumos'] += 1

    # --- cotejo: filas de sección (en negritas) y productos
    ws = _hoja(libro, 'cotejo')
    inicio = _fila_de_datos(ws, 'producto') if ws else None
    if inicio is None:
        avisos.append('No encontré la hoja de cotejo (columna "Producto"); no se cargó la lista de productos.')
    else:
        areas = {clave(a.nombre_area): a for a in db.scalars(select(models.Areas))}
        existentes = {(clave(a.nombre_area), clave(p.nombre_producto)) for a in areas.values() for p in a.productos}
        area, orden = None, 0
        for fila in ws.iter_rows(min_row=inicio, max_col=3):
            nombre = limpiar(fila[0].value)
            if not nombre or clave(nombre).startswith('verde ='):   # fila vacía o leyenda de colores
                continue
            estado, nombres_plantilla = clave(fila[1].value), limpiar(fila[2].value)

            if fila[0].font.b and not estado:                        # fila de sección: PAN, CARNES...
                area = areas.get(clave(nombre))
                if area is None:
                    area = models.Areas(nombre_area=nombre, orden=len(areas) + 1)
                    db.add(area)
                    areas[clave(nombre)] = area
                    creados['áreas'] += 1
                orden = max((p.orden for p in area.productos), default=0)
                continue
            if area is None:
                avisos.append(f'«{nombre}» está antes de la primera sección; se omite.')
                continue

            orden += 1
            llave = (clave(area.nombre_area), clave(nombre))
            if llave in existentes:
                omitidos['productos'] += 1
                continue
            existentes.add(llave)

            vinculos = []
            for nom in re.split(r'\s*/\s*', nombres_plantilla):
                if not nom:
                    continue
                if clave(nom) in insumos:
                    vinculos.append(insumos[clave(nom)])
                else:
                    avisos.append(f'«{nombre}» dice que aparece como «{nom}», pero ese insumo no está en el resumen.')
            producto = models.Productos(
                area=area, nombre_producto=nombre, orden=orden,
                revisar=(estado == 'similar' and bool(vinculos)),
            )
            producto.insumos = vinculos
            db.add(producto)
            creados['productos'] += 1

    db.commit()
    return {'creados': creados, 'omitidos': omitidos, 'avisos': avisos}


# ==================================================================== EXPORTAR
_LADO = Side(style='thin')
_BORDE = Border(left=_LADO, right=_LADO, top=_LADO, bottom=_LADO)

# Colores de la plantilla original: (celdas, título, color del grupo, color de la columna)
_GRUPOS = [
    ('A1:B1', 'Información General', 'F9CB9C', 'FCE5CD'),
    ('C1:D1', 'Información de Presentación', 'B7D7A8', 'D9EAD3'),
    ('E1:F1', 'Información proveedor', 'A4C2F4', 'C9DBF8'),
]
_ENCABEZADOS = ['Nombre de Insumo', 'Categoria', 'Nombre de presentación',
                'UdM de Presentación', 'Proveedor', 'Precio sin impuesto']
_ANCHOS = {'A': 28, 'B': 16, 'C': 28, 'D': 20, 'E': 30, 'F': 20}


def _relleno(rgb):
    # 'FF' + color: opacidad total, igual que los archivos que genera Excel
    return PatternFill('solid', fgColor='FF' + rgb)


def _consulta_insumos(db):
    return db.scalars(
        select(models.Insumos)
        .options(joinedload(models.Insumos.categoria), joinedload(models.Insumos.unidad),
                 joinedload(models.Insumos.proveedor))
        .order_by(models.Insumos.id_insumo)
    ).all()


def _hoja_resumen(ws, insumos, unidades):
    """La plantilla: dos filas de encabezado (3 grupos de color) y un insumo por fila."""
    for rango, titulo, color_grupo, _ in _GRUPOS:
        ws[rango.split(':')[0]] = titulo
        ws.merge_cells(rango)
        for fila in ws[rango]:
            for celda in fila:
                celda.fill, celda.border = _relleno(color_grupo), _BORDE
    for n, (texto, (_, _, _, color)) in enumerate(zip(_ENCABEZADOS, [g for g in _GRUPOS for _ in (0, 1)]), start=1):
        celda = ws.cell(row=2, column=n, value=texto)
        celda.fill, celda.border = _relleno(color), _BORDE
    for fila in ws['A1:F2']:
        for celda in fila:
            celda.font = Font(name='Arial', size=11, bold=True)
            celda.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for n, i in enumerate(insumos, start=3):
        valores = [
            i.nombre_insumo,
            i.categoria.nombre_categoria if i.categoria else None,
            i.nombre_presentacion,
            i.unidad.nombre_unidad if i.unidad else None,
            i.proveedor.nombre_proveedor if i.proveedor else None,
            i.precio_sin_impuesto,          # vacío = todavía sin capturar
        ]
        for col, valor in enumerate(valores, start=1):
            celda = ws.cell(row=n, column=col, value=valor)
            celda.font = Font(name='Arial', size=11)
            celda.alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=n, column=6).number_format = '#,##0.00'

    # Lista desplegable de unidades (como la plantilla original), con espacio para filas nuevas
    validacion = DataValidation(type='list', formula1='"' + ','.join(unidades) + '"', allow_blank=True)
    ws.add_data_validation(validacion)
    validacion.add('D3:D1000')

    for letra, ancho in _ANCHOS.items():
        ws.column_dimensions[letra].width = ancho
    ws.freeze_panes = 'A3'


def _hoja_catalogo(ws, titulo, nombres):
    ws['A1'] = titulo
    ws['A1'].font = Font(name='Arial', size=11, bold=True)
    for n, nombre in enumerate(nombres, start=2):
        ws.cell(row=n, column=1, value=nombre).font = Font(name='Arial', size=11)
    ws.column_dimensions['A'].width = 34


def _hoja_cotejo(ws, areas):
    """Lista de productos por área. Verde = ya está en el resumen, amarillo = parecido, sin color = falta."""
    ws['A1'] = 'Lista de insumos'
    ws['A1'].font = Font(name='Arial', size=14, bold=True)
    ws['A2'] = 'Estado de cada producto de la lista. Generado por la aplicación de inventario.'
    ws['A2'].font = Font(name='Arial', size=9, italic=True, color='FF666666')
    for col, texto in enumerate(['Producto', '¿Aparece en plantilla?', 'Nombre en la plantilla'], start=1):
        celda = ws.cell(row=4, column=col, value=texto)
        celda.font = Font(name='Arial', size=11, bold=True, color='FFFFFFFF')
        celda.fill = _relleno('1F4E78')

    fila = 5
    for area in areas:
        ws.cell(row=fila, column=1, value=area.nombre_area)
        for col in (1, 2, 3):
            ws.cell(row=fila, column=col).fill = _relleno('D9E1F2')
            ws.cell(row=fila, column=col).font = Font(name='Arial', size=11, bold=True)
        fila += 1
        for p in area.productos:
            etiqueta, color = {'listo': ('Sí', 'C6EFCE'), 'revisar': ('Similar', 'FFEB9C')}.get(p.estado, (None, None))
            ws.cell(row=fila, column=1, value=p.nombre_producto)
            ws.cell(row=fila, column=2, value=etiqueta)
            ws.cell(row=fila, column=3, value=' / '.join(i.nombre_insumo for i in sorted(p.insumos, key=lambda i: i.id_insumo)) or None)
            for col in (1, 2, 3):
                celda = ws.cell(row=fila, column=col)
                celda.font = Font(name='Arial', size=11, bold=(col == 2 and bool(etiqueta)))
                if color:
                    celda.fill = _relleno(color)
            fila += 1
    ws.cell(row=fila + 1, column=1, value='Verde = aparece en la plantilla · Amarillo = hay un producto parecido/variante, revisar · Sin color = no aparece').font = Font(name='Arial', size=9, italic=True)
    for letra, ancho in {'A': 34, 'B': 24, 'C': 52}.items():
        ws.column_dimensions[letra].width = ancho


def _bytes(libro):
    salida = io.BytesIO()
    libro.save(salida)
    return salida.getvalue()


def libro_resumen(db):
    """Excel con solo la plantilla de resumen: el resultado que se busca del inventario."""
    unidades = [u.nombre_unidad for u in db.scalars(select(models.Unidades).order_by(models.Unidades.id_unidad))]
    libro = Workbook()
    ws = libro.active
    ws.title = 'Resumen'
    _hoja_resumen(ws, _consulta_insumos(db), unidades)
    return _bytes(libro)


def libro_respaldo(db):
    """Excel con las 4 hojas del libro original; importar() puede volver a leerlo."""
    unidades = [u.nombre_unidad for u in db.scalars(select(models.Unidades).order_by(models.Unidades.id_unidad))]
    areas = db.scalars(
        select(models.Areas).options(selectinload(models.Areas.productos).selectinload(models.Productos.insumos))
        .order_by(models.Areas.orden)
    ).all()
    libro = Workbook()
    ws = libro.active
    ws.title = HOJA_PROVEEDORES
    _hoja_catalogo(ws, 'PROVEEDORES', [p.nombre_proveedor for p in db.scalars(select(models.Proveedores).order_by(models.Proveedores.id_proveedor))])
    _hoja_catalogo(libro.create_sheet(HOJA_CATEGORIAS), 'CATEGORÍAS', [c.nombre_categoria for c in db.scalars(select(models.Categorias).order_by(models.Categorias.id_categoria))])
    _hoja_cotejo(libro.create_sheet(HOJA_COTEJO), areas)
    _hoja_resumen(libro.create_sheet(HOJA_RESUMEN), _consulta_insumos(db), unidades)
    return _bytes(libro)
