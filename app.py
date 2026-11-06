"""
    Inventario digital de Casa Töhö (Flask + SQLAlchemy, aplicación monolítica)

    Todo se resuelve en el servidor: Flask consulta la base de datos y entrega HTML ya armado.
    Las modificaciones se hacen con formularios normales (POST -> redirección).

    Flujo de trabajo:
        1. Se recorre la LISTA DE COTEJO por áreas (PAN, CARNES, VERDURAS...).
        2. Cada producto se vincula con un insumo del RESUMEN, o se crea el insumo
           capturando categoría, presentación, unidad, proveedor y precio.
        3. El RESUMEN se descarga en Excel con el formato de la plantilla.

    Rutas:
        GET        /                              Inicio: avance general y áreas
        GET, POST  /areas/<id>                    Lista de productos de un área (POST: agregar producto)
        GET        /productos/<id>                Un producto: vincular, crear insumo, confirmar
        POST       /productos/<id>/vincular       Vincular con un insumo existente
        POST       /productos/<id>/desvincular    Quitar un vínculo
        POST       /productos/<id>/confirmar      Confirmar un vínculo "parecido"
        POST       /productos/<id>/borrar         Quitar el producto de la lista
        GET        /insumos                       RESUMEN (?q=&categoria=&proveedor=&pend=1)
        GET, POST  /insumos/nuevo                 Crear insumo (?producto=<id> lo vincula al guardar)
        GET, POST  /insumos/<id>                  Ver/editar un insumo
        POST       /insumos/<id>/borrar           Borrar un insumo
        GET, POST  /catalogos/<tipo>              Proveedores o categorías (POST: agregar)
        POST       /catalogos/<tipo>/<id>         Renombrar
        POST       /catalogos/<tipo>/<id>/borrar  Borrar (solo si nadie lo usa)
        GET        /exportar/resumen.xlsx         Excel con la plantilla de resumen
        GET        /exportar/respaldo.xlsx        Excel con las 4 hojas (se reimporta con seed.py)
"""

import os
import re
from collections import Counter
from datetime import date

from flask import Flask, Response, abort, g, redirect, render_template, request, url_for
from pydantic import ValidationError
from sqlalchemy import func, select
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import joinedload, selectinload

import excel
import models
import schemas
from database import Base, LocalSession, engine
from utiles import clave, limpiar, moneda, numero

from dotenv import load_dotenv

# Inicializamos los modelos de la base de datos y las unidades de medida
# Base.metadata.create_all(bind=engine)
# with LocalSession() as _db:
#     excel.asegurar_catalogos(_db)

# Iniciamos la aplicación de flask
app = Flask(__name__)

load_dotenv()

# Avisos que se muestran tras una acción. Se pasan como ?ok=clave y solo se aceptan estas claves.
MENSAJES = {
    'creado': 'Insumo creado ✓',
    'guardado': 'Insumo guardado ✓',
    'borrado': 'Insumo borrado',
    'vinculado': 'Insumo vinculado ✓',
    'desvinculado': 'Vínculo quitado',
    'confirmado': 'Producto confirmado ✓',
    'agregado': 'Producto agregado a la lista ✓',
    'producto_borrado': 'Producto quitado de la lista',
    'catalogo_agregado': 'Agregado ✓',
    'catalogo_guardado': 'Guardado ✓',
    'catalogo_borrado': 'Borrado',
}

ESTADOS = {
    'falta': 'Falta',
    'revisar': 'Revisar',
    'listo': 'Listo',
}

# Catálogos que se pueden administrar desde la app (las unidades de medida son fijas)
CATALOGOS = {
    'proveedores': {'modelo': models.Proveedores, 'id': 'id_proveedor', 'campo': 'nombre_proveedor',
                    'uso': models.Insumos.proveedor_insumo, 'titulo': 'Proveedores', 'uno': 'proveedor'},
    'categorias': {'modelo': models.Categorias, 'id': 'id_categoria', 'campo': 'nombre_categoria',
                   'uso': models.Insumos.categoria_insumo, 'titulo': 'Categorías', 'uno': 'categoría'},
}


# --------------------------------------------------------------- utilidades
def get_db():
    """Una sesión de base de datos por petición."""
    if 'db' not in g:
        g.db = LocalSession()
    return g.db


@app.teardown_appcontext
def cerrar_db(_error):
    db = g.pop('db', None)
    if db is not None:
        db.close()  # si algo quedó sin commit, se revierte


@app.before_request
def rechazar_post_de_otros_sitios():
    # Los formularios no llevan token CSRF: se rechazan los POST que el navegador
    # marca como originados en otro sitio (Sec-Fetch-Site lo calcula el navegador).
    if request.method == 'POST' and request.headers.get('Sec-Fetch-Site') == 'cross-site':
        abort(403)


@app.context_processor
def contexto():
    return {'mensaje': MENSAJES.get(request.args.get('ok')), 'estados': ESTADOS}


app.add_template_filter(numero, 'numero')
app.add_template_filter(moneda, 'moneda')


@app.template_filter('bonito')
def bonito(texto):
    """'MESA FRÍA 1' -> 'Mesa fría 1' (los nombres de área vienen en mayúsculas del Excel)."""
    return texto.capitalize() if texto.isupper() else texto


@app.template_filter('faltan_texto')
def faltan_texto(faltantes):
    """['precio'] -> 'Falta precio'; ['precio', 'unidad'] -> 'Faltan 2 datos'."""
    if not faltantes:
        return ''
    return f'Falta {faltantes[0]}' if len(faltantes) == 1 else f'Faltan {len(faltantes)} datos'


def guardar(db):
    """Confirma los cambios. Devuelve un mensaje de error o None."""
    try:
        db.commit()
    except IntegrityError:
        db.rollback()
        return 'No se pudo guardar: el nombre ya existe o hay datos inválidos.'
    return None


def porcentaje(parte, total):
    return round(parte / total * 100, 1) if total else 0


_VACIAS = {'con', 'sin', 'del', 'las', 'los', 'una', 'uno', 'por', 'que', 'para'}


def _palabras(texto):
    """Palabras significativas de un nombre, en singular y sin acentos: 'Aguacates' -> 'aguacate'."""
    palabras = set()
    for p in re.findall(r'[a-z0-9]+', clave(texto)):
        if len(p) > 4 and p.endswith('s'):
            p = p[:-1]
        if len(p) >= 3 and p not in _VACIAS:
            palabras.add(p)
    return palabras


def sugerir(nombre, insumos, limite=5):
    """Insumos cuyo nombre comparte palabras con el producto: ayuda a no duplicar insumos."""
    buscadas = _palabras(nombre)
    puntuados = []
    for insumo in insumos:
        comunes = len(buscadas & _palabras(insumo.nombre_insumo))
        if comunes:
            puntuados.append((-comunes, clave(insumo.nombre_insumo), insumo))
    puntuados.sort(key=lambda t: t[:2])
    return [t[2] for t in puntuados[:limite]]


# ------------------------------------------------------------------- inicio
@app.get('/')
def inicio():
    """Avance general y una fila por área con su barra de avance."""
    db = get_db()
    areas = db.scalars(
        select(models.Areas)
        .options(selectinload(models.Areas.productos).selectinload(models.Productos.insumos))
        .order_by(models.Areas.orden)
    ).all()
    insumos = db.scalars(select(models.Insumos)).all()

    filas, total = [], Counter()
    for area in areas:
        c = Counter(p.estado for p in area.productos)
        total.update(c)
        n = len(area.productos)
        filas.append({
            'area': area, 'total': n, 'listo': c['listo'], 'revisar': c['revisar'], 'falta': c['falta'],
            'p_listo': porcentaje(c['listo'], n), 'p_revisar': porcentaje(c['revisar'], n),
        })

    n_productos = sum(total.values())
    return render_template(
        'inicio.html',
        filas=filas,
        n_productos=n_productos,
        productos_listos=total['listo'],
        n_insumos=len(insumos),
        insumos_completos=sum(1 for i in insumos if i.completo),
        p_productos=porcentaje(total['listo'], n_productos),
    )


# -------------------------------------------------------------------- áreas
def mostrar_area(id_area, error=None, nuevo=''):
    db = get_db()
    area = db.scalar(
        select(models.Areas)
        .options(selectinload(models.Areas.productos).selectinload(models.Productos.insumos))
        .where(models.Areas.id_area == id_area)
    )
    if area is None:
        abort(404)

    q = request.args.get('q', '').strip()
    estado = request.args.get('estado') if request.args.get('estado') in ESTADOS else None
    buscado = clave(q)

    def coincide(p):
        if estado and p.estado != estado:
            return False
        if buscado:
            return buscado in clave(' '.join([p.nombre_producto, *(i.nombre_insumo for i in p.insumos)]))
        return True

    return render_template(
        'area.html',
        area=area,
        productos=[p for p in area.productos if coincide(p)],
        conteo=Counter(p.estado for p in area.productos),
        total=len(area.productos),
        q=q,
        estado=estado,
        error=error,
        nuevo=nuevo,
    ), (422 if error else 200)


@app.get('/areas/<int:id_area>')
def area(id_area):
    return mostrar_area(id_area)


@app.post('/areas/<int:id_area>')
def agregar_producto(id_area):
    """Agrega a la lista un producto que no estaba."""
    db = get_db()
    area = db.get(models.Areas, id_area)
    if area is None:
        abort(404)
    nombre = limpiar(request.form.get('nombre_producto'))
    try:
        datos = schemas.ProductosCreate.model_validate({'nombre_producto': nombre})
    except ValidationError:
        return mostrar_area(id_area, error='Escribe el nombre del producto.', nuevo=nombre)
    if any(clave(p.nombre_producto) == clave(datos.nombre_producto) for p in area.productos):
        return mostrar_area(id_area, error='Ese producto ya está en la lista de esta área.', nuevo=nombre)

    producto = models.Productos(
        area=area, nombre_producto=datos.nombre_producto,
        orden=max((p.orden for p in area.productos), default=0) + 1,
    )
    db.add(producto)
    error = guardar(db)
    if error:
        return mostrar_area(id_area, error=error, nuevo=nombre)
    return redirect(url_for('area', id_area=id_area, ok='agregado', _anchor=f'producto-{producto.id_producto}'))


# ---------------------------------------------------------------- productos
def obtener_producto(id_producto):
    producto = get_db().scalar(
        select(models.Productos)
        .options(joinedload(models.Productos.area), selectinload(models.Productos.insumos))
        .where(models.Productos.id_producto == id_producto)
    )
    if producto is None:
        abort(404)
    return producto


def volver_al_area(producto, ok):
    return redirect(url_for('area', id_area=producto.area_producto, ok=ok, _anchor=f'producto-{producto.id_producto}'))


@app.get('/productos/<int:id_producto>')
def producto(id_producto):
    db = get_db()
    producto = obtener_producto(id_producto)
    vinculados = sorted(producto.insumos, key=lambda i: clave(i.nombre_insumo))
    ya = {i.id_insumo for i in vinculados}
    disponibles = sorted(
        (i for i in db.scalars(select(models.Insumos)).all() if i.id_insumo not in ya),
        key=lambda i: clave(i.nombre_insumo),
    )
    return render_template(
        'producto.html',
        producto=producto,
        vinculados=vinculados,
        sugerencias=sugerir(producto.nombre_producto, disponibles),
        disponibles=disponibles,
    )


@app.post('/productos/<int:id_producto>/vincular')
def vincular(id_producto):
    db = get_db()
    producto = obtener_producto(id_producto)
    insumo = db.get(models.Insumos, request.form.get('id_insumo', type=int) or 0)
    if insumo is None:
        abort(400)
    if insumo not in producto.insumos:
        producto.insumos.append(insumo)
    db.commit()
    return redirect(url_for('producto', id_producto=id_producto, ok='vinculado'))


@app.post('/productos/<int:id_producto>/desvincular')
def desvincular(id_producto):
    db = get_db()
    producto = obtener_producto(id_producto)
    insumo = db.get(models.Insumos, request.form.get('id_insumo', type=int) or 0)
    if insumo is None:
        abort(400)
    if insumo in producto.insumos:
        producto.insumos.remove(insumo)
    if not producto.insumos:
        producto.revisar = False
    db.commit()
    return redirect(url_for('producto', id_producto=id_producto, ok='desvinculado'))


@app.post('/productos/<int:id_producto>/confirmar')
def confirmar(id_producto):
    producto = obtener_producto(id_producto)
    producto.revisar = False
    get_db().commit()
    return volver_al_area(producto, 'confirmado')


@app.post('/productos/<int:id_producto>/borrar')
def borrar_producto(id_producto):
    db = get_db()
    producto = obtener_producto(id_producto)
    id_area = producto.area_producto
    db.delete(producto)
    db.commit()
    return redirect(url_for('area', id_area=id_area, ok='producto_borrado'))


# ------------------------------------------------------------------ insumos
def forma_vacia(producto=None):
    return {
        'nombre': producto.nombre_producto if producto else limpiar(request.args.get('nombre')),
        'categoria': '', 'presentacion': '', 'unidad': '', 'proveedor': '', 'precio': '',
    }


def forma_desde_insumo(i):
    return {
        'nombre': i.nombre_insumo,
        'categoria': str(i.categoria_insumo or ''),
        'presentacion': i.nombre_presentacion or '',
        'unidad': str(i.unidad_presentacion or ''),
        'proveedor': str(i.proveedor_insumo or ''),
        'precio': numero(i.precio_sin_impuesto),
    }


def forma_desde_request():
    f = request.form
    return {
        'nombre': limpiar(f.get('nombre_insumo')),
        'categoria': f.get('categoria_insumo', '').strip(),
        'presentacion': limpiar(f.get('nombre_presentacion')),
        'unidad': f.get('unidad_presentacion', '').strip(),
        'proveedor': f.get('proveedor_insumo', '').strip(),
        'precio': f.get('precio_sin_impuesto', '').strip(),
    }


def validar_insumo(db, forma, id_insumo=None):
    """Valida la forma. Devuelve (datos, None) o (None, mensaje_de_error)."""
    try:
        datos = schemas.InsumosCreate.model_validate({
            'nombre_insumo': forma['nombre'],
            'categoria_insumo': forma['categoria'] or None,
            'nombre_presentacion': forma['presentacion'] or None,
            'unidad_presentacion': forma['unidad'] or None,
            'proveedor_insumo': forma['proveedor'] or None,
            # "$12,50" se acepta como 12.5
            'precio_sin_impuesto': forma['precio'].replace('$', '').replace(',', '.').strip() or None,
        })
    except ValidationError as e:
        campo = e.errors(include_url=False, include_context=False, include_input=False)[0]['loc'][0]
        return None, {
            'nombre_insumo': 'Escribe el nombre del insumo.',
            'precio_sin_impuesto': 'El precio debe ser un número mayor o igual a 0 (por ejemplo 12.50).',
        }.get(campo, 'Elige una opción válida de la lista.')

    for campo, modelo, nombre in (
        (datos.categoria_insumo, models.Categorias, 'categoría'),
        (datos.unidad_presentacion, models.Unidades, 'unidad'),
        (datos.proveedor_insumo, models.Proveedores, 'proveedor'),
    ):
        if campo is not None and db.get(modelo, campo) is None:
            return None, f'La opción elegida en «{nombre}» ya no existe. Elige otra.'

    for otro in db.execute(select(models.Insumos.id_insumo, models.Insumos.nombre_insumo)):
        if otro.id_insumo != id_insumo and clave(otro.nombre_insumo) == clave(datos.nombre_insumo):
            return None, f'Ya existe un insumo llamado «{otro.nombre_insumo}».'
    return datos, None


def mostrar_insumo(forma, insumo=None, producto=None, error=None):
    db = get_db()
    return render_template(
        'insumo_form.html',
        forma=forma,
        insumo=insumo,
        producto=producto,
        error=error,
        categorias=sorted(db.scalars(select(models.Categorias)), key=lambda c: clave(c.nombre_categoria)),
        unidades=db.scalars(select(models.Unidades).order_by(models.Unidades.id_unidad)).all(),
        proveedores=sorted(db.scalars(select(models.Proveedores)), key=lambda p: clave(p.nombre_proveedor)),
    ), (422 if error else 200)


@app.get('/insumos')
def insumos():
    """El RESUMEN: todos los insumos, agrupados por categoría."""
    db = get_db()
    q = request.args.get('q', '').strip()
    categoria = request.args.get('categoria', type=int)
    proveedor = request.args.get('proveedor', type=int)
    solo_pend = request.args.get('pend') == '1'

    todos = db.scalars(
        select(models.Insumos).options(
            joinedload(models.Insumos.categoria), joinedload(models.Insumos.unidad), joinedload(models.Insumos.proveedor)
        )
    ).all()
    buscado = clave(q)

    def coincide(i):
        if categoria and i.categoria_insumo != categoria:
            return False
        if proveedor and i.proveedor_insumo != proveedor:
            return False
        if solo_pend and i.completo:
            return False
        return not buscado or buscado in clave(i.nombre_insumo)

    grupos = {}
    for i in sorted(filter(coincide, todos), key=lambda i: clave(i.nombre_insumo)):
        grupos.setdefault(i.categoria, []).append(i)
    # grupos por nombre de categoría; "Sin categoría" al final
    secciones = sorted(grupos.items(), key=lambda kv: (kv[0] is None, clave(kv[0].nombre_categoria) if kv[0] else ''))

    categorias_usadas = {i.categoria for i in todos if i.categoria}
    proveedores_usados = {i.proveedor for i in todos if i.proveedor}
    return render_template(
        'insumos.html',
        secciones=secciones,
        chips=sorted(categorias_usadas, key=lambda c: clave(c.nombre_categoria)),
        proveedores=sorted(proveedores_usados, key=lambda p: clave(p.nombre_proveedor)),
        total=len(todos),
        completos=sum(1 for i in todos if i.completo),
        q=q, categoria=categoria, proveedor=proveedor, solo_pend=solo_pend,
    )


@app.route('/insumos/nuevo', methods=['GET', 'POST'])
def nuevo_insumo():
    db = get_db()
    id_producto = request.values.get('producto', type=int)
    producto = db.get(models.Productos, id_producto) if id_producto else None

    if request.method == 'POST':
        forma = forma_desde_request()
        datos, error = validar_insumo(db, forma)
        if not error:
            insumo = models.Insumos(**datos.model_dump())
            if producto:                       # el insumo nace vinculado al producto que se estaba revisando
                insumo.productos.append(producto)
                producto.revisar = False
            db.add(insumo)
            error = guardar(db)
            if not error:
                if producto:
                    return volver_al_area(producto, 'creado')
                return redirect(url_for('insumos', ok='creado', _anchor=f'insumo-{insumo.id_insumo}'))
        return mostrar_insumo(forma, producto=producto, error=error)
    return mostrar_insumo(forma_vacia(producto), producto=producto)


@app.route('/insumos/<int:id_insumo>', methods=['GET', 'POST'])
def editar_insumo(id_insumo):
    db = get_db()
    insumo = db.scalar(
        select(models.Insumos)
        .options(selectinload(models.Insumos.productos).joinedload(models.Productos.area))
        .where(models.Insumos.id_insumo == id_insumo)
    )
    if insumo is None:
        abort(404)

    if request.method == 'POST':
        forma = forma_desde_request()
        datos, error = validar_insumo(db, forma, id_insumo)
        if not error:
            for campo, valor in datos.model_dump().items():
                setattr(insumo, campo, valor)
            error = guardar(db)
            if not error:
                return redirect(url_for('insumos', ok='guardado', _anchor=f'insumo-{id_insumo}'))
        return mostrar_insumo(forma, insumo=insumo, error=error)
    return mostrar_insumo(forma_desde_insumo(insumo), insumo=insumo)


@app.post('/insumos/<int:id_insumo>/borrar')
def borrar_insumo(id_insumo):
    db = get_db()
    insumo = db.get(models.Insumos, id_insumo)
    if insumo is None:
        abort(404)
    db.delete(insumo)   # los productos que lo usaban vuelven a "Falta" solos
    db.commit()
    return redirect(url_for('insumos', ok='borrado'))


# ---------------------------------------------------------------- catálogos
def obtener_catalogo(tipo):
    if tipo not in CATALOGOS:
        abort(404)
    return CATALOGOS[tipo]


def mostrar_catalogo(tipo, error=None, nuevo=''):
    cat = obtener_catalogo(tipo)
    db = get_db()
    modelo, campo = cat['modelo'], cat['campo']
    usos = dict(db.execute(select(cat['uso'], func.count()).group_by(cat['uso'])).all())
    filas = sorted(
        ({'id': getattr(o, cat['id']), 'nombre': getattr(o, campo), 'usos': usos.get(getattr(o, cat['id']), 0)}
         for o in db.scalars(select(modelo))),
        key=lambda f: clave(f['nombre']),
    )
    return render_template('catalogo.html', tipo=tipo, cat=cat, filas=filas, error=error, nuevo=nuevo), (422 if error else 200)


def validar_nombre_catalogo(db, cat, nombre, id_actual=None):
    """Devuelve (nombre_limpio, None) o (None, mensaje)."""
    try:
        nombre = schemas.NombreCatalogo.model_validate({'nombre': limpiar(nombre)}).nombre
    except ValidationError:
        return None, f'Escribe el nombre del {cat["uno"]} (máximo 80 letras).'
    for otro in db.scalars(select(cat['modelo'])):
        if getattr(otro, cat['id']) != id_actual and clave(getattr(otro, cat['campo'])) == clave(nombre):
            return None, f'Ya existe «{getattr(otro, cat["campo"])}».'
    return nombre, None


@app.get('/catalogos')
def catalogos():
    return redirect(url_for('catalogo', tipo='proveedores'))


@app.get('/catalogos/<tipo>')
def catalogo(tipo):
    return mostrar_catalogo(tipo)


@app.post('/catalogos/<tipo>')
def agregar_catalogo(tipo):
    cat = obtener_catalogo(tipo)
    db = get_db()
    nombre, error = validar_nombre_catalogo(db, cat, request.form.get('nombre'))
    if not error:
        db.add(cat['modelo'](**{cat['campo']: nombre}))
        error = guardar(db)
    if error:
        return mostrar_catalogo(tipo, error=error, nuevo=limpiar(request.form.get('nombre')))
    return redirect(url_for('catalogo', tipo=tipo, ok='catalogo_agregado'))


@app.post('/catalogos/<tipo>/<int:id_>')
def renombrar_catalogo(tipo, id_):
    cat = obtener_catalogo(tipo)
    db = get_db()
    objeto = db.get(cat['modelo'], id_)
    if objeto is None:
        abort(404)
    nombre, error = validar_nombre_catalogo(db, cat, request.form.get('nombre'), id_actual=id_)
    if not error:
        setattr(objeto, cat['campo'], nombre)
        error = guardar(db)
    if error:
        return mostrar_catalogo(tipo, error=error)
    return redirect(url_for('catalogo', tipo=tipo, ok='catalogo_guardado', _anchor=f'fila-{id_}'))


@app.post('/catalogos/<tipo>/<int:id_>/borrar')
def borrar_catalogo(tipo, id_):
    cat = obtener_catalogo(tipo)
    db = get_db()
    objeto = db.get(cat['modelo'], id_)
    if objeto is None:
        abort(404)
    if db.scalar(select(func.count()).select_from(models.Insumos).where(cat['uso'] == id_)):
        return mostrar_catalogo(tipo, error=f'No se puede borrar «{getattr(objeto, cat["campo"])}»: hay insumos que lo usan.')
    db.delete(objeto)
    db.commit()
    return redirect(url_for('catalogo', tipo=tipo, ok='catalogo_borrado'))


# --------------------------------------------------------------- exportar
XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


def descargar(contenido, nombre):
    return Response(contenido, mimetype=XLSX, headers={'Content-Disposition': f'attachment; filename={nombre}'})


@app.get('/exportar/resumen.xlsx')
def exportar_resumen():
    """La plantilla de resumen (el resultado del inventario), con el formato del Excel original."""
    return descargar(excel.libro_resumen(get_db()), f'inventario-resumen-{date.today().isoformat()}.xlsx')


@app.get('/exportar/respaldo.xlsx')
def exportar_respaldo():
    """Las 4 hojas del libro original. `python seed.py archivo.xlsx` lo vuelve a cargar."""
    return descargar(excel.libro_respaldo(get_db()), f'inventario-respaldo-{date.today().isoformat()}.xlsx')


if __name__ == '__main__':
    # debug=True con host 0.0.0.0 expone el depurador de Werkzeug (ejecución remota de código)
    # a toda la red. El modo debug y el host se activan solo si se piden.
    app.run(
        debug=os.environ.get('FLASK_DEBUG') == '1',
        host=os.environ.get('HOST', '127.0.0.1'),
        port=int(os.environ.get('PORT', '5000')),
    )
