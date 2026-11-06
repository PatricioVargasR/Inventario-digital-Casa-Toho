"""
    Archivo que almacena los modelos de la
    base de datos para el ORM

    Dos mundos que se conectan:
      * La LISTA DE COTEJO (areas -> productos): lo que hay físicamente en la cocina
        y que hay que revisar.
      * El RESUMEN (insumos): la plantilla que se quiere obtener, con categoría,
        presentación, proveedor y precio.
    Un producto de la lista se vincula con uno o varios insumos del resumen
    (tabla producto_insumo).
"""

from sqlalchemy import (
    Boolean, CheckConstraint, Column, Float, ForeignKey, Integer, Table, Text, UniqueConstraint,
)
from sqlalchemy.orm import relationship

# Importar el objeto de la base de datos
from database import Base


# ------------------------------------------------------------------ catálogos
class Categorias(Base):
    __tablename__ = 'categorias'
    __table_args__ = {'sqlite_autoincrement': True}

    id_categoria = Column(Integer, primary_key=True, autoincrement=True)
    nombre_categoria = Column(Text, nullable=False, unique=True)


class Proveedores(Base):
    __tablename__ = 'proveedores'
    __table_args__ = {'sqlite_autoincrement': True}

    id_proveedor = Column(Integer, primary_key=True, autoincrement=True)
    nombre_proveedor = Column(Text, nullable=False, unique=True)


class Unidades(Base):
    """Unidades de medida de la presentación: kg, gr, lt, ml, pza, gal."""
    __tablename__ = 'unidades'
    __table_args__ = {'sqlite_autoincrement': True}

    id_unidad = Column(Integer, primary_key=True, autoincrement=True)
    nombre_unidad = Column(Text, nullable=False, unique=True)


# ------------------------------------------------------- resumen (el resultado)
class Insumos(Base):
    __tablename__ = 'insumos'
    __table_args__ = (
        CheckConstraint('precio_sin_impuesto IS NULL OR precio_sin_impuesto >= 0', name='ck_insumos_precio'),
        {'sqlite_autoincrement': True},
    )

    id_insumo = Column(Integer, primary_key=True, autoincrement=True)
    nombre_insumo = Column(Text, nullable=False, unique=True)
    # Todo lo demás es opcional: el inventario es un BORRADOR que se va completando
    categoria_insumo = Column(Integer, ForeignKey('categorias.id_categoria', ondelete='RESTRICT'), nullable=True)
    nombre_presentacion = Column(Text, nullable=True)          # Caja, Costal, Kilo, Pieza...
    unidad_presentacion = Column(Integer, ForeignKey('unidades.id_unidad', ondelete='RESTRICT'), nullable=True)
    proveedor_insumo = Column(Integer, ForeignKey('proveedores.id_proveedor', ondelete='RESTRICT'), nullable=True)
    # NULL = todavía no se captura. 0 es un precio válido (por ejemplo, producción interna).
    precio_sin_impuesto = Column(Float, nullable=True)

    categoria = relationship('Categorias')
    unidad = relationship('Unidades')
    proveedor = relationship('Proveedores')
    productos = relationship('Productos', secondary='producto_insumo', back_populates='insumos')

    @property
    def faltantes(self):
        """Datos que todavía no se han capturado."""
        faltan = []
        if self.categoria_insumo is None:
            faltan.append('categoría')
        if not (self.nombre_presentacion or '').strip():
            faltan.append('presentación')
        if self.unidad_presentacion is None:
            faltan.append('unidad')
        if self.proveedor_insumo is None:
            faltan.append('proveedor')
        if self.precio_sin_impuesto is None:
            faltan.append('precio')
        return faltan

    @property
    def completo(self):
        return not self.faltantes

    @property
    def detalle(self):
        """'Costal · kg · Super Superet' (solo lo que exista)."""
        partes = [
            (self.nombre_presentacion or '').strip(),
            self.unidad.nombre_unidad if self.unidad else '',
            self.proveedor.nombre_proveedor if self.proveedor else '',
        ]
        return ' · '.join(p for p in partes if p)


# ---------------------------------------------------- lista de cotejo (lo físico)
class Areas(Base):
    """Secciones de la cocina: PAN, CARNES, REFRIGERACIÓN, VERDURAS..."""
    __tablename__ = 'areas'
    __table_args__ = {'sqlite_autoincrement': True}

    id_area = Column(Integer, primary_key=True, autoincrement=True)
    nombre_area = Column(Text, nullable=False, unique=True)
    orden = Column(Integer, nullable=False, default=0)

    productos = relationship(
        'Productos', back_populates='area', cascade='all, delete-orphan',
        order_by='Productos.orden, Productos.id_producto',
    )


# Un producto de la lista puede corresponder a varios insumos (p. ej. "Crema" -> 3 cremas)
producto_insumo = Table(
    'producto_insumo', Base.metadata,
    Column('id_producto', Integer, ForeignKey('productos.id_producto', ondelete='CASCADE'), primary_key=True),
    Column('id_insumo', Integer, ForeignKey('insumos.id_insumo', ondelete='CASCADE'), primary_key=True),
)


class Productos(Base):
    __tablename__ = 'productos'
    __table_args__ = (
        UniqueConstraint('area_producto', 'nombre_producto', name='uq_productos_area_nombre'),
        {'sqlite_autoincrement': True},
    )

    id_producto = Column(Integer, primary_key=True, autoincrement=True)
    area_producto = Column(Integer, ForeignKey('areas.id_area', ondelete='CASCADE'), nullable=False, index=True)
    nombre_producto = Column(Text, nullable=False)
    orden = Column(Integer, nullable=False, default=0)
    # True = el vínculo con el resumen es "parecido" y alguien debe confirmarlo
    revisar = Column(Boolean, nullable=False, default=False)

    area = relationship('Areas', back_populates='productos')
    insumos = relationship('Insumos', secondary=producto_insumo, back_populates='productos')

    @property
    def estado(self):
        """falta = sin insumo en el resumen · revisar = parecido, por confirmar · listo = ya está."""
        if not self.insumos:
            return 'falta'
        return 'revisar' if self.revisar else 'listo'
