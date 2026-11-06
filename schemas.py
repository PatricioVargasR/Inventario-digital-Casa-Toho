"""
    Esquemas (Pydantic v2) para validar lo que llega de los formularios.
"""

from typing import Optional

from pydantic import BaseModel, ConfigDict, Field, field_validator


class _Esquema(BaseModel):
    # Quita espacios al inicio y al final de todos los textos
    model_config = ConfigDict(str_strip_whitespace=True)


class InsumosCreate(_Esquema):
    nombre_insumo: str = Field(min_length=1, max_length=150)
    # Todo lo demás es opcional: se puede guardar un borrador incompleto
    categoria_insumo: Optional[int] = None
    nombre_presentacion: Optional[str] = Field(default=None, max_length=80)
    unidad_presentacion: Optional[int] = None
    proveedor_insumo: Optional[int] = None
    precio_sin_impuesto: Optional[float] = Field(default=None, ge=0, le=10_000_000)

    @field_validator('nombre_presentacion', mode='after')
    @classmethod
    def _vacio_es_none(cls, valor):
        return valor or None


class ProductosCreate(_Esquema):
    nombre_producto: str = Field(min_length=1, max_length=150)


class NombreCatalogo(_Esquema):
    nombre: str = Field(min_length=1, max_length=80)
