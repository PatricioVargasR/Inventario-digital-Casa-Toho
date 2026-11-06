# Inventario digital de Casa Töhö

Herramienta para armar, entre varias personas y desde el celular, el **primer borrador del resumen de
insumos** (categoría, presentación, unidad, proveedor y precio sin impuesto) y descargarlo en Excel con el
formato de la plantilla. Misma tecnología que el Recetario: Flask + SQLAlchemy + Jinja, sin API; el
JavaScript solo se usa para un detalle visual.

## Cómo se usa

1. **Lista de cotejo.** Se elige un área (Pan, Carnes, Verduras…) y se recorre su lista de productos.
   Cada producto tiene uno de tres estados:
   - **Listo**: ya está vinculado con un insumo del resumen.
   - **Por revisar**: el Excel original lo marcaba como "Similar"; alguien debe confirmar el vínculo.
   - **Falta**: todavía no está en el resumen.
2. En un producto se puede **vincularlo** con un insumo que ya existe (la app sugiere los de nombre parecido
   para no duplicar) o **crear el insumo nuevo**, que nace vinculado.
3. **Resumen.** Muestra todos los insumos por categoría. Cada uno indica qué datos le faltan; se puede guardar
   un borrador con solo el nombre y completarlo después.
4. **Descargar Excel**: genera la hoja de resumen con el mismo formato que `Resumen_de_productos.xlsx`
   (encabezados por grupo, colores y lista desplegable de unidades).

## Puesta en marcha

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt

python seed.py                     # una sola vez: carga datos/Resumen_de_productos.xlsx
python app.py                      # http://127.0.0.1:5000
```

`python seed.py` lee las 4 hojas del Excel (Proveedores, Categorías, Cotejo Insumo y el Resumen). Es seguro
repetirlo: lo que ya existe se omite y nunca borra nada. También acepta otro libro:
`python seed.py otro.xlsx`, incluido el **respaldo** que descarga la app.

### Variables de entorno (todas opcionales)

| Variable       | Para qué sirve                                               | Por defecto               |
|----------------|--------------------------------------------------------------|---------------------------|
| `DATABASE_URL` | Base de datos (`sqlite:///...` o `postgresql+psycopg://...`) | `instance/inventario.db`  |
| `HOST`         | Interfaz donde escucha (`0.0.0.0` para abrirla en la red)    | `127.0.0.1`               |
| `PORT`         | Puerto                                                       | `5000`                    |
| `FLASK_DEBUG`  | `1` activa el modo debug (solo en desarrollo)                | apagado                   |

## Estructura

| Archivo                      | Qué hace                                                          |
|------------------------------|-------------------------------------------------------------------|
| `app.py`                     | Rutas, consultas y validación de formularios                      |
| `models.py`                  | Tablas (ver abajo)                                                |
| `schemas.py`                 | Validación de lo que llega de los formularios (Pydantic)          |
| `excel.py`                   | Importar el libro de Excel y exportar resumen / respaldo          |
| `seed.py`                    | Línea de comandos para importar                                   |
| `database.py`                | Conexión (motor, sesión, llaves foráneas de SQLite)               |
| `utiles.py`                  | Limpieza y comparación de textos                                  |
| `templates/`                 | `inicio`, `area`, `producto`, `insumos`, `insumo_form`, `catalogo`|
| `static/js/app.js`           | Un detalle visual (filtro por proveedor al cambiar el selector)   |
| `datos/`                     | El Excel original                                                 |

### Modelo de datos

```
areas ──< productos >──< producto_insumo >──< insumos >── categorias
(PAN, CARNES...)  (lista de cotejo)              (el resumen) ├─ proveedores
                                                              └─ unidades (kg, gr, lt, ml, pza, gal)
```

Un producto de la lista puede corresponder a varios insumos (por ejemplo "Crema" → 3 cremas). En `insumos`,
solo el nombre es obligatorio; `precio_sin_impuesto` vacío significa "sin capturar" y `0` es un precio válido.

## Rutas

| Método     | Ruta                              | Descripción                                          |
|------------|-----------------------------------|------------------------------------------------------|
| `GET`      | `/`                               | Avance general y áreas                               |
| `GET/POST` | `/areas/<id>`                     | Productos del área (`?q=`, `?estado=`); POST agrega  |
| `GET`      | `/productos/<id>`                 | Vincular, crear insumo, confirmar                    |
| `POST`     | `/productos/<id>/vincular`        | Vincular con un insumo                               |
| `POST`     | `/productos/<id>/desvincular`     | Quitar un vínculo                                    |
| `POST`     | `/productos/<id>/confirmar`       | Confirmar un vínculo "parecido"                      |
| `POST`     | `/productos/<id>/borrar`          | Quitar de la lista                                   |
| `GET`      | `/insumos`                        | Resumen (`?q=`, `?categoria=`, `?proveedor=`, `?pend=1`) |
| `GET/POST` | `/insumos/nuevo`                  | Crear (`?producto=<id>` lo vincula al guardar)       |
| `GET/POST` | `/insumos/<id>`                   | Ver y editar                                         |
| `POST`     | `/insumos/<id>/borrar`            | Borrar                                               |
| `GET/POST` | `/catalogos/<proveedores\|categorias>` | Ver y agregar                                   |
| `POST`     | `/catalogos/<tipo>/<id>` y `/borrar` | Renombrar; borrar si nadie lo usa                 |
| `GET`      | `/exportar/resumen.xlsx`          | Excel con la plantilla de resumen                    |
| `GET`      | `/exportar/respaldo.xlsx`         | Excel con las 4 hojas (se reimporta con `seed.py`)   |

## Cómo se trataron los datos del Excel original

- Se limpian espacios sobrantes ("Cebolla blanca " → "Cebolla blanca") y los nombres se comparan sin
  importar acentos ni mayúsculas.
- En la hoja Resumen, **precio `0.0` se importa como "sin capturar"** (75 de 97 insumos). Al exportar, los
  precios sin capturar quedan en blanco.
- En Cotejo: "Sí" → vinculado; "Similar" → vinculado **por revisar**; vacío → falta. Cuando una celda
  nombra varios insumos ("A / B") se crean todos los vínculos.
- Los catálogos se importan tal cual, sin unificar. Hay duplicados probables que conviene revisar en
  **Catálogos**: "Sin proveedor" / "Sin Provedor", "Milard" / "Maillard", "Doctora" / "Doctora Sonia".

## Despliegue y seguridad

Como el Recetario, necesita un hosting que ejecute Flask (`gunicorn app:app`). Si el servicio borra el disco
en cada despliegue, usa PostgreSQL con `DATABASE_URL`, o perderás los datos.

La app **no tiene inicio de sesión**: quien pueda abrirla puede editar y borrar. Los `POST` que el navegador
marca como originados en otro sitio se rechazan. `datos/*.xlsx` está en `.gitignore` porque contiene
nombres de proveedores y precios; no lo subas a un repositorio público.
