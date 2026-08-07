# Dashboard POA Institucional 2026 — Django

Dashboard interactivo (Django + plantillas) para consultar el Plan Operativo Anual:
búsqueda por unidad, avance de cumplimiento, actividades por unidad y observaciones.
Incluye descarga del POA completo en CSV (Excel) y JSON.

## Requisitos
- Python 3.10 o superior

## Instalación y ejecución

```bash
cd poa_django
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python manage.py runserver
```

Abre http://127.0.0.1:8000/

## Estructura

```
poa_django/
├── manage.py
├── requirements.txt
├── data/poa2026.json            # dataset extraído del Excel del POA
├── poa_project/                 # settings / urls / wsgi / asgi
└── dashboard/
    ├── views.py                 # dashboard, API JSON y descargas
    ├── urls.py
    ├── templates/dashboard/index.html
    └── static/dashboard/{app.css,app.js}
```

## Rutas

| Ruta | Descripción |
|------|-------------|
| `/` | Dashboard interactivo |
| `/api/poa/?q=&categoria=` | API JSON con búsqueda y filtro por categoría |
| `/descargar/csv/` | Todas las actividades en CSV (separador `;`, con BOM para Excel) |
| `/descargar/json/` | Dataset completo en JSON |

## Actualizar los datos

Reemplaza `data/poa2026.json` con un nuevo dataset (misma estructura) y reinicia el
servidor. La estructura esperada es:

```json
{
  "anio": 2026,
  "unidades": [
    {
      "id": "slug", "nombre": "...", "titulo": "...", "categoria": "Unidad Administrativa",
      "responsable": "...", "avance": 0.42, "totalActividades": 18,
      "actividades": [
        { "objetivo": "...", "tarea": "...", "prioridad": "...", "fecha": "...",
          "responsable": "...", "entregable": "...", "avance": 0.5, "observacion": "..." }
      ]
    }
  ],
  "proyectos": [{ "nombre": "...", "descripcion": "...", "responsable": "...", "avance": 0.3, "observacion": "..." }]
}
```

## Notas de despliegue
`settings.py` está en modo desarrollo (`DEBUG = True`, `SECRET_KEY` de ejemplo).
Para producción: define `SECRET_KEY` por variable de entorno, `DEBUG = False`,
`ALLOWED_HOSTS` concreto y sirve los estáticos con `python manage.py collectstatic`.

## Panel de administrador (subir Excel)

| Ruta | Descripción |
|------|-------------|
| `/panel/login/` | Acceso con usuario administrador |
| `/panel/` | Subir el Excel del POA y ver el estado de los datos |

1. Crea un usuario administrador (una sola vez):

```bash
python manage.py migrate
python manage.py createsuperuser
```

2. Entra a `/panel/login/`, sube el archivo `.xlsx` y elige el año.
   El sistema convierte el Excel a `data/poa2026.json`, guarda un respaldo del
   archivo anterior en `data/respaldos/` y refresca el dashboard sin reiniciar
   el servidor.

### Formato esperado del Excel
- Una hoja por unidad (el nombre de la hoja es el nombre de la unidad; se admite el
  prefijo numérico, p. ej. `2. UNIDAD DE COMUNICACIÓN`).
- Una hoja cuyo nombre contenga «Proyecto» para la lista de proyectos.
- En cada hoja, una fila de encabezados con columnas reconocibles por su título:
  Objetivo, Tarea/Actividad, Prioridad, Fecha, Responsable, Entregable,
  Avance (0-1, 0-100 o `50%`) y Observación. El orden de las columnas es libre.
- El título de la unidad y el responsable se leen de las filas previas al encabezado.
