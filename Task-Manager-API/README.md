# Task Manager API

API para gestionar tareas, hecha con Python y FastAPI.

## Tecnologias

- Python 3.10+
- FastAPI
- SQLAlchemy
- SQLite
- Pydantic
- Uvicorn

## Como usar

```bash
git clone https://github.com/acamilox/Task-Manager-API.git
cd Task-Manager-API
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

La API queda en `http://127.0.0.1:8000` y los docs en `http://127.0.0.1:8000/docs`.

## Endpoints

| Metodo | Ruta              | Descripcion          |
|--------|-------------------|----------------------|
| POST   | /tasks/           | Crear una tarea      |
| GET    | /tasks/           | Listar tareas        |
| GET    | /tasks/{id}       | Obtener tarea por ID |
| DELETE | /tasks/{id}       | Eliminar una tarea   |

### POST /tasks/

```json
{
  "title": "Comprar víveres",
  "description": "Leche, pan, huevos"
}
```

Responde con la tarea creada y codigo 201.

### GET /tasks/

Query params opcionales: `skip` (default 0) y `limit` (default 100).

### GET /tasks/{id}

Devuelve la tarea o 404 si no existe.

### DELETE /tasks/{id}

Elimina la tarea y la devuelve. 404 si no existe.

## Estructura

```
app/
  main.py       # entrada de la app, endpoints
  database.py   # conexion a SQLite con SQLAlchemy
  models.py     # modelo Task (id, title, description, completed)
  schemas.py    # esquemas Pydantic para validar datos
  crud.py       # funciones para leer/escribir en la base de datos
requirements.txt
.gitignore
README.md
```

## Licencia

MIT
