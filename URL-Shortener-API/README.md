# URL Shortener API

API para acortar URLs, hecha con Python y FastAPI.

## Tecnologias

- Python 3.10+
- FastAPI
- SQLAlchemy
- SQLite
- Pydantic

## Como usar

```bash
git clone https://github.com/acamilox/URL-Shortener-API.git
cd URL-Shortener-API
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Docs en http://localhost:8000/docs

## Endpoints

| Metodo | Ruta              | Descripcion                     |
|--------|-------------------|---------------------------------|
| POST   | /shorten          | Acortar una URL                 |
| GET    | /r/{code}         | Redirige a la URL original      |
| GET    | /stats/{code}     | Ver estadisticas de una URL     |

### POST /shorten

Enviar:

```json
{
  "url": "https://ejemplo.com/articulo-muy-largo"
}
```

Responde con la URL acortada y codigo 201.

### GET /r/{code}

Redirige automaticamente (301) a la URL original.

### GET /stats/{code}

Muestra la URL original y cuantas visitas tuvo.

## Estructura

```
app/
  main.py       # entrada, endpoints
  database.py   # conexion a SQLite
  models.py     # modelo ShortURL
  schemas.py    # esquemas Pydantic
  crud.py       # operaciones en BD
  utils.py      # generador de codigos
  config.py     # configuracion
requirements.txt
.gitignore
README.md
```

## Licencia

MIT
