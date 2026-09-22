# API de Autenticacion con JWT

API con FastAPI, PostgreSQL, JWT y Docker.

## Endpoints

| Metodo | Ruta            | Descripcion                     |
|--------|-----------------|---------------------------------|
| POST   | /auth/register  | Registrar un usuario nuevo      |
| POST   | /auth/login     | Iniciar sesion (devuelve token) |
| GET    | /users/me       | Ver perfil del usuario logueado |
| GET    | /admin/users    | Listar usuarios (solo admin)    |

## Como usar

```bash
git clone <repo-url>
cd "API de Autenticacion con JWT"
cp .env.example .env
docker-compose up --build
```

Docs en http://localhost:8000/docs
