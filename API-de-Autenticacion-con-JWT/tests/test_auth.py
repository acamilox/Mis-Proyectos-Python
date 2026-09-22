def test_register_creates_user(client):
    response = client.post(
        "/auth/register",
        json={
            "username": "camilo",
            "email": "camilo@example.com",
            "password": "secreto123",
        },
    )
    assert response.status_code == 201
    data = response.json()
    assert data["username"] == "camilo"
    assert data["email"] == "camilo@example.com"
    assert data["role"] == "user"
    assert "hashed_password" not in data


def test_register_duplicate_username(client):
    user = {
        "username": "camilo",
        "email": "camilo@example.com",
        "password": "secreto123",
    }
    client.post("/auth/register", json=user)
    response = client.post("/auth/register", json=user)
    assert response.status_code == 400


def test_register_invalid_email(client):
    response = client.post(
        "/auth/register",
        json={"username": "camilo", "email": "no-es-un-email", "password": "secreto123"},
    )
    assert response.status_code == 422


def test_login_returns_token(client):
    client.post(
        "/auth/register",
        json={
            "username": "camilo",
            "email": "camilo@example.com",
            "password": "secreto123",
        },
    )
    response = client.post(
        "/auth/login",
        data={"username": "camilo", "password": "secreto123"},
    )
    assert response.status_code == 200
    data = response.json()
    assert "access_token" in data
    assert data["token_type"] == "bearer"


def test_login_wrong_password(client):
    client.post(
        "/auth/register",
        json={
            "username": "camilo",
            "email": "camilo@example.com",
            "password": "secreto123",
        },
    )
    response = client.post(
        "/auth/login",
        data={"username": "camilo", "password": "incorrecta"},
    )
    assert response.status_code == 401


def test_users_me_requires_auth(client):
    response = client.get("/users/me")
    assert response.status_code == 401


def test_users_me_returns_profile(client):
    client.post(
        "/auth/register",
        json={
            "username": "camilo",
            "email": "camilo@example.com",
            "password": "secreto123",
        },
    )
    token = client.post(
        "/auth/login",
        data={"username": "camilo", "password": "secreto123"},
    ).json()["access_token"]

    response = client.get(
        "/users/me", headers={"Authorization": f"Bearer {token}"}
    )
    assert response.status_code == 200
    assert response.json()["username"] == "camilo"


def test_admin_users_denied_for_normal_user(client):
    client.post(
        "/auth/register",
        json={
            "username": "camilo",
            "email": "camilo@example.com",
            "password": "secreto123",
        },
    )
    token = client.post(
        "/auth/login",
        data={"username": "camilo", "password": "secreto123"},
    ).json()["access_token"]

    response = client.get(
        "/admin/users", headers={"Authorization": f"Bearer {token}"}
    )
    assert response.status_code == 403


def test_admin_users_allowed_for_admin(client):
    # Crear un usuario admin directamente en la base de datos
    from app import crud, schemas, security

    admin = schemas.UserCreate(
        username="admin", email="admin@example.com", password="admin123"
    )
    crud.create_user(db=client.app.state.db, user=admin)
    # Promover a admin
    session = client.app.state.db
    user_obj = crud.get_user_by_username(session, "admin")
    user_obj.role = "admin"
    session.commit()

    token = client.post(
        "/auth/login",
        data={"username": "admin", "password": "admin123"},
    ).json()["access_token"]

    response = client.get(
        "/admin/users", headers={"Authorization": f"Bearer {token}"}
    )
    assert response.status_code == 200
    assert any(u["username"] == "admin" for u in response.json())
