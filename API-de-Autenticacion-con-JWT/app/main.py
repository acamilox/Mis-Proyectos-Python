from fastapi import FastAPI, Depends, HTTPException, status
from fastapi.security import OAuth2PasswordRequestForm
from sqlalchemy.orm import Session
from typing import List

from . import models, schemas, crud, security, dependencies
from .config import settings
from .database import get_db

app = FastAPI(
    title="Auth API",
    description="API de autenticacion con JWT",
    version="1.0.0",
)


@app.on_event("startup")
def validate_secret_key():
    if settings.is_using_default_secret:
        raise RuntimeError(
            "SECRET_KEY tiene el valor por defecto. "
            "Define una clave segura en la variable de entorno SECRET_KEY."
        )


@app.get("/")
def root():
    return {"message": "Auth API funcionando", "docs": "/docs"}


@app.post("/auth/register", response_model=schemas.UserOut, status_code=201)
def register(user: schemas.UserCreate, db: Session = Depends(get_db)):
    if crud.get_user_by_username(db, user.username):
        raise HTTPException(status_code=400, detail="El usuario ya existe")
    if crud.get_user_by_email(db, user.email):
        raise HTTPException(status_code=400, detail="El email ya esta registrado")
    return crud.create_user(db, user)


@app.post("/auth/login", response_model=schemas.Token)
def login(
    form_data: OAuth2PasswordRequestForm = Depends(),
    db: Session = Depends(get_db),
):
    user = crud.authenticate_user(db, form_data.username, form_data.password)
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Usuario o contrasena incorrectos",
        )
    token = security.create_access_token({"sub": user.username, "role": user.role})
    return {"access_token": token, "token_type": "bearer"}


@app.get("/users/me", response_model=schemas.UserOut)
def read_me(current_user: models.User = Depends(security.get_current_user)):
    return current_user


@app.get("/admin/users", response_model=List[schemas.UserOut])
def list_users(
    skip: int = 0,
    limit: int = 100,
    db: Session = Depends(get_db),
    current_user: models.User = Depends(dependencies.require_admin),
):
    return crud.get_users(db, skip=skip, limit=limit)
