from sqlalchemy.orm import Session
from app import models, schemas

# Función para OBTENER todas las tareas
def get_tasks(db: Session, skip: int = 0, limit: int = 100):
    return db.query(models.Task).offset(skip).limit(limit).all()

# Función para CREAR una tarea
def create_task(db: Session, task: schemas.TaskCreate):
    # Crea un objeto Task de SQLAlchemy a partir de los datos del schema
    db_task = models.Task(title=task.title, description=task.description)
    db.add(db_task)  # Añade el objeto a la sesión de la base de datos
    db.commit()      # Guarda los cambios en la base de datos
    db.refresh(db_task) # Refresca el objeto para obtener el ID asignado por la DB
    return db_task

# (Añade esto al final del archivo crud.py)

# Función para OBTENER una sola tarea por su ID
def get_task(db: Session, task_id: int):
    return db.query(models.Task).filter(models.Task.id == task_id).first()

# Función para ACTUALIZAR una tarea
def update_task(db: Session, task_id: int, task: schemas.TaskUpdate):
    db_task = db.query(models.Task).filter(models.Task.id == task_id).first()
    if db_task:
        task_data = task.model_dump(exclude_unset=True)
        for key, value in task_data.items():
            setattr(db_task, key, value)
        db.commit()
        db.refresh(db_task)
    return db_task

# Función para BORRAR una tarea
def delete_task(db: Session, task_id: int):
    db_task = db.query(models.Task).filter(models.Task.id == task_id).first()
    if db_task:
        db.delete(db_task)
        db.commit()
        return db_task
    return None