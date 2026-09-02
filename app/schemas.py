from pydantic import BaseModel

# Schema base con los campos comunes
class TaskBase(BaseModel):
    title: str
    description: str | None = None

# Schema para la creación de una tarea (hereda de TaskBase)
class TaskCreate(TaskBase):
    pass

# Schema para actualizar una tarea (todos los campos opcionales)
class TaskUpdate(BaseModel):
    title: str | None = None
    description: str | None = None
    completed: bool | None = None

# Schema para leer/devolver una tarea (incluye campos de la DB)
class Task(TaskBase):
    id: int
    completed: bool

    class Config:
        from_attributes = True