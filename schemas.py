from datetime import date, datetime
from enum import Enum
import json
from typing import Any

from pydantic import BaseModel, field_validator


class AulaEnum(str, Enum):
    AULA_206_INF_A = "Aula 206 (Inf A)"
    AULA_210_INF_B = "Aula 210 (Inf B)"
    AULA_209_INF_C = "Aula 209 (Inf C)"
    BIBLIOTECA = "Biblioteca"


class FranjaHorariaEnum(str, Enum):
    H1 = "08:40-09:30"
    H2 = "09:35-10:25"
    H3 = "10:30-11:20"
    H4 = "11:50-12:40"
    H5 = "12:45-13:35"
    H6 = "13:40-14:30"


class RolEnum(str, Enum):
    ADMIN = "admin"
    PROFESOR = "profesor"


class UsuarioBase(BaseModel):
    nombre: str
    email: str
    rol: RolEnum
    activo: bool = True


class UsuarioCreate(UsuarioBase):
    password: str


class UsuarioOut(UsuarioBase):
    id: int

    class Config:
        from_attributes = True


class LoginRequest(BaseModel):
    email: str
    password: str


class TokenOut(BaseModel):
    token: str
    rol: RolEnum
    nombre: str
    email: str


class ResetPasswordOut(BaseModel):
    id: int
    email: str
    nueva_password: str


class ChangePasswordRequest(BaseModel):
    password_actual: str
    password_nueva: str


class ReservaBase(BaseModel):
    profesor: str
    aula: AulaEnum
    fecha: date
    franja_horaria: FranjaHorariaEnum

    @field_validator("profesor")
    @classmethod
    def profesor_no_vacio(cls, v: str) -> str:
        if not v.strip():
            raise ValueError("El nombre del profesor no puede estar vacío")
        return v


class ReservaCreate(ReservaBase):
    pass


class ReservaOut(ReservaBase):
    id: int

    class Config:
        from_attributes = True


class ReservaRecurrenteCreate(BaseModel):
    profesor: str | None = None
    aula: AulaEnum
    fecha_inicio: date
    franja_horaria: FranjaHorariaEnum


class AuditoriaOut(BaseModel):
    id: int
    creado_en: datetime

    actor_usuario_id: int | None = None
    actor_nombre: str | None = None
    actor_rol: RolEnum | None = None
    actor_email: str | None = None
    actor_ip: str | None = None

    accion: str
    entidad: str
    entidad_id: int | None = None
    detalle: dict[str, Any] | None = None

    @field_validator("detalle", mode="before")
    @classmethod
    def parse_detalle_json(cls, v: Any) -> Any:
        if v is None:
            return None
        if isinstance(v, dict):
            return v
        if isinstance(v, str):
            try:
                return json.loads(v)
            except Exception:
                return {"raw": v}
        return {"raw": str(v)}

    class Config:
        from_attributes = True


class AuditoriaGroupedOut(BaseModel):
    """
    Estructura agrupada para el registro de auditoría.
    - reservas: altas/bajas/modificaciones relacionadas con reservas.
    - otros: resto de acciones (usuarios, auth, etc.).
    """
    reservas: list[AuditoriaOut]
    otros: list[AuditoriaOut]


class UsuarioRolUpdate(BaseModel):
    rol: RolEnum

