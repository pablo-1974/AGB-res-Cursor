from sqlalchemy import Column, Integer, String, Date, Boolean, DateTime, Text
from sqlalchemy.sql import func

from .database import Base


class Usuario(Base):
    __tablename__ = "usuarios"

    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String, unique=True, nullable=False)
    email = Column(String, unique=True, nullable=False, index=True)
    # "admin" o "profesor"
    rol = Column(String, nullable=False)
    # Para simplificar el ejemplo se guarda en texto plano.
    # En un entorno real se debería almacenar un hash.
    password = Column(String, nullable=False)
    # Token de acceso tipo "Bearer"
    api_token = Column(String, unique=True, index=True, nullable=False)
    # El administrador puede activar/desactivar usuarios
    activo = Column(Boolean, nullable=False, default=True)


class Reserva(Base):
    __tablename__ = "reservas"

    id = Column(Integer, primary_key=True, index=True)
    # Nombre del profesor que hace la reserva
    profesor = Column(String, nullable=False)
    # Aula: 206 (Inf A), 210 (Inf B), 209 (Inf C), Biblioteca
    aula = Column(String, nullable=False)
    # Fecha de la reserva
    fecha = Column(Date, nullable=False)
    # Franja horaria (por ejemplo "08:40-09:30")
    franja_horaria = Column(String, nullable=False)


class Auditoria(Base):
    __tablename__ = "auditoria"

    id = Column(Integer, primary_key=True, index=True)
    creado_en = Column(DateTime, nullable=False, server_default=func.now(), index=True)

    # Actor (quién realiza la acción). Puede ser None en acciones internas.
    actor_usuario_id = Column(Integer, nullable=True, index=True)
    actor_nombre = Column(String, nullable=True)
    actor_rol = Column(String, nullable=True)
    actor_email = Column(String, nullable=True, index=True)
    actor_ip = Column(String, nullable=True)

    # Qué pasó
    accion = Column(String, nullable=False, index=True)
    entidad = Column(String, nullable=False, index=True)  # "Usuario", "Reserva", etc.
    entidad_id = Column(Integer, nullable=True, index=True)
    detalle = Column(Text, nullable=True)  # JSON serializado (string)

