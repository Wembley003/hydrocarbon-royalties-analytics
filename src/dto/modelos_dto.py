"""
CAPA TRANSVERSAL DE OBJETOS DE TRANSFERENCIA DE DATOS (10 DTOs)
"""
from dataclasses import dataclass


@dataclass
class UsuarioDTO:
    cve_usuario: str = ""
    password: str = ""
    nombre: str = ""
    a_paterno: str = ""
    a_materno: str = ""
    cve_perfil: str = "U"
    email: str = ""
    cve_estatus: str = "ACTIVO"


@dataclass
class ContratoDTO:
    id_contrato: str = ""
    cve_operador: str = ""
    modalidad_fiscal: str = ""
    tasa_regalia: float = 0.0
    limite_recuperacion_costos: float = 0.0
    cuenca_petrolera: str = ""


@dataclass
class OperadorDTO:
    cve_operador: str = ""
    nombre_operador: str = ""
    tipo_operador: str = ""
    calificacion_tecnica: str = "AAA"


@dataclass
class PerfilDTO:
    cve_perfil: str = ""
    descripcion: str = ""


@dataclass
class EstatusDTO:
    cve_estatus: str = ""
    descripcion: str = ""


@dataclass
class MonedaDTO:
    cve_moneda: str = ""
    descripcion: str = ""
    tipo_cambio_fix: float = 1.0


@dataclass
class TasaRegaliaDTO:
    cve_tasa: str = ""
    descripcion: str = ""
    valor_tasa: float = 0.0


@dataclass
class BloqueDTO:
    id_bloque: str = ""
    operador: str = ""
    cuenca: str = ""


@dataclass
class ContraprestacionesDTO:
    valor_contractual_mdp: float = 0.0
    contraprestacion_mdp: float = 0.0
    derecho_promedio_mdp: float = 0.0
    equivalente_usd_mdd: float = 0.0


@dataclass
class CuencaDTO:
    cve_cuenca: str = ""
    nombre_cuenca: str = ""
