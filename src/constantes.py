"""
MÓDULO TRANSVERSAL: constantes.py
=============================================================================
Centraliza las constantes del Sistema de Fiscalización, Valoración Contractual
y Liquidación de Regalías de Hidrocarburos (Upstream E&P).
Arquitectura de 3 Capas: Presentación -> Negocio (BO) -> Persistencia (DAO + SQL).
"""
import os

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
PROJECT_NAME = "Sistema de Fiscalización, Valoración Contractual y Regalías de Hidrocarburos"
UNIDAD_ADMINISTRATIVA = "Dirección Corporativa de Economía y Fiscalización de Hidrocarburos"
AREA_TECNICA = "Gerencia de Verificación de Volúmenes, Regalías y Contratos E&P"

# Base de datos física externa alojada en la subcarpeta BD/ (fuera del módulo principal)
DB_DIR = os.path.join(PROJECT_ROOT, "BD")
DB_NAME = "HIDROCARBUROS_UPSTREAM.sqlite"
DB_PATH = os.path.join(DB_DIR, DB_NAME)

# Plantilla genérica con campos vacíos y fórmulas vivas SUMIFS/SUM (nunca se modifica el original)
TEMPLATES_DIR = os.path.join(PROJECT_ROOT, "Templates")
TEMPLATE_GENERICO = "Generico.xlsx"
TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, TEMPLATE_GENERICO)

INSUMOS_DIR = os.path.join(PROJECT_ROOT, "Insumos")
REPORTES_DIR = os.path.join(PROJECT_ROOT, "Reportes_Generados")
DICTAMENES_DIR = os.path.join(PROJECT_ROOT, "Dictamenes_Generados")
BI_DIR = os.path.join(PROJECT_ROOT, "bi_powerbi_tableau")

PERFIL_ADMINISTRADOR = "A"
PERFIL_VERIFICADOR = "U"
CLAVE_ESTATUS_ACTIVO = "ACTIVO"
CLAVE_ESTATUS_INACTIVO = "INACTIVO"
PASSWORD_PREDETERMINADA = "Wembley1"

# Las 3 bases de datos independientes de insumos por régimen contractual + repositorio consolidado
TABLA_INSUMO_ASIGNACIONES = "BD_INSUMO_ASIGNACIONES"
TABLA_INSUMO_LICENCIAS = "BD_INSUMO_LICENCIAS"
TABLA_INSUMO_PRODUCCION_COMPARTIDA = "BD_INSUMO_PRODUCCION_COMPARTIDA"
TABLA_CONSOLIDADA_CONTRAPRESTACIONES = "PADRON_CONTRAPRESTACIONES_CONSOLIDADO"

HOJA_TPL_REPORTE = "REPORTE"
HOJA_TPL_DATOS = "DATOS"
HOJA_TPL_CALCULO = "CALCULO"
HOJA_TPL_CALCINSTR = "CALC_INSTR"

TPL_PORTAFOLIO_1 = "1. Asignaciones Directas de Extraccion (Derechos sobre Hidrocarburos)"
TPL_PORTAFOLIO_2 = "2. Contratos de Licencia (Regalias Basicas, Adicionales y Cuotas)"
TPL_PORTAFOLIO_3 = "3. Contratos de Produccion Compartida (Utilidad Operativa y Costos)"

# Esquema oficial de 53 columnas para fiscalización de hidrocarburos
COLUMNAS_53_OFICIALES = [
    "P", "REGIMEN", "G", "MODALIDAD_CONTRACTUAL", "ESTATA", "CUENCA_PETROLERA",
    "SG", "AREA_CONTRACTUAL", "CAMPOA", "OPERADOR_PETROLERO", "RFC", "MONEDA2",
    "MON", "MONEDA", "EF", "ENT_FED", "CONCATENADO", "CTA",
    "SCATAS", "SCTA", "SSCTA", "SSSCTA", "TIPO_YACIMIENTO", "C",
    "CONCEPTO_FISCAL", "SC", "SCONCEPTO", "SSCONCEPTO", "DESTINO_PRODUCCION",
    "SDINMO", "SDFIMO", "PROMMO", "SDFIMN", "SDPRMN", "SDFIDL", "SDPRDL",
    "MESPRO", "ANOPRO", "ACT", "CAMPO_EXTRACCION", "GRUCON", "GRUPO_CONT_FISCAL",
    "GRUCONTNORM", "ARTICULO_NORMATIVO", "CODIGO_HIDROCARBURO",
    "DESC_HIDROCARBURO", "CODIGO_VERIFICADOR", "NOMBRE_VERIFICADOR",
    "DESCRIPCIÓN CUENTA", "SITUACION DEL CONTRATO", "CONCEPTO_ETAPA",
    "FECHA_INSUMO", "ID_CONTRATO"
]

COLUMNAS_ENTERAS_1BASED = [1, 3, 7, 9, 13, 15, 18, 19, 20, 21, 22, 24, 26, 37, 38, 39, 41, 43, 45, 47]
