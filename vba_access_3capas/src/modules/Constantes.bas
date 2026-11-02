Attribute VB_Name = "Constantes"
Option Compare Database
Option Explicit

' =============================================================================
' MÓDULO TRANSVERSAL: Constantes.bas
' -----------------------------------------------------------------------------
' Centraliza todas las constantes institucionales del Módulo de Administración
' Integral de Riesgos de Hidrocarburos: rutas relativas hacia la base de datos
' externa (BD\RIESGOSBD.accdb), plantillas, credenciales maestras de cifrado,
' claves de perfil/estatus, nombres de tablas de insumo y la variable global
' de sesión activa (usuarioSesionDTO).
' =============================================================================

'******** Identificación del Aplicativo y Rutas de Archivos ********
Public Const PROJECT_NAME As String = "Módulo de Administración Integral de Riesgos"
Public Const DB_PATH As String = "\BD\"
Public Const DB_NAME As String = "RIESGOSBD.accdb"
Public Const PASSWORD_PREDETERMINADA As String = "Wembley1"
Public Const PATH_IMAGENES As String = "\Archivos\Imagenes\"
Public Const PROYECTO As String = "\Modulo_Administracion_Riesgos.accdb"

'******** Plantilla Oficial de Cálculo en Excel ********
Public Const PATH_TEMPLATES As String = "\Templates\"
' Sin acento a propósito para evitar problemas de codificación en rutas de Windows
Public Const TEMPLATE_GENERICO As String = "Generico.xlsx"

'******** Perfiles de Usuario (Control de Acceso Basado en Roles) ********
Public Const PERFIL_USUARIO As String = "U"
Public Const PERFIL_ADMINISTRADOR As String = "A"
Public Const PERFIL_EMPLEADO As String = "U"

'******** Estatus de Cuenta en el Catálogo USUARIOS ********
Public Const CLAVE_ESTATU_ACTIVO As String = "ACTIVO"
Public Const CLAVE_ESTATU_INACTIVO As String = "INACTIVO"

'******** Mensaje Institucional Estándar ante Errores ********
Public Const MENSAJE_ERROR As String = "Ocurrió un error en la operación. Favor de contactar al Área de Riesgos."

'******** Variable Global de Sesión Activa (Instancia UsuarioDTO) ********
' Conserva en memoria los datos del empleado que inició sesión en Form_Acceso
Public usuarioSesionDTO As UsuarioDTO

'******** Tablas Independientes de Insumos (alojadas en BD\RIESGOSBD.accdb) ********
Public Const TABLA_INSUMO_PRIMER_PISO As String = "BD_INSUMO_PRIMER_PISO"
Public Const TABLA_INSUMO_TESORERIA As String = "BD_INSUMO_TESORERIA"
Public Const TABLA_INSUMO_SEGUNDO_PISO As String = "BD_INSUMO_SEGUNDO_PISO"

'******** Constantes para el puntero del ratón en operaciones de espera ********
Public Const MOUSE_POINTER_WAIT As Integer = 11
Public Const MOUSE_POINTER_DEFAULT As Integer = 1

'******** Bandera de Caché de Vínculos e Índices del Backend en la Sesión Actual ********
Public gBackendVerificadoEnSesion As Boolean
