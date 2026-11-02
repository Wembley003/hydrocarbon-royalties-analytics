VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CLASE DTO: UsuarioDTO (Objeto de Transferencia de Datos de Empleado/Usuario)
' -----------------------------------------------------------------------------
' Representa un registro de la tabla USUARIOS de BD\RIESGOSBD.accdb.
' Transporta los datos del empleado autenticado (usuarioSesionDTO) o del
' empleado que se da de alta, consulta, modifica o inactiva entre las 3 capas
' (Presentación <-> FondoBO <-> FondoDAO) sin exponer el Recordset a las pantallas.
' =============================================================================

Private CVEUSUARIO As String
Private password As String
Private NOMBRE As String
Private APATERNO As String
Private AMATERNO As String
Private cvePerfil As String
Private PERFIL As String
Private EMAIL As String
Private CVEESTATUS As String
Private estatus As String

' Asigna y obtiene la clave de empleado de 4 dígitos (ej. "1000" a "1100")
Public Function setCveUsuario(ByVal C As String)
    CVEUSUARIO = Trim(C)
End Function

Public Function getCveUsuario() As String
    getCveUsuario = CVEUSUARIO
End Function

' Asigna y obtiene la contraseña (en texto plano al capturar o cifrada ENC$... desde BD)
Public Function setPassword(ByVal C As String)
    password = C
End Function

Public Function getPassword() As String
    getPassword = password
End Function

' Asigna y obtiene el(los) nombre(s) de pila del empleado
Public Function setNombre(ByVal C As String)
    NOMBRE = Trim(C)
End Function

Public Function getNombre() As String
    getNombre = NOMBRE
End Function

' Asigna y obtiene el apellido paterno del empleado
Public Function setAPaterno(ByVal C As String)
    APATERNO = Trim(C)
End Function

Public Function getAPaterno() As String
    getAPaterno = APATERNO
End Function

' Asigna y obtiene el apellido materno del empleado
Public Function setAMaterno(ByVal C As String)
    AMATERNO = Trim(C)
End Function

Public Function getAMaterno() As String
    getAMaterno = AMATERNO
End Function

' Asigna y obtiene la clave de perfil ("A" = Administrador, "U" = Usuario Operativo)
Public Function setCvePerfil(ByVal C As String)
    cvePerfil = Trim(C)
End Function

Public Function getCvePerfil() As String
    getCvePerfil = cvePerfil
End Function

' Asigna y obtiene la descripción textual del perfil
Public Function setPerfil(ByVal C As String)
    PERFIL = Trim(C)
End Function

Public Function getPerfil() As String
    getPerfil = PERFIL
End Function

' Asigna y obtiene el correo electrónico institucional (@energia-upstream.org)
Public Function setEmail(ByVal C As String)
    EMAIL = Trim(C)
End Function

Public Function getEmail() As String
    getEmail = EMAIL
End Function

' Asigna y obtiene la clave de estatus ("ACTIVO" o "INACTIVO")
Public Function setCveEstatus(ByVal C As String)
    CVEESTATUS = Trim(C)
End Function

Public Function getCveEstatus() As String
    getCveEstatus = CVEESTATUS
End Function

' Asigna y obtiene la descripción del estatus
Public Function setEstatus(ByVal C As String)
    estatus = Trim(C)
End Function

Public Function getEstatus() As String
    getEstatus = estatus
End Function

' Propiedades equivalentes conforme al estándar del Marco de Desarrollo
Public Property Get PropCveUsuario() As String
    PropCveUsuario = CVEUSUARIO
End Property

Public Property Let PropCveUsuario(ByVal valor As String)
    CVEUSUARIO = Trim(valor)
End Property

Public Property Get PropNombre() As String
    PropNombre = NOMBRE
End Property

Public Property Let PropNombre(ByVal valor As String)
    NOMBRE = Trim(valor)
End Property
