VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerfilDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CLASE DTO: PerfilDTO (Objeto de Transferencia del Catálogo PERFILES)
' -----------------------------------------------------------------------------
' Representa un rol de seguridad del sistema ("A" = Administrador, "U" = Operativo).
' =============================================================================

Private pCvePerfil As String
Private pDescripcion As String

Public Sub setCvePerfil(ByVal val As String): pCvePerfil = Trim(val): End Sub
Public Sub setDescripcion(ByVal val As String): pDescripcion = Trim(val): End Sub
Public Function getCvePerfil() As String: getCvePerfil = pCvePerfil: End Function
Public Function getDescripcion() As String: getDescripcion = pDescripcion: End Function
