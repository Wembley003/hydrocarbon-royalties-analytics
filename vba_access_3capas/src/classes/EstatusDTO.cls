VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstatusDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CLASE DTO: EstatusDTO (Objeto de Transferencia del Catálogo ESTATUS)
' -----------------------------------------------------------------------------
' Representa el estado operativo de una cuenta ("ACTIVO" o "INACTIVO").
' =============================================================================

Private pCveEstatus As String
Private pDescripcion As String

Public Sub setCveEstatus(ByVal val As String): pCveEstatus = Trim(val): End Sub
Public Sub setDescripcion(ByVal val As String): pDescripcion = Trim(val): End Sub
Public Function getCveEstatus() As String: getCveEstatus = pCveEstatus: End Function
Public Function getDescripcion() As String: getDescripcion = pDescripcion: End Function
