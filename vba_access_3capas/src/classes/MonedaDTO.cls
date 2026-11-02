VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MonedaDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CLASE DTO: MonedaDTO (Objeto de Transferencia del Catálogo CAT_MONEDAS)
' -----------------------------------------------------------------------------
' Representa una divisa de valuación (MXN, USD, UDI) registrada en CAT_MONEDAS.
' =============================================================================

Private pCveMoneda As String
Private pDescripcion As String

Public Sub setCveMoneda(ByVal val As String): pCveMoneda = Trim(val): End Sub
Public Sub setDescripcion(ByVal val As String): pDescripcion = Trim(val): End Sub
Public Function getCveMoneda() As String: getCveMoneda = pCveMoneda: End Function
Public Function getDescripcion() As String: getDescripcion = pDescripcion: End Function
