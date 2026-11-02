VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TasaDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CLASE DTO: TasaDTO (Objeto de Transferencia del Catálogo CAT_TASAS)
' -----------------------------------------------------------------------------
' Representa una tasa de referencia financiera (TIIE_28, CETES_28, SOFR_USD).
' =============================================================================

Private pCveTasa As String
Private pValor As Double

Public Sub setCveTasa(ByVal val As String): pCveTasa = Trim(val): End Sub
Public Sub setValor(ByVal val As Double): pValor = val: End Sub
Public Function getCveTasa() As String: getCveTasa = pCveTasa: End Function
Public Function getValor() As Double: getValor = pValor: End Function
