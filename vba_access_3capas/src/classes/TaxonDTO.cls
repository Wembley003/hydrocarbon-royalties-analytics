VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaxonDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CLASE DTO: TaxonDTO (Objeto de Transferencia de Clasificación / Límite NORMATIVA)
' -----------------------------------------------------------------------------
' Encapsula un segmento o taxón de riesgo con su clave, descripción y límite
' porcentual de exposición normativa.
' =============================================================================

Private pCveTaxon As String
Private pDescripcion As String
Private pPorcentajeLimite As Double

Public Function setCveTaxon(ByVal C As String): pCveTaxon = Trim(C): End Function
Public Function getCveTaxon() As String: getCveTaxon = pCveTaxon: End Function
Public Function setDescripcion(ByVal C As String): pDescripcion = Trim(C): End Function
Public Function getDescripcion() As String: getDescripcion = pDescripcion: End Function
Public Function setPorcentajeLimite(ByVal V As Double): pPorcentajeLimite = V: End Function
Public Function getPorcentajeLimite() As Double: getPorcentajeLimite = pPorcentajeLimite: End Function
