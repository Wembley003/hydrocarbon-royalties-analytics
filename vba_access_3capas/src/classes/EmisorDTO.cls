VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmisorDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CLASE DTO: EmisorDTO (Objeto de Transferencia del Catálogo EMISORES)
' -----------------------------------------------------------------------------
' Encapsula los atributos de un emisor financiero registrado en EMISORES
' (clave de emisor, razón social, sector económico y calificación de riesgo).
' =============================================================================

Private pCveEmisor As String
Private pNombreEmisor As String
Private pSector As String
Private pCalificacionRiesgo As String

Public Sub setCveEmisor(ByVal val As String): pCveEmisor = Trim(val): End Sub
Public Sub setNombreEmisor(ByVal val As String): pNombreEmisor = Trim(val): End Sub
Public Sub setSector(ByVal val As String): pSector = Trim(val): End Sub
Public Sub setCalificacionRiesgo(ByVal val As String): pCalificacionRiesgo = Trim(val): End Sub

Public Function getCveEmisor() As String: getCveEmisor = pCveEmisor: End Function
Public Function getNombreEmisor() As String: getNombreEmisor = pNombreEmisor: End Function
Public Function getSector() As String: getSector = pSector: End Function
Public Function getCalificacionRiesgo() As String: getCalificacionRiesgo = pCalificacionRiesgo: End Function
