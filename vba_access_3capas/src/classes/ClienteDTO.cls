VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClienteDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CLASE DTO: ClienteDTO (Objeto de Transferencia de Acreditado / Contraparte)
' -----------------------------------------------------------------------------
' Encapsula los datos de identificación de un acreditado en los portafolios NORMATIVA
' (clave, nombre o razón social, RFC y sector económico).
' =============================================================================

Private pCveCliente As String
Private pNombreCliente As String
Private pRfc As String
Private pSector As String

Public Function setCveCliente(ByVal C As String): pCveCliente = Trim(C): End Function
Public Function getCveCliente() As String: getCveCliente = pCveCliente: End Function
Public Function setNombreCliente(ByVal C As String): pNombreCliente = Trim(C): End Function
Public Function getNombreCliente() As String: getNombreCliente = pNombreCliente: End Function
Public Function setRfc(ByVal C As String): pRfc = Trim(C): End Function
Public Function getRfc() As String: getRfc = pRfc: End Function
Public Function setSector(ByVal C As String): pSector = Trim(C): End Function
Public Function getSector() As String: getSector = pSector: End Function
