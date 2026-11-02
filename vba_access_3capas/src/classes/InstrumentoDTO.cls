VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstrumentoDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CLASE DTO: InstrumentoDTO (Objeto de Transferencia de Instrumento Financiero)
' -----------------------------------------------------------------------------
' Representa un registro del catálogo INSTRUMENTOS de BD\RIESGOSBD.accdb.
' Se utiliza para dar de alta, consultar y actualizar instrumentos financieros
' (bonos, CETES, UDIBONOS, certificados bursátiles) entre Presentación, BO y DAO.
' =============================================================================

Private pIdInstrumento As String
Private pCveEmisor As String
Private pTipoInstrumento As String
Private pTasa As Double
Private pPlazo As Long
Private pFechaEmision As Date
Private pFechaVencimiento As Date

' Métodos de escritura (Setters)
Public Sub setIdInstrumento(ByVal val As String): pIdInstrumento = Trim(val): End Sub
Public Sub setCveEmisor(ByVal val As String): pCveEmisor = Trim(val): End Sub
Public Sub setTipoInstrumento(ByVal val As String): pTipoInstrumento = Trim(val): End Sub
Public Sub setTasa(ByVal val As Double): pTasa = val: End Sub
Public Sub setPlazo(ByVal val As Long): pPlazo = val: End Sub
Public Sub setFechaEmision(ByVal val As Date): pFechaEmision = val: End Sub
Public Sub setFechaVencimiento(ByVal val As Date): pFechaVencimiento = val: End Sub

' Métodos de lectura (Getters)
Public Function getIdInstrumento() As String: getIdInstrumento = pIdInstrumento: End Function
Public Function getCveEmisor() As String: getCveEmisor = pCveEmisor: End Function
Public Function getTipoInstrumento() As String: getTipoInstrumento = pTipoInstrumento: End Function
Public Function getTasa() As Double: getTasa = pTasa: End Function
Public Function getPlazo() As Long: getPlazo = pPlazo: End Function
Public Function getFechaEmision() As Date: getFechaEmision = pFechaEmision: End Function
Public Function getFechaVencimiento() As Date: getFechaVencimiento = pFechaVencimiento: End Function
