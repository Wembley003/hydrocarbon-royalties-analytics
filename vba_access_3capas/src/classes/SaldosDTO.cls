VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SaldosDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CLASE DTO: SaldosDTO (Objeto de Transferencia de Saldos Valuados NORMATIVA)
' -----------------------------------------------------------------------------
' Encapsula los importes de saldo inicial (SDINMO), saldo final (SDFIMN) y
' saldo promedio (SDPRMN) agrupados por cuenta y moneda.
' =============================================================================

Private pCuenta As String
Private pMoneda As String
Private pSaldoInicial As Double
Private pSaldoFinal As Double
Private pSaldoPromedio As Double

Public Function setCuenta(ByVal C As String): pCuenta = Trim(C): End Function
Public Function getCuenta() As String: getCuenta = pCuenta: End Function
Public Function setMoneda(ByVal C As String): pMoneda = Trim(C): End Function
Public Function getMoneda() As String: getMoneda = pMoneda: End Function
Public Function setSaldoInicial(ByVal V As Double): pSaldoInicial = V: End Function
Public Function getSaldoInicial() As Double: getSaldoInicial = pSaldoInicial: End Function
Public Function setSaldoFinal(ByVal V As Double): pSaldoFinal = V: End Function
Public Function getSaldoFinal() As Double: getSaldoFinal = pSaldoFinal: End Function
Public Function setSaldoPromedio(ByVal V As Double): pSaldoPromedio = V: End Function
Public Function getSaldoPromedio() As Double: getSaldoPromedio = pSaldoPromedio: End Function
