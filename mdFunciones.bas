Attribute VB_Name = "mdFunciones"
Option Explicit

Function formarCodBarras(ByVal codBarrasOr As String) As String
  formarCodBarras = codBarrasOr & comprobarDigitoControl(codBarrasOr)
End Function

Function comprobarDigitoControl(ByVal codigoBarras As String) As Byte
  Dim digito As Byte, calTotal As Byte
  Dim codTmp As String, bPal As Byte, numC As Byte

  Select Case Len(codigoBarras)
  Case 7, 12
    codTmp = Right$("0000000000000000" & codigoBarras, 17)
    bPal = 3
    For numC = 1 To 17
        calTotal = calTotal + Val(Mid$(codTmp, numC, 1)) * bPal
        bPal = 4 - bPal
    Next
    digito = calTotal Mod 10
    digito = IIf(digito = 0, 0, 10 - digito)
  End Select
  comprobarDigitoControl = digito
End Function

Sub msgAviso(ByVal textoAviso As String)
  MsgBox textoAviso, vbExclamation, App.Title
End Sub

Function medIn(ByVal vTextoTmp, ByVal vPosicion)
  medIn = CInt(Mid(vTextoTmp, vPosicion, 1))
End Function
