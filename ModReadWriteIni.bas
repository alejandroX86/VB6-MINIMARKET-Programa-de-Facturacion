Attribute VB_Name = "ModReadWriteIni"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As _
String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As _
String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As _
Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'##########################################################
'######################LEER / ESCRIBIR INI #######################
'##########################################################

Public Function ReadINI(CORCHETEALEER As Variant, PROPIEDADALEER As String) As String
Dim valor As String
iniPath = App.Path & "\config.ini"
valor = String$(255, 0)
i = GetPrivateProfileString(CORCHETEALEER, PROPIEDADALEER, "", valor, Len(valor), iniPath)
If i > 0 Then
ReadINI = Trim(valor)
End If
End Function

Public Function LeerIni(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault) As String
    lpFileName = App.Path & "\config.ini"
    Dim lpString As String
    Dim LTmp As Long
    Dim sRetVal As String
    If IsMissing(vDefault) Then
        lpString = ""
    Else
        lpString = vDefault
    End If
    sRetVal = String$(255, 0)
    LTmp = GetPrivateProfileString(lpAppName, lpKeyName, lpString, sRetVal, Len(sRetVal), lpFileName)
    If LTmp = 0 Then
        LeerIni = lpString
    Else
        LeerIni = Left(sRetVal, LTmp)
    End If
End Function

Public Sub WriteIni(CORCHETEAESCRIBIR As Variant, PROPIEDADAESCRIBIR As String, NuevoValor As String)
iniPath = App.Path & "\config.ini"
i = WritePrivateProfileString(CORCHETEAESCRIBIR, PROPIEDADAESCRIBIR, NuevoValor, iniPath)
End Sub


