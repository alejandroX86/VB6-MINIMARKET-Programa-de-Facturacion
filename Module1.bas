Attribute VB_Name = "Module1"
'DECLARO LA CONEXION
Public DB As New Connection
Public rs As New Recordset
Public rsEmpresa As New Recordset
'PROCEDIMIENTOS ABRIR/CERRAR BASE

Public Sub AbrirBase()
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & ("\DB01.mdb") & ";" & _
     "Jet OLEDB:Database Password=admin"

End Sub

Public Sub CerrarBase()
DB.Close
Set DB = Nothing
End Sub


Public Sub SeleccionarImpresora(strTipoComp As Variant)
strSql = "SELECT * FROM IMPRESORA WHERE DESCRIPCION LIKE " & "'" & strTipoComp & "'"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
For Each xPrinter In Printers
If xPrinter.DeviceName = IIf(IsNull(rs!ruta), "", rs!ruta) Then
Set Printer = xPrinter
End If
Next
Else
MsgBox "Error de Impresora. Se usará la predeterminada.", vbExclamation
End If
rs.Close
Set rs = Nothing
Set xPrinter = Nothing
End Sub

Public Sub LimitarUsos()
AbrirBase
strSql = "Select * From Security Where ID=1"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.EOF Then
    rs.AddNew
    rs!id = 1
    End If
    
    'If Val(rs!Cantidad) < Val(60) Then
    'rs!Cantidad = rs!Cantidad + 1
    'rs.Update
    'Else
    'MsgBox "El periodo de prueba ha finalizado, por favor comuniquese con el proveedor del Sistema", vbExclamation
    'End
    'End If



CerrarBase
End Sub


