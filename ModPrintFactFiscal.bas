Attribute VB_Name = "ModPrintFactFiscal"
Public Sub devolverstock(tipofact, numfact)

strSql = "select * from DetalleFacturas" & tipofact & " Where Numfact=" & numfact
Dim rsdf As New Recordset
rsdf.Open strSql, DB, adOpenKeyset, adLockOptimistic

If Not rsdf.EOF Then
While Not rsdf.EOF

CodArt = rsdf!codarticulo
cantart = rsdf!Cantidad

    strSql = "SELECT * FROM ARTICULOS WHERE ID=" & CodArt
    Dim rsart As New Recordset
    rsart.Open strSql, DB, adOpenKeyset, adLockOptimistic
    If Not rsart.EOF Then
    rsart!Existencias = rsart!Existencias + cantart
    rsart.Update
    End If
    rsart.Close
    Set rsart = Nothing

rsdf.MoveNext
Wend

End If
rsdf.Close
Set rsdf = Nothing

End Sub

Public Sub borrarfactura(tipofact, numfact)
strSql = "DELETE FROM FACTURAS" & tipofact & " WHERE numfact=" & numfact
Dim rsF As New Recordset
rsF.Open strSql, DB, adOpenKeyset, adLockOptimistic
strSql = "DELETE FROM DETALLEFACTURAS" & tipofact & " WHERE numfact=" & numfact
Dim rsF2 As New Recordset
rsF2.Open strSql, DB, adOpenKeyset, adLockOptimistic
End Sub

Public Function ObtenerPrefijoCategIVA(categiva As String) As String
'resp_comprador= responsabilidad iva del comprador (M:monotributo, I:inscripto, F:Consumidor Final)
'resp_vendedor= generalmente es I de responsable inscripto... aunque no es siempre asi..
'se usa igual que resp_comprador
ObtenerPrefijoCategIVA = ""

If InStr(1, categiva, "tributo", vbTextCompare) > 0 Then
ObtenerPrefijoCategIVA = "M"
End If

If InStr(1, categiva, "final", vbTextCompare) > 0 Then
ObtenerPrefijoCategIVA = "F"
End If

If InStr(1, categiva, "cripto", vbTextCompare) > 0 Then
ObtenerPrefijoCategIVA = "I"
End If
End Function

Public Function QuitarGuionesCuit(ncuit As String) As String
QuitarGuionesCuit = Trim(Replace(ncuit, "-", ""))
End Function

Public Function ObtenerProxRemito()
Sql = "Select * From Remitos"
Dim rsRem As New Recordset
rsRem.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsRem.EOF Then
rsRem.MoveLast
tt = rsRem!numfact + 1
Else
tt = 1
End If
ObtenerProxRemito = tt
End Function

