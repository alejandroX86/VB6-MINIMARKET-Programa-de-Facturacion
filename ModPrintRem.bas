Attribute VB_Name = "ModPrintRem"
Public Sub PrintRemito(ComboTipoFact, nrem, ComboCopias)
AbrirBase
SeleccionarImpresora "Remitos"

strSql = "Select * From Remitos WHERE NumFact=" & Val(nrem)
Dim rsFacturas As New Recordset
rsFacturas.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsFacturas.EOF Then

    sql2 = "Select * From Clientes Where ID=" & Val(rsFacturas!CodCliente)
    Dim rsClientes As New Recordset
    rsClientes.Open sql2, DB, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsClientes.EOF Then
    
        For i = 1 To ComboCopias
        ImprimirTitulos rsFacturas!tipofact, rsFacturas!Numfact, rsFacturas!Fecha
        ImprimirDatosCliente rsClientes!id, rsClientes!nombre, rsClientes!categiva, rsClientes!cuit, rsClientes!domicilio, rsClientes!telefono
        ImprimirCondicionesDeVenta rsFacturas!condventa, rsFacturas!codvendedor, 0
        ImprimirTituloCampos
        ImprimirDetalle
        ImprimirTotales rsFacturas!tipofact, rsFacturas!subtotal, rsFacturas!iva, rsFacturas!total
        Printer.NewPage
        Next i
        Printer.EndDoc
   
    End If
    rsClientes.Close


End If
rsFacturas.Close
CerrarBase
End Sub

Private Sub ImprimirTitulos(inVarTipoFact, InVarNumFact, InVarFecha As Variant)
'###########ENCABEZADO########################
'Tipo de Papel A4
'Printer.PaperSize = 9
'Printer.Height = 20
'Printer.Width = 22
'Impresion en blanco y negro
'Printer.ColorMode = 1
Printer.ScaleMode = vbCentimeters
'Printer.ScaleHeight = 20.1
'Printer.ScaleWidth = 22.1
'BORDES HORIZONTALES
'Printer.Line (0, 0)-(20, 0)
'Printer.Line (0, 22)-(20, 22)
'BORDES VERTICALES
'Printer.Line (0, 0)-(0, 22)
'Printer.Line (20, 0)-(20, 22)
Printus "Numero Remito", inVarTipoFact & " " & Format(InVarNumFact, "0001-########"), 0, 0, 0
Printus "Fecha Remito", InVarFecha & " - " & Format(Time, "hh:nn am/pm"), 0, 0, 0
End Sub

Private Sub ImprimirDatosCliente(id, nombre, categiva, cuit, domicilio, telefono)
Printus "Codigo Cliente", " (" & id & ")", 0, 0, 0
Printus "Nombre Cliente", nombre, 0, 0, 0
Printus "Domicilio Cliente", domicilio, 0, 0, 0
Printus "Telefono Cliente", telefono, 0, 0, 0
Printus "CategIva Cliente", categiva, 0, 0, 0
Printus "Cuit Cliente", cuit, 0, 0, 0
End Sub

Private Sub ImprimirCondicionesDeVenta(condventa, codvendedor, InBoolPrintRem)
Printus "Condiciones de Venta", condventa, 0, 0, 0
Printus "Codigo Vendedor", codvendedor, 0, 0, 0
'If InBoolPrintRem = 1 Then
'Printer.CurrentX = 8
'Printer.Print "" & ObtenerProxNRemito
'End If
End Sub

Private Function ObtenerProxRemito()
Sql = "Select * From Remitos"
Dim rsRem As New Recordset
rsRem.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsRem.EOF Then
rsRem.MoveLast
tt = rsRem!Numfact + 1
Else
tt = 1
End If
ObtenerProxRemito = tt
End Function

Private Sub ImprimirTituloCampos()
End Sub

Private Sub ImprimirDetalle()
x = "Select * From Carrito"
Dim rsPrintItem As New Recordset
rsPrintItem.Open (x), DB, adOpenKeyset, adLockOptimistic, adCmdText
While (Not rsPrintItem.EOF)
R = R + 0.5
Printus "Cantidad Articulo", Format(rsPrintItem!Cantidad, "0###"), 0, 1, R
Printus "Codigo Articulo", Format(rsPrintItem!codarticulo, "0###"), 0, 1, R
Printus "Descripcion Articulo", rsPrintItem!Descripcion, 0, 1, R
Printus "Precio Unitario Articulo", rsPrintItem!P_Unitario, 1, 1, R
Printus "Precio Neto Articulo", rsPrintItem!P_NETO, 1, 1, R
rsPrintItem.MoveNext
Wend
rsPrintItem.Close
End Sub
Private Sub ImprimirTotales(tipofact, subtotal, iva, total)
If tipofact = "A" Then
Printus "Subtotal Remito", subtotal, 1, 0, 0
Printus "Iva Remito", iva, 1, 0, 0
End If
Printus "Total Remito", total, 1, 0, 0
End Sub

Private Sub Printus(tipo, cadena, ismoney, isitem, paramot)
strSql = "Select * From CordRem Where Type like " & "'" & tipo & "'"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then

    If rs!Visible = 1 Then
    Printer.FontName = rs!Font
    Printer.Font.Size = rs!Size
        
        If isitem = 1 Then
        Printer.CurrentY = rs!CordY + paramot
        Else
        Printer.CurrentY = rs!CordY
        End If
        
        If ismoney = 1 Then
        Printer.CurrentX = rs!CordX - Printer.TextWidth(Format(cadena, "#,###.#0"))
        Printer.Print "" & Format(cadena, "#,###.#0")
        Else
        Printer.CurrentX = rs!CordX
        Printer.Print "" & cadena
        End If
    
   
    End If
End If
rs.Close
End Sub

