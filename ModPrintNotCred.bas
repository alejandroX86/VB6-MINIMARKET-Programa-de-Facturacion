Attribute VB_Name = "ModPrintNotCred"
'Dim R As Variant
Public Sub PrintNotCred(ComboTipoFact, nfact, nfactRef, ComboCopias)
AbrirBase

ReplicarFactura ComboTipoFact, nfact, nfactRef

SeleccionarImpresora "Facturas" & ComboTipoFact


strSql = "Select * From Facturas" & ComboTipoFact & " WHERE NumFact=" & Val(nfact)
Dim rsFacturas As New Recordset
rsFacturas.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rsFacturas.EOF Then

    sql2 = "Select * From Clientes Where ID=" & Val(rsFacturas!CodCliente)
    Dim rsClientes As New Recordset
    rsClientes.Open sql2, DB, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsClientes.EOF Then
    
        For i = 1 To ComboCopias
        ImprimirDatosEmpresa
        ImprimirTitulos rsFacturas!tipofact, rsFacturas!NumFact, rsFacturas!Fecha
        ImprimirLeyenda
        ImprimirDatosCliente rsClientes!id, rsClientes!nombre, rsClientes!categiva, rsClientes!cuit, rsClientes!domicilio, rsClientes!telefono
        ImprimirCondicionesDeVenta rsFacturas!condventa, rsFacturas!codvendedor, 0
        ImprimirTituloCampos
        ImprimirDetalle
        
        postotal = Printer.CurrentX
        
        ImprimirTotales rsFacturas!tipofact, rsFacturas!subtotal, rsFacturas!iva, rsFacturas!total, postotal
        Printer.NewPage
        Next i
        Printer.EndDoc
   
    End If
    rsClientes.Close


End If
rsFacturas.Close
CerrarBase
End Sub

Public Sub ReplicarFactura(ComboTipoFact, nfact, nfactRef)

Dim rs As New Recordset
strSql = "Select * FROM FACTURAS" & ComboTipoFact & " WHERE Numfact=" & nfactRef
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    strSql = "Select * FROM FACTURAS" & ComboTipoFact
    Dim rs1 As New Recordset
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic
    rs1.AddNew
    rs1!NumFact = nfact
    rs1!tipofact = rs!tipofact
    rs1!Fecha = rs!Fecha
    rs1!Hora = rs!Hora
    rs1!CodCliente = rs!CodCliente
    rs1!condventa = rs!condventa
    rs1!Vencimiento = rs!Vencimiento
    rs1!subtotal = rs!subtotal
    rs1!iva = rs!iva
    rs1!total = rs!total
    rs1!Anulada = rs!Anulada
    rs1!NotaCredito = 1
    rs1.Update
    rs1.Close
    
    GuardarDetalleNotaCredito ComboTipoFact, nfact, nfactRef
    
    RellenarCarrito ComboTipoFact, nfact, nfactRef
  
    AcreditarEnCtaCte rs!condventa, rs!CodCliente, nfact, rs!total
    ResetearUtilidad nfact, ComboTipoFact
    
End If

rs.Close

End Sub

Private Sub ResetearUtilidad(nfact, ComboTipoFact)
Dim tipof As String
tipof = "FACTURA" & ComboTipoFact

strSql = "DELETE FROM UTILIDAD WHERE DESCRIPCION LIKE " & "'" & tipof & "'" & " AND ID=" & Val(nfact)
Dim rsU As New Recordset
rsU.Open strSql, DB, adOpenKeyset, adLockOptimistic

End Sub


Private Sub RellenarCarrito(ComboTipoFact, nfact, nfactRef)

Dim rsDcart As New Recordset
strSql = "DELETE FROM CARRITO"
rsDcart.Open strSql, DB, adOpenKeyset, adLockOptimistic


Dim rsDF As New Recordset
strSql = "Select * FROM DETALLEFACTURAS" & ComboTipoFact & " WHERE Numfact=" & nfactRef
rsDF.Open strSql, DB, adOpenKeyset, adLockOptimistic

If Not rsDF.EOF Then

    While Not rsDF.EOF
    
        Dim rsCarrito As New Recordset
        strSql = "Select * FROM CARRITO"
        rsCarrito.Open strSql, DB, adOpenKeyset, adLockOptimistic
        
        rsCarrito.AddNew
        rsCarrito!codarticulo = rsDF!codarticulo
        rsCarrito!Descripcion = ObtenerDescripcionArticulo(rsDF!codarticulo)
        rsCarrito!Cantidad = rsDF!Cantidad
        rsCarrito!P_Unitario = rsDF!P_Unitario
        rsCarrito!P_NETO = rsDF!P_NETO
        rsCarrito.Update
        rsCarrito.Close
        Set rsCarrito = Nothing

        
    rsDF.MoveNext
    Wend

End If
rsDF.Close
Set rsDF = Nothing


End Sub

Private Function ObtenerDescripcionArticulo(codarticulo As Variant) As String

Dim rsD As New Recordset
strSql = "SELECT * FROM Articulos WHERE ID=" & Val(codarticulo)
rsD.Open strSql, DB, adOpenKeyset, adLockOptimistic

If Not rsD.EOF Then
ObtenerDescripcionArticulo = "" & rsD!Descripcion
Else
ObtenerDescripcionArticulo = ""
End If

End Function



Private Sub GuardarDetalleNotaCredito(ComboTipoFact, nfact, nfactRef)

Dim rs3 As New Recordset
strSql = "Select * FROM DETALLEFACTURAS" & ComboTipoFact & " WHERE Numfact=" & nfactRef
rs3.Open strSql, DB, adOpenKeyset, adLockOptimistic

If Not rs3.EOF Then

    While Not rs3.EOF
    
        Dim rs4 As New Recordset
        strSql = "Select * FROM DETALLEFACTURAS" & ComboTipoFact
        rs4.Open strSql, DB, adOpenKeyset, adLockOptimistic
        
        If Not rs4.EOF Then
        rs4.MoveLast
        cont = rs4!CodDetalleFactura + 1
        Else
        cont = 1
        End If
        
        rs4.AddNew
        rs4!CodDetalleFactura = cont
        rs4!NumFact = nfact
        rs4!codarticulo = rs3!codarticulo
        rs4!Cantidad = rs3!Cantidad
        rs4!P_Unitario = rs3!P_Unitario
        rs4!P_NETO = rs3!P_NETO
        rs4.Update
        rs4.Close
        Set rs4 = Nothing

    
    DevolverStock rs3!codarticulo, rs3!Cantidad
    
    rs3.MoveNext
    Wend

End If
rs3.Close
Set rs3 = Nothing

End Sub



Public Sub AcreditarEnCtaCte(ComboCondVenta, txtCodCliente, txtNumFact, txtTotal)

If ComboCondVenta = "Cuenta Corriente" Then

    Dim rsCta As New Recordset
    strSql = "Select * From CuentasCorrientes"
    rsCta.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
    Dim nx As Long
    
        If rsCta.EOF Then
        rsCta.AddNew
        rsCta!id = "1"
        Else
        rsCta.MoveLast
        nx = Val(rsCta!id)
        rsCta.AddNew
        rsCta!id = nx + 1
        End If
    
    rsCta!CodCliente = Val(txtCodCliente)
    rsCta!NumFact = txtNumFact
    rsCta!Fecha = Date
    rsCta!Descripcion = "Nota de Crédito"
    rsCta!Haber = txtTotal
    rsCta.Update

End If
End Sub



Private Sub DevolverStock(CodArt, Cantidad)
strSql = "Select * From Articulos WHERE ID=" & CodArt
Dim rs5 As New Recordset
rs5.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rs5.EOF Then
rs5!Existencias = rs5!Existencias + Cantidad
rs5.Update
End If
rs5.Close
Set rs5 = Nothing
End Sub

Private Sub ImprimirDatosEmpresa()
Dim emp As String
emp = "SELECT * FROM EMPRESA WHERE IdEmpresa = 1"
'AbrirBase
Dim rsEmpresa As New Recordset
rsEmpresa.Open (emp), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rsEmpresa.BOF And rsEmpresa.EOF) Then
Printus "Nombre Empresa", rsEmpresa!nombre, 0, 0, 0
Printus "Direccion Empresa", rsEmpresa!Direccion, 0, 0, 0
End If
rsEmpresa.Close
'CerrarBase
End Sub

Private Sub ImprimirLeyenda()
Printus "COMANDA DE USO INTERNO", "COMANDA DE USO INTERNO", 0, 0, 0
Printus "DOCUMENTO NO VALIDO COMO FACTURA", "DOCUMENTO NO VALIDO COMO FACTURA", 0, 0, 0
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
Printus "Nota Credito", "Nota de Crédito", 0, 0, 0
Printus "Numero Factura", "Numero: " & inVarTipoFact & " " & Format(InVarNumFact, "0001-########"), 0, 0, 0
Printus "Fecha Factura", "Fecha: " & InVarFecha & " - " & Format(Time, "hh:nn am/pm"), 0, 0, 0
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
Printus "Codigo Vendedor", "Mostrador " & codvendedor, 0, 0, 0
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
tt = rsRem!NumFact + 1
Else
tt = 1
End If
ObtenerProxRemito = tt
End Function

Private Sub ImprimirTituloCampos()
Printus "COLUMNA CANTIDAD", "Cant.", 0, 0, 0
Printus "COLUMNA DESCRIPCION", "DESCRIPCION", 0, 0, 0
Printus "COLUMNA PU", "P.U", 0, 0, 0
Printus "COLUMNA PN", "TOT", 0, 0, 0

End Sub

Private Sub ImprimirDetalle()
x = "Select * From Carrito"
Dim rsPrintItem As New Recordset
rsPrintItem.Open (x), DB, adOpenKeyset, adLockOptimistic, adCmdText
While (Not rsPrintItem.EOF)
R = R + 0.4
Printus "Cantidad Articulo", Format(rsPrintItem!Cantidad, "0###"), 0, 1, R
Printus "Codigo Articulo", Format(rsPrintItem!codarticulo, "0###"), 0, 1, R
Printus "Descripcion Articulo", rsPrintItem!Descripcion, 0, 1, R
Printus "Precio Unitario Articulo", rsPrintItem!P_Unitario, 1, 1, R
Printus "Precio Neto Articulo", rsPrintItem!P_NETO, 1, 1, R
rsPrintItem.MoveNext
Wend
rsPrintItem.Close
End Sub
Private Sub ImprimirTotales(tipofact, subtotal, iva, total, postotal)
If tipofact = "A" Then
Printus "Subtotal Factura", subtotal, 1, 0, 0
Printus "Iva Factura", iva, 1, 0, 0
End If
Printus "Total Factura", total, 1, 0, 0


End Sub

Private Sub Printus(tipo, cadena, ismoney, isitem, paramot)
strSql = "Select * From CordFact Where Type like " & "'" & tipo & "'"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then
    
    If rs!Visible = 1 Then
        
        Printer.FontName = rs!Font
        Printer.Font.Size = rs!Size
        Printer.CurrentX = rs!CordX
     
        If isitem = 1 Then
        Printer.CurrentY = rs!CordY + paramot
        Else
        
            If tipo = "Total Factura" Then
            Printer.CurrentY = Printer.CurrentY + 0.5
            Else
            Printer.CurrentY = rs!CordY
            End If
            
        End If
        
        If ismoney = 1 Then
            
            If tipo = "Total Factura" Then
            Printer.CurrentX = Printer.CurrentX - Printer.TextWidth("TOTAL $" & Format(cadena, "standard"))
            Printer.Print "TOTAL $" & Format(cadena, "standard")
            Else
            Printer.CurrentX = Printer.CurrentX - Printer.TextWidth(Format(cadena, "standard"))
            Printer.Print "" & Format(cadena, "standard")
            End If
            
        Else
        
        Printer.Print "" & cadena
        End If
    
   
    End If
End If
rs.Close
End Sub
