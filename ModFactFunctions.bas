Attribute VB_Name = "ModFactFunctions"
Public Sub VaciarCarrito()
Dim strSql As String
AbrirBase
strSql = "DELETE FROM Carrito"
Dim rsCarrito As New Recordset
rsCarrito.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
'CalcularTotales "" & txtSubTotal, "" & txtIva, "" & txtTotal
CerrarBase
End Sub
Public Function NuevoNumeroFactura(tipo As Variant) As Variant
strSql = "SELECT * FROM Facturas" & tipo
AbrirBase
Dim rsFacturas As New Recordset
rsFacturas.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
        If Not (rsFacturas.BOF And rsFacturas.EOF) Then
        rsFacturas.MoveLast
        NuevoNumeroFactura = Format((Val(rsFacturas!NumFact) + 1), "0#######")
        Else
        NuevoNumeroFactura = Format(1, "0#######")
        End If
rsFacturas.Close
CerrarBase
End Function

Public Function ObtenerUsuario(txtID, txtName, ComboCondVenta) As Variant
Dim rsUsuarios As New Recordset
Dim strSql As String
AbrirBase
strSql = "SELECT * FROM Vendedores Where ID=" & Val(txtID)
rsUsuarios.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rsUsuarios.BOF And rsUsuarios.EOF) Then
txtName.Text = "" & rsUsuarios!nombre
txtID.Enabled = False
txtName.Enabled = True
ComboCondVenta.SetFocus
Else
MsgBox "Código de vendedor incorrecto"
End If
CerrarBase
End Function
Public Sub ObtenerCliente(txtCodCliente As TextBox, txtNombre As TextBox, txtDomicilio As TextBox, txtTelefono As TextBox, ComboCategIva As ComboBox, txtCuit As TextBox)
Dim strSql As String
AbrirBase
strSql = "SELECT * FROM Clientes WHERE ID=" & Val(txtCodCliente)
Dim rsClientes As New Recordset
rsClientes.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
' Si existe
If Not (rsClientes.BOF And rsClientes.EOF) Then
txtNombre = "" & rsClientes!nombre
txtDomicilio = "" & rsClientes!domicilio
txtTelefono = "" & rsClientes!telefono
ComboCategIva = "" & rsClientes!categiva
txtCuit = "" & rsClientes!cuit
'Else
'MsgBox "Codigo Inexistente. Agrega Nuevo Cliente!", vbExclamation
End If
' Deshabilito el campo clave
txtCodCliente.Enabled = False
' Habilito todos los demás campos
txtNombre.Enabled = True
txtDomicilio.Enabled = True
txtTelefono.Enabled = True
ComboCategIva.Enabled = True
txtCuit.Enabled = True
txtNombre.SetFocus

CerrarBase
End Sub


Public Sub GuardarCliente(txtCodCliente As TextBox, txtNombre As TextBox, txtDomicilio As TextBox, txtTelefono As TextBox, ComboCategIva As ComboBox, txtCuit As TextBox)
Dim strSql As String
Dim rsClientes As New Recordset
AbrirBase
strSql = "SELECT * FROM Clientes WHERE ID=" & Val(txtCodCliente)
rsClientes.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
        
If (rsClientes.BOF And rsClientes.EOF) Then
rsClientes.AddNew
End If
rsClientes!id = "" & txtCodCliente
rsClientes!nombre = "" & txtNombre
rsClientes!domicilio = "" & txtDomicilio
rsClientes!telefono = "" & txtTelefono
rsClientes!categiva = "" & ComboCategIva
rsClientes!cuit = "" & txtCuit
rsClientes.Update

CerrarBase

End Sub



Public Sub GetProduct(txtCodArticulo As TextBox, txtDescripcion As TextBox, txtPrecio As TextBox, txtCantidad As TextBox)
Dim rsPreciosProv As New Recordset
Dim strSql As String
AbrirBase
strSql = "SELECT * FROM PreciosClientes " & _
"WHERE CodProveedor =" & Val(txtCodCliente) & _
   " AND CodArticulo =" & Val(txtCodArticulo)

rsPreciosProv.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (rsPreciosProv.BOF And rsPreciosProv.EOF) Then
    txtDescripcion.Text = "" & rsPreciosProv!Descripcion & " " & txtMarca & " " & txtTalle
    txtPrecio.Text = "" & Replace(Format(rsPreciosProv!Precio, "fixed"), ",", ".")
    Else
    GetProductGen txtCodArticulo, txtDescripcion, txtPrecio, txtCantidad
    End If

CerrarBase
End Sub

Public Sub GetProductGen(txtCodArticulo As TextBox, txtDescripcion As TextBox, txtPrecio As TextBox, txtCantidad As TextBox)
Dim strSql As String
strSql = "SELECT * FROM Articulos " & _
"WHERE ID =" & Val(txtCodArticulo)
Dim rsArticulos As New Recordset
rsArticulos.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
    If Not (rsArticulos.BOF And rsArticulos.EOF) Then
    ' Si existe
    txtDescripcion.Text = "" & rsArticulos!Descripcion & " " & txtMarca & " " & txtTalle
    txtPrecio.Text = "" & Replace(Format(rsArticulos!Precio, "fixed"), ",", ".")
    txtCodArticulo.Enabled = False
    txtCantidad.Enabled = True
    txtPrecio.Enabled = True
    txtCantidad.SetFocus
    Else
    MsgBox "Número Incorrecto"
    End If
    
End Sub



Public Sub AgregarProducto(ComboTipoFact As ComboBox, txtCodArticulo As TextBox, txtDescripcion As TextBox, txtPrecio As TextBox, txtCantidad As TextBox)

Dim rsVal As New Recordset
Dim rsCar As New Recordset

AbrirBase
'averiguamos si hay suficiente stock
strCarrito = "SELECT * FROM CARRITO WHERE CODARTICULO=" & Val(txtCodArticulo)
rsCar.Open strCarrito, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsCar.EOF Then
num = Val(rsCar!Cantidad)
End If
num2 = Val(txtCantidad) + Val(num)
strSql = "Select * From Articulos Where ID=" & Val(txtCodArticulo)
rsVal.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rsVal.EOF Then

    If rsVal!Existencias < Val(num2) Then
    MsgBox "ERROR: Verifique Stock de Mercaderías", vbCritical
    Else
    'hay suficiente, agregamos el item
    AgregarItem ComboTipoFact, txtCodArticulo, txtDescripcion, txtPrecio, txtCantidad
    End If
End If
CerrarBase

End Sub

Public Sub AgregarItem(ComboTipoFact As ComboBox, txtCodArticulo As TextBox, txtDescripcion As TextBox, txtPrecio As TextBox, txtCantidad As TextBox)
Dim strSql As String
Dim rsCarrito As New Recordset
strSql = "Select * from Carrito"
rsCarrito.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If rsCarrito.RecordCount <= 10 Then
rsCarrito.AddNew
rsCarrito!codarticulo = Format(Val(txtCodArticulo), "0000")
rsCarrito!Descripcion = txtDescripcion
rsCarrito!Cantidad = Format(Val(txtCantidad), "0000")

varimp = ObtenerImpuestoProducto(Val(txtCodArticulo))

If ComboTipoFact.Text = "A" Or ComboTipoFact.Text = "X" Then
rsCarrito!P_Unitario = Format(Val(txtPrecio) / varimp, "standard")
rsCarrito!P_NETO = Format((Val(txtPrecio) * Val(txtCantidad)) / varimp, "standard")
Else
rsCarrito!P_Unitario = Format(Val(txtPrecio), "standard")
rsCarrito!P_NETO = Format(Val(txtPrecio) * Val(txtCantidad), "standard")
End If


rsCarrito.Update

txtCodArticulo.Enabled = True
txtMarca = ""
txtTalle = ""
'lblCodInt = ""
txtCodArticulo = ""
txtDescripcion = ""
txtPrecio = ""
txtCantidad = ""
txtCodArticulo.SetFocus

Else
MsgBox "Solo hasta 10 articulos por comprobante", vbExclamation
End If

End Sub

Public Sub CalcularTotales(ComboTipoFact As ComboBox, txtSubTotal As TextBox, txtIva As TextBox, txtTotal As TextBox)

Dim R As Double
Dim strSql As String






acumsiniva = 0
acumIva = 0

tipoFactura = "" & ComboTipoFact.Text
strSql = "SELECT * FROM Carrito"
Dim rsTotales As New Recordset
AbrirBase
rsTotales.Open ("Carrito"), DB, adOpenKeyset, adLockOptimistic, adCmdTable
While Not rsTotales.EOF

varimp = ObtenerImpuestoProducto(Val(rsTotales!codarticulo))
precioprod = ObtenerPrecioProducto(Val(rsTotales!codarticulo))
cant = Val(rsTotales!Cantidad)

'primero si es factura b no se va a sumar el neto del carrito
    If tipoFactura = "A" Then
    
    preciosiniva = precioprod / varimp
    totiva = precioprod - preciosiniva
    
    
    Else
    totconImp = ObtenerPrecioProducto(Val(rsTotales!codarticulo)) * varimp
    totiva = totconImp - ObtenerPrecioProducto(Val(rsTotales!codarticulo))
    End If
    


acumsiniva = acumsiniva + (preciosiniva * cant)
acumIva = acumIva + (totiva * cant)


IVI = IVI + totiva
R = R + rsTotales!P_NETO

rsTotales.MoveNext
Wend
rsTotales.Close
CerrarBase


If ComboTipoFact.Text = "A" Then
    txtSubTotal = acumsiniva
    txtIva = acumIva
    txtTotal = acumsiniva + acumIva
    Else
    txtTotal = R
    txtIva = IVI
    txtSubTotal = R - IVI
End If

End Sub


Public Function ObtenerImpuestoProducto(codarticulo As Variant) As String

strSql = "Select * From cnTaxProduct WHERE IDARTICULO=" & Val(codarticulo)
Dim rsx As New Recordset
rsx.Open strSql, DB, adOpenKeyset, adLockOptimistic

If Not rsx.EOF Then

strSql = "Select * FROM Impuestos WHERE ID=" & Val(rsx!IDIMPUESTO)
Dim rsy As New Recordset
rsy.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rsy.EOF Then
ObtenerImpuestoProducto = "1," & Replace(rsy!total, ",", "")
Else
ObtenerImpuestoProducto = "1,21"
'MsgBox "El Articulo " & codarticulo & " No tiene Alicuota Asignada, se asignó un 21.00%", vbExclamation
End If

Else
ObtenerImpuestoProducto = "1,21"
MsgBox "El Articulo " & codarticulo & " No tiene Alicuota Asignada, se asignó un 21.00%", vbExclamation
End If
rsx.Close

End Function

Public Function ObtenerPrecioProducto(codarticulo) As Double

strSql = "Select * From Articulos WHERE ID=" & Val(codarticulo)
Dim rsx As New Recordset
rsx.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rsx.EOF Then

'If rsx!Descripcion <> "ENVIO" Then
ObtenerPrecioProducto = rsx!Precio
'Else
'ObtenerPrecioProducto = 0
'End If

Else
ObtenerPrecioProducto = 0
MsgBox ObtenerPrecioProducto
End If
rsx.Close
End Function



Public Sub EliminarItem(MSHFlexGrid1 As MSHFlexGrid)
Dim itemselecto As Integer
Dim strSql As String
AbrirBase
MSHFlexGrid1.Col = 0
itemselecto = Val(MSHFlexGrid1.Text)
strSql = "SELECT * FROM Carrito Where CodArticulo=" & Val(itemselecto)
Dim rsCarrito As New Recordset
rsCarrito.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rsCarrito.BOF And rsCarrito.EOF) Then
rsCarrito.MoveLast
rsCarrito.Delete
rsCarrito.Update
End If
CerrarBase
End Sub

Public Sub AplicarRecargoGeneral(txtPorcentaje As TextBox)
strSql = "Select * From Carrito"
AbrirBase
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
Dim valorPU As Double
Dim valorPN As Double
If Not rs.EOF Then
    While Not rs.EOF
    valorPU = 0
    valorPN = 0
    valorPU = valor + rs!P_Unitario + (rs!P_Unitario * Format(Val(txtPorcentaje), "#,###.#0") / 100)
    valorPN = valor + rs!P_NETO + (rs!P_NETO * Format(Val(txtPorcentaje), "#,###.#0") / 100)
    rs!P_Unitario = valorPU
    rs!P_NETO = valorPN
    rs.MoveNext
    Wend
    
End If
CerrarBase
End Sub

Public Sub AplicarRecargoIndividual(MSHFlexGrid1 As MSHFlexGrid, txtPorcentaje As TextBox)

MSHFlexGrid1.Col = 0
strSql = "Select * From Carrito WHERE CodArticulo=" & Val(MSHFlexGrid1.Text)
AbrirBase
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
Dim valorPU As Double
Dim valorPN As Double
If Not rs.EOF Then
    valorPU = 0
    valorPN = 0
    valorPU = valor + rs!P_Unitario + (rs!P_Unitario * Format(Val(txtPorcentaje), "#,###.#0") / 100)
    valorPN = valor + rs!P_NETO + (rs!P_NETO * Format(Val(txtPorcentaje), "#,###.#0") / 100)
    rs!P_Unitario = valorPU
    rs!P_NETO = valorPN
    rs.MoveNext
End If
CerrarBase

End Sub






Public Function GuardarDatosDeFactura(ComboTipoFact, txtCodUsuario, txtCodCliente, ComboCondVenta, dtpVto, txtSubTotal, txtIva, txtTotal) As Variant
Dim n As Integer
Dim v As String
Dim strSql As String

Select Case ComboCondVenta
Case "Cta./Cte. 30 días"
v = Date + 30
Case "Cta./Cte. 60 días"
v = Date + 60
Case "Cta./Cte. 90 días"
v = Date + 90
Case Else
v = Date
End Select
        
strSql = "SELECT * FROM Facturas" & ComboTipoFact
Dim rsFacturas As New Recordset
rsFacturas.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
      
            If Not (rsFacturas.BOF And rsFacturas.EOF) Then
            rsFacturas.MoveLast
            n = Val(rsFacturas!NumFact) + 1
            Else
            n = 1
            End If
            
            rsFacturas.AddNew
            rsFacturas!NumFact = n
            rsFacturas!tipofact = "" & ComboTipoFact
            rsFacturas!Fecha = Date
            rsFacturas!Hora = Time
            rsFacturas!codvendedor = Val(txtCodUsuario)
            rsFacturas!CodCliente = "" & txtCodCliente
            rsFacturas!condventa = "" & ComboCondVenta
            rsFacturas!Vencimiento = "" & dtpVto
            rsFacturas!subtotal = "" & txtSubTotal
            rsFacturas!iva = "" & txtIva
            rsFacturas!total = "" & txtTotal
            rsFacturas.Update
  
GuardarDetalleDeFactura ComboTipoFact, n
GuardarDatosDeFactura = n
End Function
Public Sub GuardarDetalleDeFactura(ComboTipoFact, n)
Dim rsCarrito As New Recordset
Dim rsDetalleFacturas As New Recordset
Dim strSql As String
Dim x As String
Dim cdf As Long

Dim subtoX As Double


x = "SELECT * FROM Carrito"
strSql = "SELECT * FROM DetalleFacturas" & ComboTipoFact
rsCarrito.Open (x), DB, adOpenKeyset, adLockOptimistic, adCmdText
rsDetalleFacturas.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not (rsDetalleFacturas.BOF And rsDetalleFacturas.EOF) Then
While (Not rsCarrito.EOF)
rsDetalleFacturas.MoveLast
cdf = rsDetalleFacturas!CodDetalleFactura
rsDetalleFacturas.AddNew
rsDetalleFacturas!CodDetalleFactura = cdf + 1
rsDetalleFacturas!NumFact = n
rsDetalleFacturas!codarticulo = rsCarrito!codarticulo
rsDetalleFacturas!Cantidad = rsCarrito!Cantidad
rsDetalleFacturas!P_Unitario = rsCarrito!P_Unitario
rsDetalleFacturas!P_NETO = rsCarrito!P_NETO

varimp = 0
varimp = ObtenerImpuestoProducto(Val(rsCarrito!codarticulo))
subtoX = 0
subtoX = rsCarrito!P_NETO / varimp

If ComboTipoFact = "A" Then
rsDetalleFacturas!iva = rsCarrito!P_NETO * varimp
Else
rsDetalleFacturas!iva = rsCarrito!P_NETO - subtoX
End If

rsDetalleFacturas.MoveNext
rsCarrito.MoveNext
Wend
Else
While (Not rsCarrito.EOF)
cdf = cdf + 1
rsDetalleFacturas.AddNew
rsDetalleFacturas!CodDetalleFactura = cdf
rsDetalleFacturas!NumFact = n
rsDetalleFacturas!codarticulo = rsCarrito!codarticulo
rsDetalleFacturas!Cantidad = rsCarrito!Cantidad
rsDetalleFacturas!P_Unitario = rsCarrito!P_Unitario
rsDetalleFacturas!P_NETO = rsCarrito!P_NETO
varimp = 0
varimp = ObtenerImpuestoProducto(Val(rsCarrito!codarticulo))
subtoX = 0
subtoX = rsCarrito!P_NETO / varimp

If ComboTipoFact = "A" Then
rsDetalleFacturas!iva = rsCarrito!P_NETO * varimp
Else
rsDetalleFacturas!iva = rsCarrito!P_NETO - subtoX
End If

rsDetalleFacturas.MoveNext
rsCarrito.MoveNext
Wend
End If

End Sub

Public Function GuardarDatosDeRemito(ComboTipoFact, txtCodUsuario, txtCodCliente, ComboCondVenta, dtpVto, txtSubTotal, txtIva, txtTotal) As Variant
Dim n As Integer
Dim v As String
Dim strSql As String

Select Case ComboCondVenta
Case "Cta./Cte. 30 días"
v = Date + 30
Case "Cta./Cte. 60 días"
v = Date + 60
Case "Cta./Cte. 90 días"
v = Date + 90
Case Else
v = Date
End Select
        
strSql = "SELECT * FROM Remitos"
Dim rsRemitos As New Recordset
rsRemitos.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
      
            If Not (rsRemitos.BOF And rsRemitos.EOF) Then
            rsRemitos.MoveLast
            n = Val(rsRemitos!NumFact) + 1
            Else
            n = 1
            End If
            
            rsRemitos.AddNew
            rsRemitos!NumFact = n
            rsRemitos!tipofact = "R"
            rsRemitos!Fecha = Date
            rsRemitos!Hora = Time
            rsRemitos!codvendedor = Val(txtCodUsuario)
            rsRemitos!CodCliente = txtCodCliente
            rsRemitos!condventa = ComboCondVenta
            rsRemitos!Vencimiento = dtpVto
            rsRemitos!subtotal = txtSubTotal
            rsRemitos!iva = txtIva
            rsRemitos!total = txtTotal
            rsRemitos.Update
            
GuardarDetalleDeRemito n
GuardarDatosDeRemito = n
End Function


Public Sub GuardarDetalleDeRemito(n)
Dim rsCarrito As New Recordset
Dim rsDetalleRemitos As New Recordset

Dim strSql As String
Dim x As String
Dim cdf As Long

x = "SELECT * FROM Carrito"
strSql = "SELECT * FROM DetalleRemitos"
rsCarrito.Open (x), DB, adOpenKeyset, adLockOptimistic, adCmdText
rsDetalleRemitos.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not (rsDetalleRemitos.BOF And rsDetalleRemitos.EOF) Then
While (Not rsCarrito.EOF)
rsDetalleRemitos.MoveLast
cdf = rsDetalleRemitos!CodDetalleFactura
rsDetalleRemitos.AddNew
rsDetalleRemitos!CodDetalleFactura = cdf + 1
rsDetalleRemitos!NumFact = n
rsDetalleRemitos!codarticulo = rsCarrito!codarticulo
rsDetalleRemitos!Cantidad = rsCarrito!Cantidad
rsDetalleRemitos!P_Unitario = rsCarrito!P_Unitario
rsDetalleRemitos!P_NETO = rsCarrito!P_NETO
rsDetalleRemitos.MoveNext
rsCarrito.MoveNext
Wend
Else
While (Not rsCarrito.EOF)
cdf = cdf + 1
rsDetalleRemitos.AddNew
rsDetalleRemitos!CodDetalleFactura = cdf
rsDetalleRemitos!NumFact = n
rsDetalleRemitos!codarticulo = rsCarrito!codarticulo
rsDetalleRemitos!Cantidad = rsCarrito!Cantidad
rsDetalleRemitos!P_Unitario = rsCarrito!P_Unitario
rsDetalleRemitos!P_NETO = rsCarrito!P_NETO
rsDetalleRemitos.MoveNext
rsCarrito.MoveNext
Wend
End If

End Sub

Public Sub GuardarUtilidad(txtNumFact, ComboTipoFact, ComboCondVenta)
Dim rsUtil As New Recordset
Dim rsU2 As New Recordset
strSql = "Select Carrito.CodArticulo AS CodArt,Carrito.Descripcion AS DESCR,Carrito.Cantidad AS CANT,Carrito.P_UNITARIO AS UNIT, Carrito.P_NETO AS NETO, Articulos.ID,Articulos.PrecioProv AS PREPROV From Carrito INNER JOIN Articulos ON Carrito.CodArticulo=Articulos.ID"
rsU2.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

While Not rsU2.EOF
TotVent = TotVent + Val(Str(rsU2!Neto))
TotCompr = TotCompr + rsU2!PREPROV * rsU2!cant
TotGan = TotGan + rsU2!Neto - rsU2!PREPROV * rsU2!cant
rsU2.MoveNext
Wend
strUtil = "Select * From Utilidad"
rsUtil.Open strUtil, DB, adOpenKeyset, adLockOptimistic, adCmdText
rsUtil.AddNew
rsUtil!id = txtNumFact
rsUtil!Fecha = Date
rsUtil!Hora = Time
rsUtil!Descripcion = "FACTURA " & ComboTipoFact
rsUtil!condventa = ComboCondVenta
rsUtil!TotalCompra = Format(TotCompr, "#,###.#0")
rsUtil!TotalVenta = Format(TotVent, "#,###.#0")
rsUtil!Ganancia = Format(TotGan, "#,###.#0")
rsUtil.Update
End Sub

Public Sub ActualizarStock()
Dim rsActualizacionStock As New Recordset
Dim strSql As String
strSql = "Select Carrito.CodArticulo, Carrito.Cantidad,Articulos.ID, Articulos.Existencias FROM Carrito INNER JOIN Articulos ON Articulos.ID = Carrito.CodArticulo"
rsActualizacionStock.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
While (Not rsActualizacionStock.EOF)
rsActualizacionStock!Existencias = Val(rsActualizacionStock!Existencias) - Val(rsActualizacionStock!Cantidad)
rsActualizacionStock.MoveNext
Wend
End Sub
Public Sub ActualizarCuentasCorrientes(ComboCondVenta, txtCodCliente, txtNumFact, txtTotal)

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
    rsCta!Descripcion = "Venta en Cuenta Corriente"
    rsCta!Debe = txtTotal
    rsCta.Update

End If
End Sub


