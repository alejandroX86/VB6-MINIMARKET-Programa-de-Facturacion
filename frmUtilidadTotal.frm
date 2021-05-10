VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUtilidadTotal 
   BackColor       =   &H00000000&
   Caption         =   "Listado de Ventas Por Comprobante"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmUtilidadTotal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNotaCred 
      Caption         =   "Nota de Credito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1170
      TabIndex        =   14
      Top             =   7545
      Width           =   1560
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BorrarTodo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   7545
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exportar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7875
      TabIndex        =   12
      Top             =   7545
      Width           =   1245
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ver Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4365
      TabIndex        =   11
      Top             =   7545
      Width           =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar Venta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6075
      TabIndex        =   10
      Top             =   7545
      Width           =   1605
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   9
      Top             =   7545
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   435
      Left            =   4560
      TabIndex        =   3
      Top             =   180
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   141164545
      CurrentDate     =   38284
   End
   Begin MSComCtl2.DTPicker dtp1 
      Height          =   435
      Left            =   1500
      TabIndex        =   2
      Top             =   180
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   141164545
      CurrentDate     =   38284
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CERRAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   1
      Top             =   7545
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmUtilidadTotal.frx":57E2
      Height          =   6705
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   11827
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   65280
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6390
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblGanancia 
      BackColor       =   &H00000000&
      Caption         =   "lblGanancia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8010
      TabIndex        =   8
      Top             =   7125
      Width           =   3300
   End
   Begin VB.Label lblVentas 
      BackColor       =   &H00000000&
      Caption         =   "lblVentas"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   135
      TabIndex        =   7
      Top             =   7125
      Width           =   3375
   End
   Begin VB.Label lblCosto 
      BackColor       =   &H00000000&
      Caption         =   "lblCosto"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4140
      TabIndex        =   6
      Top             =   7125
      Width           =   3225
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Peridodo Hasta:"
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   300
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Peridodo Desde:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   300
      Width           =   1275
   End
End
Attribute VB_Name = "frmUtilidadTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdNotaCred_Click()
'imprimir nota de credito


End Sub

'close
Private Sub Command1_Click()
Unload Me
End Sub

'ver detalle
Private Sub Command3_Click()
MSHFlexGrid1_DblClick
End Sub

'export to excel
Private Sub Command4_Click()
FlexGrid_To_Excel Me.MSHFlexGrid1, MSHFlexGrid1.Rows, MSHFlexGrid1.Cols, "Ventas por Comprobante"
End Sub

'imprimir
Private Sub cmdImprimir_Click()
'ImprimirFlex MSHFlexGrid1, "Listado de Ventas por Comprobante"
End Sub

'Borrar Todo
Private Sub Command5_Click()
PreguntarBorrarRegistros
VerUtilidad
End Sub

Private Sub PreguntarBorrarRegistros()

If MsgBox("Atención: ¿Borrar Registros de Ventas?", vbYesNo + vbExclamation) = vbYes Then

Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"

AbrirBase
strSql = "Delete FROM Utilidad WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

BorrarVenta "FACTURASA"
BorrarVenta "FACTURASB"
BorrarVenta "FACTURASC"
BorrarVenta "FACTURASX"
CerrarBase
MsgBox "Los registros han sido borrados", vbInformation
End If
End Sub

Private Sub BorrarVenta(tipofact As String)
Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"


strSql = "Select * FROM " & tipofact & " WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

While Not rs.EOF

Dim rs1 As New Recordset
strSql = "SELECT * FROM DETALLE" & tipofact & " WHERE Numfact =" & rs!numfact
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic
While Not rs1.EOF
rs1.Delete
rs1.MoveNext
Wend
rs1.Close

rs.Delete
rs.MoveNext
Wend
rs.Close

End Sub

Private Sub dtp1_Change()
VerUtilidad
End Sub
Private Sub dtp2_Change()
VerUtilidad
End Sub

Private Sub Form_Load()
dtp1 = Date
dtp2 = Date
VerUtilidad
End Sub

Private Sub VerUtilidad()

titulos = " Número|Fecha|Hora|DESCRIPCION|TIPO VENTA|TOTAL|COSTO REP.|GANANCIA"

    With MSHFlexGrid1
        .Clear
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With

Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"


strSql = "SELECT * FROM Utilidad WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY Fecha,Hora ASC"

AbrirBase
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

While Not rs.EOF
                i = i + 1
                
                linea = rs!id _
                & Chr(9) & rs!Fecha _
                & Chr(9) & rs!Hora _
                & Chr(9) & rs!Descripcion _
                & Chr(9) & rs!condventa _
                & Chr(9) & Format(rs!TotalVenta, "standard") _
                & Chr(9) & Format(rs!TotalCompra, "standard") _
                & Chr(9) & Format(rs!Ganancia, "standard")

MSHFlexGrid1.AddItem linea, i

If rs!condventa <> "ANULADO" Then
TotVent = TotVent + rs!TotalVenta
totComp = totComp + rs!TotalCompra
totUt = totUt + rs!Ganancia
End If
rs.MoveNext
Wend

lblVentas = "TOTAL VENTAS: $ " & Format(TotVent, "#,###.#0")
lblCosto = "COSTO MERCADERIA: $ " & Format(totComp, "#,###.#0")
lblGanancia = "GANANCIA PERIODO: $ " & Format(totUt, "#,###.#0")

CerrarBase
AutoFlex Me.MSHFlexGrid1

End Sub

'ver detalle
Private Sub MSHFlexGrid1_DblClick()
MSHFlexGrid1.Col = 0
frmGananciaDetalle.txtID = Val(MSHFlexGrid1.Text)
frmGananciaDetalle.Show
End Sub
Private Sub Command2_Click()
'anular venta
MSHFlexGrid1.Col = 0
If IsNumeric(MSHFlexGrid1.Text) Then
    If MsgBox("¿Anula este Item?", vbExclamation + vbYesNoCancel) = vbYes Then
    EliminarItem
    End If
End If
End Sub


Private Sub EliminarItem()
Dim itemselecto As Integer
Dim strSql As String
AbrirBase

' DEVUELVO EL STOCK
MSHFlexGrid1.Col = 0
itemselecto = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 2
TipoFacturaSelecta = MSHFlexGrid1.Text

If TipoFacturaSelecta = "FACTURA X" Then
strSql = "Select DetalleFacturasX.CodArticulo, DetalleFacturasX.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleFacturasX INNER JOIN Articulos ON Articulos.ID = DetalleFacturasX.CodArticulo WHERE DetalleFacturasX.Numfact = " & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA A" Then
strSql = "Select DetalleFacturasA.CodArticulo, DetalleFacturasA.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleFacturasA INNER JOIN Articulos ON Articulos.ID = DetalleFacturasA.CodArticulo WHERE DetalleFacturasA.Numfact = " & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA B" Then
strSql = "Select DetalleFacturasB.CodArticulo, DetalleFacturasB.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleFacturasB INNER JOIN Articulos ON Articulos.ID = DetalleFacturasB.CodArticulo WHERE DetalleFacturasB.Numfact = " & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA C" Then
strSql = "Select DetalleFacturasC.CodArticulo, DetalleFacturasC.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleFacturasC INNER JOIN Articulos ON Articulos.ID = DetalleFacturasC.CodArticulo WHERE DetalleFacturasC.Numfact = " & Val(itemselecto)
End If

Dim rsActualizacionStock As New Recordset
rsActualizacionStock.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsActualizacionStock.EOF Then
While (Not rsActualizacionStock.EOF)
rsActualizacionStock!CANTSTOCK = Val(rsActualizacionStock!CANTSTOCK) + Val(rsActualizacionStock!CANTVENT)
rsActualizacionStock.MoveNext
Wend
End If



' NO BORRO SOLO MARCO LA VENTA COMO ANULADA
If TipoFacturaSelecta = "FACTURA X" Then
v = "select * from FacturasX where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA A" Then
v = "select * from FacturasA where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA B" Then
v = "select * from FacturasB where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA C" Then
v = "select * from FacturasC where numfact=" & Val(itemselecto)
End If

rs.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rs.BOF And rs.EOF) Then
'While Not rs.EOF
rs!condventa = "ANULADO"
rs!Anulada = 1
'rs.Delete
rs.Update
'Wend
End If


' ELIMINO DETALLES DE LA VENTA

If TipoFacturaSelecta = "FACTURA X" Then
v = "select * from detalleFacturasX where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA A" Then
v = "select * from detalleFacturasA where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA B" Then
v = "select * from detalleFacturasB where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA C" Then
v = "select * from detalleFacturasC where numfact=" & Val(itemselecto)
End If

Dim rs1 As New Recordset
rs1.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rs1.BOF And rs1.EOF) Then
While Not rs1.EOF
rs1.Delete
rs1.MoveNext
Wend
End If


' BORRO LA VENTA SI FUE EN CUENTA CORRIENTE
Dim rs2 As New Recordset
v = "select * from CuentasCorrientes where NumFact=" & Val(itemselecto)
rs2.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rs2.BOF And rs2.EOF) Then
rs2.Delete
rs2.Update
End If


' ELIMINO LA VENTA DE LA UTILIDAD

If TipoFacturaSelecta = "FACTURA X" Then
strSql = "SELECT * FROM Utilidad Where Descripcion like 'FACTURA X' AND ID=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA A" Then
strSql = "SELECT * FROM Utilidad Where Descripcion like 'FACTURA A' AND ID=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA B" Then
strSql = "SELECT * FROM Utilidad Where Descripcion like 'FACTURA B' AND ID=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "FACTURA C" Then
strSql = "SELECT * FROM Utilidad Where Descripcion like 'FACTURA C' AND ID=" & Val(itemselecto)
End If

Dim rsCarrito As New Recordset
rsCarrito.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rsCarrito.BOF And rsCarrito.EOF) Then
rsCarrito.MoveLast
rsCarrito.Delete
rsCarrito.Update
End If
CerrarBase
VerUtilidad
End Sub
