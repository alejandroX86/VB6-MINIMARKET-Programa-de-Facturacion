VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVentasPorComprob 
   BackColor       =   &H00000000&
   Caption         =   "Listado de Ventas Por Comprobante"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmVentasPorComprob.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Buscar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   765
      Left            =   6060
      TabIndex        =   11
      Top             =   0
      Width           =   6930
      Begin VB.TextBox txtBuscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         MaxLength       =   30
         TabIndex        =   14
         Top             =   270
         Width           =   2175
      End
      Begin VB.ComboBox comboCriterio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmVentasPorComprob.frx":57E2
         Left            =   3540
         List            =   "frmVentasPorComprob.frx":57EC
         TabIndex        =   13
         Text            =   "NumFact"
         Top             =   270
         Width           =   2115
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   5790
         TabIndex        =   12
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Buscar por:"
         Height          =   195
         Index           =   12
         Left            =   2700
         TabIndex        =   15
         Top             =   330
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5775
      Left            =   45
      TabIndex        =   10
      Top             =   810
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   10186
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
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
      Left            =   2205
      TabIndex        =   9
      Top             =   7185
      Width           =   1560
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BorrarTodo"
      Enabled         =   0   'False
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
      Left            =   90
      TabIndex        =   8
      Top             =   7185
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
      Left            =   7680
      TabIndex        =   7
      Top             =   7185
      Width           =   1290
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
      Left            =   4200
      TabIndex        =   6
      Top             =   7185
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
      Left            =   5880
      TabIndex        =   5
      Top             =   7185
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
      Left            =   9120
      TabIndex        =   4
      Top             =   7185
      Width           =   1200
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
      TabIndex        =   0
      Top             =   7185
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1530
      Top             =   7095
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   435
      Left            =   4320
      TabIndex        =   16
      Top             =   225
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
      Left            =   1260
      TabIndex        =   17
      Top             =   225
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
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Peridodo Desde:"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   345
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Peridodo Hasta:"
      Height          =   195
      Left            =   3090
      TabIndex        =   18
      Top             =   345
      Width           =   1275
   End
   Begin VB.Label lblGanancia 
      BackColor       =   &H0000FFFF&
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
      Height          =   255
      Left            =   8010
      TabIndex        =   3
      Top             =   6765
      Width           =   3300
   End
   Begin VB.Label lblVentas 
      BackColor       =   &H0000FFFF&
      Caption         =   "lblVentas"
      Height          =   255
      Left            =   135
      TabIndex        =   2
      Top             =   6765
      Width           =   3375
   End
   Begin VB.Label lblCosto 
      BackColor       =   &H0000FFFF&
      Caption         =   "lblCosto"
      Height          =   255
      Left            =   4140
      TabIndex        =   1
      Top             =   6765
      Width           =   3225
   End
End
Attribute VB_Name = "frmVentasPorComprob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totsub As Variant
Dim totsubNC As Variant
Dim totiva As Variant
Dim totivaNC As Variant
Dim tottotal As Variant
Dim tottotalNC As Variant


Private Sub cmdNotaCred_Click()
'imprimir nota de credito
Dim strTipoFact As String
Dim VarNumfact As Variant
Dim VarNumFactRef As Variant

Me.MSHFlexGrid1.Col = 0
VarNumFactRef = Val(Me.MSHFlexGrid1.Text)
Dim TipoFacturaSelecta As String
Me.MSHFlexGrid1.Col = 1
TipoFacturaSelecta = Me.MSHFlexGrid1.Text

If TipoFacturaSelecta = "X" Then
strTipoFact = "X"
ElseIf TipoFacturaSelecta = "A" Then
strTipoFact = "A"
ElseIf TipoFacturaSelecta = "B" Then
strTipoFact = "B"
ElseIf TipoFacturaSelecta = "C" Then
strTipoFact = "C"
End If

VarNumfact = NuevoNumeroFactura(strTipoFact)
PrintNotCred strTipoFact, VarNumfact, VarNumFactRef, ReadINI("NUMPAGE", "VALUE")
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
ImprimirFlex MSHFlexGrid1, "Listado de Ventas por Comprobante"
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
titulos = " NumFact|TipoFact|Fecha|Hora|Vendedor|Cliente|CondVenta|Subtotal|IVA|Total|Anulada|NotaCredito"

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


AbrirBase

Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"

totsub = 0
totsubNC = 0
totiva = 0
totivaNC = 0
tottotal = 0
tottotalNC = 0

MostrarFacturas "FACTURASA", strDesde, strHasta
MostrarFacturas "FACTURASB", strDesde, strHasta
MostrarFacturas "FACTURASC", strDesde, strHasta
MostrarFacturas "FACTURASX", strDesde, strHasta

lblVentas = "TOTAL SUBTOTAL: $ " & Format(totsub - totsubNC, "standard")
lblCosto = "TOTAL IVA: $ " & Format(totiva - totivaNC, "standard")
lblGanancia = "TOTAL: $ " & Format(tottotal - tottotalNC, "standard")

CerrarBase

AutoFlex MSHFlexGrid1
FlexRayado MSHFlexGrid1, &HFFFFFF, &H8000000F
End Sub



Private Sub MostrarFacturas(tipo As String, strDesde As String, strHasta As String)

strSql = "SELECT * FROM " & tipo & " WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY Fecha,Hora ASC"
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
While Not rs.EOF
                i = i + 1

                linea = rs!numfact _
                & Chr(9) & rs!tipofact _
                & Chr(9) & rs!Fecha _
                & Chr(9) & rs!Hora _
                & Chr(9) & rs!codvendedor _
                & Chr(9) & rs!CodCliente _
                & Chr(9) & rs!condventa _
                & Chr(9) & Format(rs!subtotal, "standard") _
                & Chr(9) & Format(rs!iva, "standard") _
                & Chr(9) & Format(rs!total, "standard") _
                & Chr(9) & rs!Anulada _
                & Chr(9) & rs!NotaCredito


MSHFlexGrid1.AddItem linea, i


    If rs!Anulada = 0 Then
    
        If rs!NotaCredito <> 1 Then
        totsub = totsub + rs!subtotal
        
            'If rs!tipofact = "A" Then
            totiva = totiva + rs!iva
            'End If
        tottotal = tottotal + rs!total
        
        Else
    
        totsubNC = totsubNC + rs!subtotal
        
        'If rs!tipofact = "A" Then
            totivaNC = totivaNC + rs!iva
         '   End If
        
        tottotalNC = tottotalNC + rs!total
        
        End If
    
    End If

rs.MoveNext
Wend
rs.Close
Set rs = Nothing

End Sub




'################PROCEDIMIENTO BUSCAR ####################

Private Sub cmdBuscar_Click()
AceptarBuscar txtBuscar.Text
End Sub

Private Sub txtBuscar_Change()
AceptarBuscar txtBuscar.Text
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        AceptarBuscar txtBuscar.Text
    End If
End Sub
Private Sub AceptarBuscar(ByVal txtBuscar As Variant)
titulos = " NumFact|TipoFact|Fecha|Hora|Vendedor|Cliente|CondVenta|Subtotal|IVA|Total|Anulada|NotaCredito"

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


AbrirBase

Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"

totsub = 0
totsubNC = 0
totiva = 0
totivaNC = 0
tottotal = 0
tottotalNC = 0

FiltrarFacturas "FACTURASA", strDesde, strHasta, CStr(txtBuscar)
FiltrarFacturas "FACTURASB", strDesde, strHasta, CStr(txtBuscar)
FiltrarFacturas "FACTURASC", strDesde, strHasta, CStr(txtBuscar)
FiltrarFacturas "FACTURASX", strDesde, strHasta, CStr(txtBuscar)

lblVentas = "TOTAL SUBTOTAL: $ " & Format(totsub - totsubNC, "standard")
lblCosto = "TOTAL IVA: $ " & Format(totiva - totivaNC, "standard")
lblGanancia = "TOTAL: $ " & Format(tottotal - tottotalNC, "standard")

CerrarBase

AutoFlex MSHFlexGrid1
FlexRayado MSHFlexGrid1, &HFFFFFF, &H8000000F
End Sub

Private Sub FiltrarFacturas(tipo As String, strDesde As String, strHasta As String, Criterio As String)

strSql = "SELECT * FROM " & tipo & " WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY Fecha,Hora ASC"
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText


Select Case comboCriterio.Text
Case "NumFact"
    If txtBuscar <> "" Then
    rs.Filter = "NumFact =" & Val(txtBuscar)
    End If
Case "TipoFact"
    If txtBuscar <> "" Then
    rs.Filter = "TipoFact LIKE '*" + txtBuscar + "*'"
    End If
Case "Cliente"
    If txtBuscar <> "" Then
    rs.Filter = "CodCliente LIKE '*" + txtBuscar + "*'"
    End If

Case "Talle"
    If txtBuscar <> "" Then
    rs.Filter = "Talle LIKE '*" + txtBuscar + "*'"
    End If

Case Else
rs.Filter = ""
End Select




While Not rs.EOF
                i = i + 1

                linea = rs!numfact _
                & Chr(9) & rs!tipofact _
                & Chr(9) & rs!Fecha _
                & Chr(9) & rs!Hora _
                & Chr(9) & rs!codvendedor _
                & Chr(9) & rs!CodCliente _
                & Chr(9) & rs!condventa _
                & Chr(9) & Format(rs!subtotal, "standard") _
                & Chr(9) & Format(rs!iva, "standard") _
                & Chr(9) & Format(rs!total, "standard") _
                & Chr(9) & rs!Anulada _
                & Chr(9) & rs!NotaCredito


MSHFlexGrid1.AddItem linea, i


    If rs!Anulada = 0 Then
    
        If rs!NotaCredito <> 1 Then
        totsub = totsub + rs!subtotal
        
            'If rs!tipofact = "A" Then
            totiva = totiva + rs!iva
            'End If
        tottotal = tottotal + rs!total
        
        Else
    
        totsubNC = totsubNC + rs!subtotal
        
        'If rs!tipofact = "A" Then
            totivaNC = totivaNC + rs!iva
         '   End If
        
        tottotalNC = tottotalNC + rs!total
        
        End If
    
    End If

rs.MoveNext
Wend
rs.Close
Set rs = Nothing

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
MSHFlexGrid1.Col = 1
TipoFacturaSelecta = MSHFlexGrid1.Text

If TipoFacturaSelecta = "X" Then
strSql = "Select DetalleFacturasX.CodArticulo, DetalleFacturasX.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleFacturasX INNER JOIN Articulos ON Articulos.ID = DetalleFacturasX.CodArticulo WHERE DetalleFacturasX.Numfact = " & Val(itemselecto)
ElseIf TipoFacturaSelecta = "A" Then
strSql = "Select DetalleFacturasA.CodArticulo, DetalleFacturasA.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleFacturasA INNER JOIN Articulos ON Articulos.ID = DetalleFacturasA.CodArticulo WHERE DetalleFacturasA.Numfact = " & Val(itemselecto)
ElseIf TipoFacturaSelecta = "B" Then
strSql = "Select DetalleFacturasB.CodArticulo, DetalleFacturasB.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleFacturasB INNER JOIN Articulos ON Articulos.ID = DetalleFacturasB.CodArticulo WHERE DetalleFacturasB.Numfact = " & Val(itemselecto)
ElseIf TipoFacturaSelecta = "C" Then
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
If TipoFacturaSelecta = "X" Then
v = "select * from FacturasX where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "A" Then
v = "select * from FacturasA where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "B" Then
v = "select * from FacturasB where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "C" Then
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

If TipoFacturaSelecta = "X" Then
v = "select * from detalleFacturasX where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "A" Then
v = "select * from detalleFacturasA where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "B" Then
v = "select * from detalleFacturasB where numfact=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "C" Then
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
'Dim rs2 As New Recordset
'v = "select * from CuentasCorrientes where NumFact=" & Val(itemselecto)
'rs2.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
'If Not (rs2.BOF And rs2.EOF) Then
'rs2.Delete
'rs2.Update
'End If


' ELIMINO LA VENTA DE LA UTILIDAD

If TipoFacturaSelecta = "X" Then
strSql = "SELECT * FROM Utilidad Where Descripcion like 'FACTURA X' AND ID=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "A" Then
strSql = "SELECT * FROM Utilidad Where Descripcion like 'FACTURA A' AND ID=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "B" Then
strSql = "SELECT * FROM Utilidad Where Descripcion like 'FACTURA B' AND ID=" & Val(itemselecto)
ElseIf TipoFacturaSelecta = "C" Then
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
