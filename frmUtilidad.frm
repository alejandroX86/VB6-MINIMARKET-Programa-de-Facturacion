VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUtilidad 
   BackColor       =   &H00000000&
   Caption         =   "Informe De Ventas $$$"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   Icon            =   "frmUtilidad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmUtilidad.frx":57E2
      Left            =   6420
      List            =   "frmUtilidad.frx":57EF
      TabIndex        =   12
      Text            =   "Ordenes de Pedido"
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ver Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   11
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCELAR VENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7260
      TabIndex        =   9
      Top             =   4320
      Width           =   1335
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
      Height          =   375
      Left            =   8700
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmUtilidad.frx":581E
      Height          =   3075
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   5424
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   8438015
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
      Left            =   2700
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   6780
      TabIndex        =   8
      Top             =   3900
      Width           =   3075
   End
   Begin VB.Label lblVentas 
      BackColor       =   &H0000FFFF&
      Caption         =   "lblVentas"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   3900
      Width           =   3375
   End
   Begin VB.Label lblCosto 
      BackColor       =   &H0000FFFF&
      Caption         =   "lblCosto"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   3900
      Width           =   2955
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
Attribute VB_Name = "frmUtilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
VerUtilidad
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command3_Click()
MSHFlexGrid1_DblClick
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

titulos = " Número|Fecha|DESCRIPCION|TIPO VENTA|TOTAL|COSTO REP.|GANANCIA"

 MSHFlexGrid1.Clear
    With MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000
        .ColWidth(2) = 2000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With

Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"

If Combo1.Text = "Ordenes de Pedido" Then
strSql = "SELECT * FROM Utilidad WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta & "AND DESCRIPCION like 'ORDEN DE PEDIDO' ORDER BY ID ASC"
ElseIf Combo1.Text = "Facturas A" Then
strSql = "SELECT * FROM Utilidad WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta & "AND DESCRIPCION like 'FACTURA A' ORDER BY ID ASC"
ElseIf Combo1.Text = "Facturas B" Then
strSql = "SELECT * FROM Utilidad WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta & "AND DESCRIPCION like 'FACTURA B' ORDER BY ID ASC"
End If


AbrirBase
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

While Not rs.EOF
                i = i + 1
                
                linea = rs!id _
                & Chr(9) & rs!Fecha _
                & Chr(9) & rs!Descripcion _
                & Chr(9) & rs!condventa _
                & Chr(9) & Format(rs!TotalVenta, "#,###.#0") _
                & Chr(9) & Format(rs!TotalCompra, "#,###.#0") _
                & Chr(9) & Format(rs!Ganancia, "#,###.#0")

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
End Sub


Private Sub VerCodigodeCobrador()
'c2 = "Select * From tabCobradores Where Nomcob Like " & "'" & ComboVendedor.Text & "'"
'
'Dim rs2 As New Recordset
'rs2.Open c2, DB, adOpenKeyset, adLockOptimistic, adCmdText
'
'If Not (rs2.BOF And rs2.EOF) Then
'lblNroCob = rs2!NroCob
'Else
'lblNroCob = "0"
'End If
End Sub

Private Sub cmdImprimir_Click()
    'On Error GoTo eror
    
    AbrirBase
    ImprimirSeleccion
    CerrarBase
    
'Frame3.Visible = False

'eror:
'Resume Next

End Sub

Private Sub ImprimirSeleccion()
    cd1.ShowPrinter
    
'On Error Resume Next
    
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"

    
    
    Dim RstSocios As New Recordset
    RstSocios.CursorLocation = adUseClient
    
    
If Combo1.Text = "Ordenes de Pedido" Then
strSql = "SELECT * FROM Utilidad WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta & "AND DESCRIPCION like 'ORDEN DE PEDIDO' ORDER BY ID ASC"
ElseIf Combo1.Text = "Facturas A" Then
strSql = "SELECT * FROM Utilidad WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta & "AND DESCRIPCION like 'FACTURA A' ORDER BY ID ASC"
ElseIf Combo1.Text = "Facturas B" Then
strSql = "SELECT * FROM Utilidad WHERE FECHA BETWEEN " & strDesde & " AND " & strHasta & "AND DESCRIPCION like 'FACTURA B' ORDER BY ID ASC"
End If

    RstSocios.Open strSql, DB, adOpenDynamic, adLockOptimistic
    'Label1.Caption = RstSocios.RecordCount
    
    ImpTitulos
    
    ContRay = 0
    ContLin = 0
    espacioCelda = 0.4
    
    
    Do While Not RstSocios.EOF
                   
If ContRay > 0 Then
Printer.CurrentY = Printer.CurrentY + espacioCelda
End If
                     
Dim CadenaCod As String
Dim LongitudCod As Long
CadenaCod = Format(CadenaCod, "0###")
LongitudCod = Len(CadenaCod)

Printer.CurrentX = 2 + (Val(LongitudCod)) - (Printer.TextWidth(Format(RstSocios!id, "0###")))
Printer.Print Format(RstSocios!id, "0###");


       
       Printer.CurrentX = 2.8
       Printer.Print Left(RstSocios!Fecha, 22);
       
       
       
       
    Printer.CurrentX = 5
    Printer.Print Left(RstSocios!Descripcion, 22);

    Printer.CurrentX = 8
    Printer.Print Left(RstSocios!condventa, 22);


'       Printer.CurrentX = 11
'       Printer.Print Left(RstSocios!CodCliente, 22);

Dim Cadena2 As String
Dim Longitud2 As Long
Cadena2 = Format(Cadena2, "#,###.#0")
Longitud2 = Len(Cadena2)

Printer.CurrentX = 15 + (Val(Longitud2)) - (Printer.TextWidth(Format(RstSocios!TotalVenta, "#,###.#0")))
Printer.Print Format(RstSocios!TotalVenta, "#,###.#0");

Dim Cadena3 As String
Dim Longitud3 As Long
Cadena3 = Format(Cadena3, "#,###.#0")
Longitud3 = Len(Cadena3)

Printer.CurrentX = 18.5 + (Val(Longitud3)) - (Printer.TextWidth(Format(RstSocios!TotalCompra, "#,###.#0")))
Printer.Print Format(RstSocios!TotalCompra, "#,###.#0");


Dim Cadena4 As String
Dim Longitud4 As Long
Cadena4 = Format(Cadena4, "#,###.#0")
Longitud4 = Len(Cadena4)

Printer.CurrentX = 20 + (Val(Longitud4)) - (Printer.TextWidth(Format(RstSocios!Ganancia, "#,###.#0")))
Printer.Print Format(RstSocios!Ganancia, "#,###.#0");

TotVent = TotVent + RstSocios!TotalVenta
totComp = totComp + RstSocios!TotalCompra
totUt = totUt + RstSocios!Ganancia

       ContRay = ContRay + 1
       ContLin = ContLin + 1
       
       If Printer.CurrentY > 28 Then
          Printer.NewPage
          ImpTitulos
          ContLin = 0
          ContRay = 0
       End If
       
       If ContRay = 5 Then
       Printer.DrawStyle = 2
           Printer.CurrentY = Printer.CurrentY + espacioCelda
           Printer.Line (1, Printer.CurrentY)-(20, Printer.CurrentY)
           ContRay = 0
       End If
       RstSocios.MoveNext
   Loop
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.CurrentX = 2
    Printer.Print "TOTAL VENTAS: $ " & Format(TotVent, "#,###.#0");
    Printer.CurrentX = 6
    Printer.Print "COSTO MERCADERIA: $ " & Format(totComp, "#,###.#0");
    Printer.CurrentX = 14
    Printer.Print "GANANCIA PERIODO: $ " & Format(totUt, "#,###.#0")
    
    Printer.EndDoc
    'cmdsalir_click

End Sub



Private Sub ImpTitulos()

Titulo = "Informe de Ventas"


If Combo1.Text = "Ordenes de Pedido" Then
Titulo = "Informe de Ventas"
ElseIf Combo1.Text = "Facturas A" Then
Titulo = "Informe de Facturas A"
ElseIf Combo1.Text = "Facturas B" Then
Titulo = "Informe de Facturas B"
End If



    'Imprime el Template de la hoja
Printer.ScaleMode = vbCentimeters
Printer.PaperSize = 9
Printer.Orientation = 1
Printer.FontName = "ARIAL"
Printer.DrawStyle = 0
Printer.Font.Size = 14

'Printer.CurrentX = 1
'Printer.CurrentY = 1
'Printer.Print "DYVCOM S.R.L."
Printer.PaintPicture Picture1.Picture, 1, 1

Printer.Font.Size = 8

Printer.CurrentX = 18.8
Printer.CurrentY = 1
Printer.Print "Página " & Printer.Page


Printer.CurrentX = 18.5
Printer.CurrentY = 1.5
Printer.Print Date '& " " & Format(Time, "hh:nn am/pm")

Printer.Font.Size = 14



'Dim Cadena3 As String
'Dim Longitud3 As Long

'Cadena3 = Format(Titulo, String(Len(Titulo), "_"))
'Longitud3 = Len(Cadena3)

'Printer.CurrentX = 8 + (Val(Longitud3)) - (Printer.TextWidth(Len(Titulo) ))


x = 21 / 2
Printer.CurrentX = x - (Printer.TextWidth(Titulo)) + (Printer.TextWidth(Titulo)) / 2
Printer.CurrentY = 3
Printer.Print Titulo

Printer.Font.Size = 8


' TITULO DE LOS CAMPOS

Printer.CurrentY = 4


Printer.CurrentX = 1
Printer.Print "Numero";

Printer.CurrentX = 2.8
Printer.Print "Fecha";

Printer.CurrentX = 5
Printer.Print "Descripcion";

Printer.CurrentX = 8
Printer.Print "Cond. Venta";

'Printer.CurrentX = 11
'Printer.Print "Cliente";

Printer.CurrentX = 14
Printer.Print "TotalVenta";

Printer.CurrentX = 17.5
Printer.Print "Costo.M.";

Printer.CurrentX = 19.2
Printer.Print "Utilidad"
    
Printer.Line (1, (Printer.CurrentY + espacioCelda))-(20, (Printer.CurrentY + espacioCelda))

End Sub




Private Sub MSHFlexGrid1_DblClick()

MSHFlexGrid1.Col = 0
frmGananciaDetalle.txtID = Val(MSHFlexGrid1.Text)
frmGananciaDetalle.Show


End Sub
Private Sub Command2_Click()
MSHFlexGrid1.Col = 0
If IsNumeric(MSHFlexGrid1.Text) Then
If MsgBox("¿Anula este Item?", vbExclamation + vbYesNoCancel) = vbYes Then

EliminarItem

End If
End If

End Sub


Private Sub EliminarItem()
Dim rsCarrito As New Recordset
Dim itemselecto As Integer
Dim strSql As String
AbrirBase


MSHFlexGrid1.Col = 0
itemselecto = MSHFlexGrid1.Text




Dim rsActualizacionStock As New Recordset
'Dim strSql As String

MSHFlexGrid1.Col = 2
cobselecto = MSHFlexGrid1.Text




If Combo1.Text = "Ordenes de Pedido" Then
strSql = "Select DetalleVentas.CodArticulo, DetalleVentas.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleVentas INNER JOIN Articulos ON Articulos.ID = DetalleVentas.CodArticulo WHERE DetalleVentas.Numfact = " & Val(itemselecto)
ElseIf Combo1.Text = "Facturas A" Then
strSql = "Select DetalleFacturasA.CodArticulo, DetalleFacturasA.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleFacturasA INNER JOIN Articulos ON Articulos.ID = DetalleFacturasA.CodArticulo WHERE DetalleFacturasA.Numfact = " & Val(itemselecto)
ElseIf Combo1.Text = "Facturas B" Then
strSql = "Select DetalleFacturasB.CodArticulo, DetalleFacturasB.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleFacturasB INNER JOIN Articulos ON Articulos.ID = DetalleFacturasB.CodArticulo WHERE DetalleFacturasB.Numfact = " & Val(itemselecto)
End If


'strSql = "Select DetalleVentas.CodArticulo, DetalleVentas.Cantidad AS CANTVENT,Articulos.ID, Articulos.Existencias AS CANTSTOCK FROM DetalleVentas INNER JOIN Articulos ON Articulos.ID = DetalleVentas.CodArticulo WHERE DetalleVentas.Numfact = " & Val(itemselecto)
rsActualizacionStock.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsActualizacionStock.EOF Then
While (Not rsActualizacionStock.EOF)
rsActualizacionStock!CANTSTOCK = Val(rsActualizacionStock!CANTSTOCK) + Val(rsActualizacionStock!CANTVENT)
rsActualizacionStock.MoveNext
Wend
End If



If Combo1.Text = "Ordenes de Pedido" Then
v = "select * from ventas where numfact=" & Val(itemselecto)
ElseIf Combo1.Text = "Facturas A" Then
v = "select * from FacturasA where numfact=" & Val(itemselecto)
ElseIf Combo1.Text = "Facturas B" Then
v = "select * from FacturasB where numfact=" & Val(itemselecto)
End If
'v = "select * from ventas where numfact=" & Val(itemselecto)
rs.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rs.BOF And rs.EOF) Then
'While Not rs.EOF
rs!condventa = "ANULADO"
'rs.Delete
rs.Update
'Wend
End If

'Dim rs1 As New Recordset
'
'If Combo1.Text = "Ordenes de Pedido" Then
'v = "select * from detalleventas where numfact=" & Val(itemselecto)
'ElseIf Combo1.Text = "Facturas A" Then
'v = "select * from detalleFacturasA where numfact=" & Val(itemselecto)
'ElseIf Combo1.Text = "Facturas B" Then
'v = "select * from detalleFacturasB where numfact=" & Val(itemselecto)
'End If
'
''v = "select * from detalleventas where numfact=" & Val(itemselecto)
'rs1.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
'If Not (rs1.BOF And rs1.EOF) Then
'While Not rs1.EOF
'rs1.Delete
'rs1.MoveNext
'Wend
'End If

Dim rs2 As New Recordset
v = "select * from CuentasCorrientes where NumFact=" & Val(itemselecto)
rs2.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rs2.BOF And rs2.EOF) Then
rs2.Delete
rs2.Update
End If

'Dim rs2 As New Recordset
'v = "select * from ConsignacionesACobrar where NumFact=" & Val(itemselecto)
'rs2.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
'If Not (rs2.BOF And rs2.EOF) Then
'rs2.Delete
'rs2.Update
'End If



If Combo1.Text = "Ordenes de Pedido" Then
strSql = "SELECT * FROM Utilidad Where Descripcion like 'ORDEN DE PEDIDO' AND ID=" & Val(itemselecto)
ElseIf Combo1.Text = "Facturas A" Then
strSql = "SELECT * FROM Utilidad Where Descripcion like 'FACTURA A' AND ID=" & Val(itemselecto)
ElseIf Combo1.Text = "Facturas B" Then
strSql = "SELECT * FROM Utilidad Where Descripcion like 'FACTURA B' AND ID=" & Val(itemselecto)
End If
'strSql = "SELECT * FROM Utilidad Where Descripcion like ID=" & Val(itemselecto)
rsCarrito.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rsCarrito.BOF And rsCarrito.EOF) Then
rsCarrito.MoveLast
rsCarrito.Delete
rsCarrito.Update
End If
CerrarBase
VerUtilidad
End Sub
