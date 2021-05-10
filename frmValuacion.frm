VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmValuacion 
   BackColor       =   &H0080FFFF&
   Caption         =   "Valuación de STOCK"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmValuacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2400
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir Listado de Reposición"
      Height          =   495
      Left            =   180
      TabIndex        =   9
      Top             =   5280
      Width           =   1635
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ver Archivo..."
      Height          =   495
      Left            =   2100
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Archivar..."
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   6180
      TabIndex        =   6
      Top             =   5280
      Width           =   1515
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Listado de Reposición"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   60
      Width           =   2535
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Valuacion de Stock de Venta"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   60
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Valuacion de Stock de Compra"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Value           =   -1  'True
      Width           =   2595
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmValuacion.frx":57E2
      Height          =   4095
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   7223
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
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   4740
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4020
      TabIndex        =   3
      Top             =   4740
      Width           =   3615
   End
End
Attribute VB_Name = "frmValuacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
frmArchivoValuacion.Show
End Sub

Private Sub Option1_Click()
AplicarOpcion
End Sub
Private Sub Option2_Click()
AplicarOpcion
End Sub
Private Sub Option3_Click()
AplicarOpcion
End Sub
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmArchivoValuacion.Show
End Sub

Private Sub Form_Load()
AplicarOpcion
End Sub

Private Sub AplicarOpcion()
If Option1.Value = True Then
VerValuacionCompra
ElseIf Option2.Value = True Then
VerValuacionVenta
ElseIf Option3.Value = True Then
VerListadoReposición

End If



End Sub


Private Sub VerValuacionCompra()

titulos = "Codigo|Descripcion|Cantidad|Precio Compra|Valuacion"

MSHFlexGrid1.Clear
    With MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .ColWidth(0) = 700
        .ColWidth(1) = 2500
        .ColWidth(2) = 1000
        .ColWidth(3) = 1300
        .ColWidth(4) = 1300
        '.ColWidth(5) = 1200
        '.ColWidth(6) = 1200
        '.ColWidth(7) = 1200
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With

Dim totStock As Double
Dim rsStock As New Recordset
st = "Select * From Articulos ORDER BY ID"

AbrirBase
rsStock.Open (st), DB, adOpenKeyset, adLockOptimistic, adCmdText
'
While Not rsStock.EOF


                i = i + 1
                
                linea = Format(rsStock!id, "0###") _
                & Chr(9) & rsStock!Descripcion & " " & rsStock!Marca _
                & Chr(9) & rsStock!Existencias _
                & Chr(9) & Format(rsStock!PrecioProv, "standard") _
                & Chr(9) & Format(rsStock!Existencias * rsStock!PrecioProv, "standard") _

MSHFlexGrid1.AddItem linea, i

totStock = totStock + rsStock!PrecioProv * rsStock!Existencias
rsStock.MoveNext
Wend
Label1 = "Valuación de Stock= $ " & Format(totStock, "#,###.#0")
Label2 = "Total De Registros= " & rsStock.RecordCount
CerrarBase
End Sub



Private Sub VerValuacionVenta()

titulos = "Codigo|Descripcion|Cantidad|Precio Venta|Valuacion"

MSHFlexGrid1.Clear
    With MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .ColWidth(0) = 700
        .ColWidth(1) = 2500
        .ColWidth(2) = 1000
        .ColWidth(3) = 1300
        .ColWidth(4) = 1300
        '.ColWidth(5) = 1200
        '.ColWidth(6) = 1200
        '.ColWidth(7) = 1200
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With

Dim totStock As Double
Dim rsStock As New Recordset
st = "Select * From Articulos ORDER BY ID"

AbrirBase
rsStock.Open (st), DB, adOpenKeyset, adLockOptimistic, adCmdText
'
While Not rsStock.EOF


                i = i + 1
                
                linea = Format(rsStock!id, "0000") _
                & Chr(9) & rsStock!Descripcion & " " & rsStock!Marca _
                & Chr(9) & rsStock!Existencias _
                & Chr(9) & Format(rsStock!Precio, "standard") _
                & Chr(9) & Format(rsStock!Existencias * rsStock!Precio, "Fixed") _

MSHFlexGrid1.AddItem linea, i

totStock = totStock + rsStock!Precio * rsStock!Existencias
rsStock.MoveNext
Wend
Label1 = "Valuación de Stock= $" & Format(totStock, "#,###.#0")
Label2 = "Total De Registros= " & rsStock.RecordCount
CerrarBase
AutoFlex Me.MSHFlexGrid1

End Sub

Private Sub VerListadoReposición()
titulos = "Codigo|Descripcion|Cantidad|StockMinimo|Diferencia"

MSHFlexGrid1.Clear
    With MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .ColWidth(0) = 700
        .ColWidth(1) = 2500
        .ColWidth(2) = 1000
        .ColWidth(3) = 1300
        .ColWidth(4) = 1300
        '.ColWidth(5) = 1200
        '.ColWidth(6) = 1200
        '.ColWidth(7) = 1200
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With

Dim totStock As Double
Dim rsStock As New Recordset
st = "Select * From Articulos WHERE Existencias<StockMinimo"

AbrirBase
rsStock.Open (st), DB, adOpenKeyset, adLockOptimistic, adCmdText
'
While Not rsStock.EOF


                i = i + 1
                
                linea = Format(rsStock!id, "0000") _
                & Chr(9) & rsStock!Descripcion & " " & rsStock!Marca _
                & Chr(9) & rsStock!Existencias _
                & Chr(9) & rsStock!StockMinimo _
                & Chr(9) & rsStock!Existencias - rsStock!StockMinimo _

MSHFlexGrid1.AddItem linea, i

totStock = totStock + rsStock!StockMinimo - rsStock!Existencias
rsStock.MoveNext
Wend
Label1 = "Diferencia = " & totStock
Label2 = "Total De Registros = " & rsStock.RecordCount
CerrarBase

AutoFlex Me.MSHFlexGrid1


End Sub

















Private Sub cmdImprimir_Click()
cd1.ShowPrinter
    'On Error GoTo eror
    
    AbrirBase
    ImprimirSeleccion
    CerrarBase
    
'Frame3.Visible = False

'eror:
'Resume Next

End Sub

Private Sub ImprimirSeleccion()
    
    
'On Error Resume Next
    

    
    
Dim RstSocios As New Recordset
RstSocios.CursorLocation = adUseClient
    
    strSql = "Select * From Articulos WHERE Existencias<StockMinimo"

    RstSocios.Open strSql, DB, adOpenDynamic, adLockOptimistic
    Label1.Caption = RstSocios.RecordCount
    
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
       Printer.Print RstSocios!Descripcion & " " & RstSocios!Marca;
       
       
       
       
'    Printer.CurrentX = 6
'       Printer.Print Left(RstSocios!Existencias, 22);

       
'       Printer.CurrentX = 11
'       Printer.Print Left(RstSocios!StockMinimo, 22);

Dim Cadena2 As String
Dim Longitud2 As Long
Cadena2 = Format(Cadena2, "Standard")
Longitud2 = Len(Cadena2)

Printer.CurrentX = 15 + (Val(Longitud2)) - (Printer.TextWidth(Format(RstSocios!Existencias, "Standard")))
Printer.Print Format(RstSocios!Existencias, "Standard");

Dim Cadena3 As String
Dim Longitud3 As Long
Cadena3 = Format(Cadena3, "Standard")
Longitud3 = Len(Cadena3)

Printer.CurrentX = 17.5 + (Val(Longitud3)) - (Printer.TextWidth(Format(RstSocios!StockMinimo, "Standard")))
Printer.Print Format(RstSocios!StockMinimo, "Standard");


Dim Cadena4 As String
Dim Longitud4 As Long
Cadena4 = Format(Cadena4, "Standard")
Longitud4 = Len(Cadena4)

Printer.CurrentX = 20 + (Val(Longitud4)) - (Printer.TextWidth(Format(RstSocios!Existencias - RstSocios!StockMinimo, "Standard")))
Printer.Print Format(RstSocios!Existencias - RstSocios!StockMinimo, "Standard");

'TotVent = TotVent + RstSocios!TotalVenta
'totComp = totComp + RstSocios!TotalCompra
'totUt = totUt + RstSocios!Ganancia

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
'    Printer.CurrentY = Printer.CurrentY + 1
'    Printer.CurrentX = 2
'    Printer.Print "TOTAL VENTAS: $ " & Format(TotVent, "#,###.#0");
'    Printer.CurrentX = 6
'    Printer.Print "COSTO MERCADERIA: $ " & Format(totComp, "#,###.#0");
'    Printer.CurrentX = 14
'    Printer.Print "GANANCIA PERIODO: $ " & Format(totUt, "#,###.#0")
    
    Printer.EndDoc
    'cmdsalir_click

End Sub



Private Sub ImpTitulos()

Titulo = "Listado De Reposición"


    'Imprime el Template de la hoja
Printer.ScaleMode = vbCentimeters
Printer.PaperSize = 9
Printer.Orientation = 1
Printer.FontName = "ARIAL"
Printer.DrawStyle = 0
Printer.Font.Size = 14

'Printer.CurrentX = 1
'Printer.CurrentY = 1
'Printer.PaintPicture frmLogo.Picture1.Picture, 1, 1
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
Printer.CurrentY = 1.5
Printer.Print Titulo

Printer.Font.Size = 8


' TITULO DE LOS CAMPOS

Printer.CurrentY = 2.5


Printer.CurrentX = 1
Printer.Print "CodArticulo";

Printer.CurrentX = 2.8
Printer.Print "Descripcion";

'Printer.CurrentX = 6
'Printer.Print "Existencias";

'Printer.CurrentX = 11
'Printer.Print "Stock Minimo";

Printer.CurrentX = 14
Printer.Print "Existencias";

Printer.CurrentX = 16.5
Printer.Print "StockMinimo";

Printer.CurrentX = 19
Printer.Print "Diferencia"
    
Printer.Line (1, (Printer.CurrentY + espacioCelda))-(20, (Printer.CurrentY + espacioCelda))

End Sub

