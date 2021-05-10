VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGananciaDetalle 
   BackColor       =   &H00000000&
   Caption         =   "Ver Detalle de Utilidad..."
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   Icon            =   "frmGananciaDetalle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   2700
      TabIndex        =   6
      Text            =   "txtDescripcion"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Text            =   "txtID"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   8040
      TabIndex        =   2
      Top             =   4020
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmGananciaDetalle.frx":57E2
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   6059
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   16761024
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
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label4"
      Height          =   195
      Left            =   2700
      TabIndex        =   4
      Top             =   3660
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label3"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3660
      Width           =   2475
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   195
      Left            =   5580
      TabIndex        =   1
      Top             =   3660
      Width           =   3075
   End
End
Attribute VB_Name = "frmGananciaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Actualizar
End Sub

Private Sub Form_Load()
frmVentasPorComprob.MSHFlexGrid1.Col = 0
txtID = Val(frmVentasPorComprob.MSHFlexGrid1.Text)
frmVentasPorComprob.MSHFlexGrid1.Col = 1
txtDescripcion = frmVentasPorComprob.MSHFlexGrid1.Text

Actualizar
End Sub

Private Sub Actualizar()

Dim TipoFacturaSelecta As String
frmVentasPorComprob.MSHFlexGrid1.Col = 1
TipoFacturaSelecta = frmVentasPorComprob.MSHFlexGrid1.Text

If TipoFacturaSelecta = "X" Then
strSql = "Select DetalleFacturasX.CodArticulo AS CodArt,Articulos.Descripcion AS DESCR,DetalleFacturasX.Cantidad AS CANT,DetalleFacturasX.P_UNITARIO AS UNIT, DetalleFacturasX.P_NETO AS NETO, Articulos.ID, Articulos.PrecioProv as PREPROV From DetalleFacturasX INNER JOIN Articulos ON DetalleFacturasX.CodArticulo=Articulos.ID WHERE DetalleFacturasX.NumFact=" & Val(txtID)
ElseIf TipoFacturaSelecta = "A" Then
strSql = "Select DetalleFacturasA.CodArticulo AS CodArt,Articulos.Descripcion AS DESCR,DetalleFacturasA.Cantidad AS CANT,DetalleFacturasA.P_UNITARIO AS UNIT, DetalleFacturasA.P_NETO AS NETO, Articulos.ID, Articulos.PrecioProv as PREPROV From DetalleFacturasA INNER JOIN Articulos ON DetalleFacturasA.CodArticulo=Articulos.ID WHERE DetalleFacturasA.NumFact=" & Val(txtID)
ElseIf TipoFacturaSelecta = "B" Then
strSql = "Select DetalleFacturasB.CodArticulo AS CodArt,Articulos.Descripcion AS DESCR,DetalleFacturasB.Cantidad AS CANT,DetalleFacturasB.P_UNITARIO AS UNIT, DetalleFacturasB.P_NETO AS NETO, Articulos.ID, Articulos.PrecioProv as PREPROV  From DetalleFacturasB INNER JOIN Articulos ON DetalleFacturasB.CodArticulo=Articulos.ID WHERE DetalleFacturasB.NumFact=" & Val(txtID)
ElseIf TipoFacturaSelecta = "C" Then
strSql = "Select DetalleFacturasC.CodArticulo AS CodArt,Articulos.Descripcion AS DESCR,DetalleFacturasC.Cantidad AS CANT,DetalleFacturasC.P_UNITARIO AS UNIT, DetalleFacturasC.P_NETO AS NETO, Articulos.ID, Articulos.PrecioProv as PREPROV   From DetalleFacturasC INNER JOIN Articulos ON DetalleFacturasC.CodArticulo=Articulos.ID WHERE DetalleFacturasC.NumFact=" & Val(txtID)
End If

titulos = " COD.|DESCRIPCION |CANT.|P.UNITARIO|P. NETO|P. PROV|GAN. NETA"

 MSHFlexGrid1.Clear
    With MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .ColWidth(0) = 600
        .ColWidth(1) = 2500
        .ColWidth(2) = 600
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200

        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With

AbrirBase

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

While Not rs.EOF
                i = i + 1
                
                linea = Format(rs!CodArt, "0000") _
                & Chr(9) & rs!DESCR _
                & Chr(9) & Format(rs!cant, "0000") _
                & Chr(9) & Format(rs!UNIT, "#,###.#0") _
                & Chr(9) & Format(rs!Neto, "#,###.#0") _
                & Chr(9) & Format(rs!PREPROV * rs!cant, "#,###.#0") _
                & Chr(9) & Format(rs!Neto - rs!PREPROV * rs!cant, "#,###.#0")

MSHFlexGrid1.AddItem linea, i
TotVent = TotVent + Val(Str(rs!Neto))
totComp = totComp + rs!PREPROV * rs!cant
Tot = Tot + rs!Neto - rs!PREPROV * rs!cant
rs.MoveNext
Wend

Label2 = "GANANCIA " & Format(Tot, "#,###.#0")
Label3 = "TOTAL DE ESTA VENTA " & Format(TotVent, "#,###.#0")
Label4 = "COSTO DE MERCADERÍA " & Format(totComp, "#,###.#0")


CerrarBase


End Sub

