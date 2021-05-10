VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVentasPorCtaCte 
   BackColor       =   &H00000000&
   Caption         =   "Informe De Ventas por Cta Cte."
   ClientHeight    =   7995
   ClientLeft      =   -30
   ClientTop       =   165
   ClientWidth     =   11880
   Icon            =   "frmVentasPorCtaCte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "EXPORTAR"
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
      Left            =   7440
      TabIndex        =   6
      Top             =   7515
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "IMPRIMIR"
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
      Left            =   8835
      TabIndex        =   5
      Top             =   7515
      Width           =   1335
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
      Left            =   10275
      TabIndex        =   1
      Top             =   7515
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmVentasPorCtaCte.frx":57E2
      Height          =   6210
      Left            =   45
      TabIndex        =   0
      Top             =   810
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   10954
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   255
      ForeColorFixed  =   16777215
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2700
      Top             =   7020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   435
      Left            =   4320
      TabIndex        =   7
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
      TabIndex        =   8
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
      TabIndex        =   10
      Top             =   345
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Peridodo Hasta:"
      Height          =   195
      Left            =   3090
      TabIndex        =   9
      Top             =   345
      Width           =   1275
   End
   Begin VB.Label lblGanancia 
      BackColor       =   &H00000000&
      Caption         =   "lblGanancia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6525
      TabIndex        =   4
      Top             =   7140
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblVentas 
      BackColor       =   &H00000000&
      Caption         =   "lblVentas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   7140
      Width           =   3375
   End
   Begin VB.Label lblCosto 
      BackColor       =   &H00000000&
      Caption         =   "lblCosto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   7140
      Visible         =   0   'False
      Width           =   2955
   End
End
Attribute VB_Name = "frmVentasPorCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
etiqueta = "Ventas por cliente en cta cte."
Me.Caption = etiqueta

dtp1 = Date
dtp2 = Date
VerVentasPorArticulo
End Sub

'imprimir
Private Sub cmdImprimir_Click()
ImprimirFlex MSHFlexGrid1, Me.Caption
End Sub
'close
Private Sub Command1_Click()
Unload Me
End Sub

'exportar
Private Sub Command2_Click()
FlexGrid_To_Excel Me.MSHFlexGrid1, MSHFlexGrid1.Rows, MSHFlexGrid1.Cols, Me.Caption
End Sub

Private Sub dtp1_Change()
VerVentasPorArticulo
End Sub
Private Sub dtp2_Change()
VerVentasPorArticulo
End Sub



Private Sub VerVentasPorArticulo()






' AHORA MUESTRO LA TABLA


titulos = " NumFact|Tipo|Fecha|Cliente|Cond Venta|Total  "

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
        .ColAlignmentFixed = 3
        .ColWidth(5) = 1500
    End With








Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"

AbrirBase




strSql = "Select * From Clientes Order BY ID ASC"

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rs.EOF Then
While Not rs.EOF

TotVent = 0


    Dim rs1 As New Recordset
    strSql = "SELECT * FROM FACTURASA WHERE CondVenta LIKE " & "'" & "Cuenta Corriente" & "'" & " AND CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
   
    If Not rs1.EOF Then
    While Not rs1.EOF
    i = i + 1
                    
                    linea = rs1!numfact _
                    & Chr(9) & rs1!tipofact _
                    & Chr(9) & rs1!Fecha _
                    & Chr(9) & rs!nombre _
                    & Chr(9) & rs1!condventa _
                    & Chr(9) & Format(rs1!total, "standard")
    
    MSHFlexGrid1.AddItem linea, i
    TotVent = CDbl(TotVent) + CDbl(rs1!total)
    rs1.MoveNext
    Wend
    End If
    rs1.Close



    strSql = "SELECT * FROM FACTURASB WHERE CondVenta LIKE " & "'" & "Cuenta Corriente" & "'" & " AND CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
        rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    If Not rs1.EOF Then
    While Not rs1.EOF
    i = i + 1
                    
                    linea = rs1!numfact _
                    & Chr(9) & rs1!tipofact _
                    & Chr(9) & rs1!Fecha _
                    & Chr(9) & rs!nombre _
                    & Chr(9) & rs1!condventa _
                    & Chr(9) & Format(rs1!total, "standard")
    
    MSHFlexGrid1.AddItem linea, i
    TotVent = CDbl(TotVent) + CDbl(rs1!total)
    rs1.MoveNext
    Wend
    End If
    rs1.Close


    strSql = "SELECT * FROM FACTURASC WHERE CondVenta LIKE " & "'" & "Cuenta Corriente" & "'" & " AND CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rs1.EOF Then
While Not rs1.EOF
i = i + 1
                
                    linea = rs1!numfact _
                    & Chr(9) & rs1!tipofact _
                    & Chr(9) & rs1!Fecha _
                    & Chr(9) & rs!nombre _
                    & Chr(9) & rs1!condventa _
                    & Chr(9) & Format(rs1!total, "standard")

MSHFlexGrid1.AddItem linea, i
TotVent = CDbl(TotVent) + CDbl(rs1!total)
rs1.MoveNext
Wend
End If
rs1.Close


    strSql = "SELECT * FROM FACTURASX WHERE CondVenta LIKE " & "'" & "Cuenta Corriente" & "'" & " AND CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rs1.EOF Then
While Not rs1.EOF
i = i + 1
                
                    linea = rs1!numfact _
                    & Chr(9) & rs1!tipofact _
                    & Chr(9) & rs1!Fecha _
                    & Chr(9) & rs!nombre _
                    & Chr(9) & rs1!condventa _
                    & Chr(9) & Format(rs1!total, "standard")

MSHFlexGrid1.AddItem linea, i
TotVent = CDbl(TotVent) + CDbl(rs1!total)
rs1.MoveNext
Wend
End If
rs1.Close


If TotVent > 0 Then

i = i + 1
                
                linea = "" _
                & Chr(9) & "" _
                & Chr(9) & "" _
                & Chr(9) & "" _
                & Chr(9) & "TOTAL" _
                & Chr(9) & Format(TotVent, "standard")

MSHFlexGrid1.AddItem linea, i

TotVentGlobal = CDbl(TotVentGlobal) + CDbl(TotVent)
End If


rs.MoveNext
Wend
End If


lblVentas = "Total de ventas en Cta.Cte. $ " & Format(TotVentGlobal, "standard")
'lblCosto = "COSTO MERCADERIA: $ " & Format(totComp, "#,###.#0")
'lblGanancia = "GANANCIA PERIODO: $ " & Format(totUt, "#,###.#0")


AutoFlex Me.MSHFlexGrid1

CerrarBase
End Sub

