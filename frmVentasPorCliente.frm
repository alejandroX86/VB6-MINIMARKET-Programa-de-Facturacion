VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVentasPorCliente 
   BackColor       =   &H00000000&
   Caption         =   "Informe De Ventas por Cliente"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmVentasPorCliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
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
      TabIndex        =   7
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
         TabIndex        =   10
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
         ItemData        =   "frmVentasPorCliente.frx":57E2
         Left            =   3540
         List            =   "frmVentasPorCliente.frx":57EC
         TabIndex        =   9
         Text            =   "Nombre"
         Top             =   270
         Width           =   2115
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   5790
         TabIndex        =   8
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Buscar por:"
         Height          =   195
         Index           =   12
         Left            =   2700
         TabIndex        =   11
         Top             =   330
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6090
      Left            =   45
      TabIndex        =   6
      Top             =   810
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   10742
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
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
      Left            =   6660
      TabIndex        =   5
      Top             =   7080
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
      Left            =   8055
      TabIndex        =   4
      Top             =   7080
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
      Left            =   9480
      TabIndex        =   0
      Top             =   7080
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5400
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   435
      Left            =   4320
      TabIndex        =   12
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
      TabIndex        =   13
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
      TabIndex        =   15
      Top             =   345
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Peridodo Hasta:"
      Height          =   195
      Left            =   3090
      TabIndex        =   14
      Top             =   345
      Width           =   1275
   End
   Begin VB.Label lblGanancia 
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
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   8685
      Width           =   4875
   End
   Begin VB.Label lblVentas 
      BackColor       =   &H0000FFFF&
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   4860
   End
   Begin VB.Label lblCosto 
      BackColor       =   &H0000FFFF&
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7320
      Width           =   4890
   End
End
Attribute VB_Name = "frmVentasPorCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
etiqueta = "Estadistica de venta por día y por mes"
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
titulos = " NumFact|Tipo|Fecha|Cliente|Cond Venta|Subtotal|IVA|Total|Anulada|NotaCredito"

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
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
    
    End With




Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"



Dim total As Double
Dim subtotal As Double
Dim iva As Double
Dim Globaltotal As Double
Dim GlobalIva As Double
Dim GlobalSubtotal As Double



AbrirBase


strSql = "Select * From Clientes Order BY ID ASC"

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rs.EOF Then
While Not rs.EOF

iva = 0
subtotal = 0
total = 0


    Dim rs1 As New Recordset
    strSql = "SELECT * FROM FACTURASA WHERE CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
   
    If Not rs1.EOF Then
    While Not rs1.EOF
    i = i + 1
                    
                    linea = rs1!numfact _
                    & Chr(9) & rs1!tipofact _
                    & Chr(9) & rs1!Fecha _
                    & Chr(9) & rs!nombre _
                    & Chr(9) & rs1!condventa _
                    & Chr(9) & Format(rs1!subtotal, "standard") _
                    & Chr(9) & Format(rs1!iva, "standard") _
                    & Chr(9) & Format(rs1!total, "standard") _
                    & Chr(9) & rs1!Anulada _
                    & Chr(9) & rs1!NotaCredito
    
    
    MSHFlexGrid1.AddItem linea, i
    
    
    
    
    If rs1!Anulada = 0 Then
        If rs1!NotaCredito = 0 Then
        subtotal = CDbl(subtotal) + CDbl(rs1!subtotal)
        iva = CDbl(iva) + CDbl(rs1!iva)
        total = CDbl(total) + CDbl(rs1!total)
        Else
        subtotal = CDbl(subtotal) - CDbl(rs1!subtotal)
        iva = CDbl(iva) - CDbl(rs1!iva)
        total = CDbl(total) - CDbl(rs1!total)
        End If
    End If
    
    rs1.MoveNext
    Wend
    End If
    rs1.Close



    strSql = "SELECT * FROM FACTURASB WHERE CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    If Not rs1.EOF Then
    While Not rs1.EOF
    i = i + 1
                    
                    linea = rs1!numfact _
                    & Chr(9) & rs1!tipofact _
                    & Chr(9) & rs1!Fecha _
                    & Chr(9) & rs!nombre _
                    & Chr(9) & rs1!condventa _
                    & Chr(9) & Format(rs1!subtotal, "standard") _
                    & Chr(9) & Format(rs1!iva, "standard") _
                    & Chr(9) & Format(rs1!total, "standard") _
                    & Chr(9) & rs1!Anulada _
                    & Chr(9) & rs1!NotaCredito
    
    
    MSHFlexGrid1.AddItem linea, i
    
    
    If rs1!Anulada = 0 Then
        If rs1!NotaCredito = 0 Then
        subtotal = CDbl(subtotal) + CDbl(rs1!subtotal)
        iva = CDbl(iva) + CDbl(rs1!iva)
        total = CDbl(total) + CDbl(rs1!total)
        Else
        subtotal = CDbl(subtotal) - CDbl(rs1!subtotal)
        iva = CDbl(iva) - CDbl(rs1!iva)
        total = CDbl(total) - CDbl(rs1!total)
        End If
    End If
    
    
    
    
    rs1.MoveNext
    Wend
    End If
    rs1.Close


    strSql = "SELECT * FROM FACTURASC WHERE CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    If Not rs1.EOF Then
    While Not rs1.EOF
    i = i + 1
                    
                        linea = rs1!numfact _
                        & Chr(9) & rs1!tipofact _
                        & Chr(9) & rs1!Fecha _
                        & Chr(9) & rs!nombre _
                        & Chr(9) & rs1!condventa _
                        & Chr(9) & Format(rs1!subtotal, "standard") _
                        & Chr(9) & Format(rs1!iva, "standard") _
                        & Chr(9) & Format(rs1!total, "standard") _
                        & Chr(9) & rs1!Anulada _
                        & Chr(9) & rs1!NotaCredito
    
        
        MSHFlexGrid1.AddItem linea, i
        
        
    
    If rs1!Anulada = 0 Then
        If rs1!NotaCredito = 0 Then
        subtotal = CDbl(subtotal) + CDbl(rs1!subtotal)
        iva = CDbl(iva) + CDbl(rs1!iva)
        total = CDbl(total) + CDbl(rs1!total)
        Else
        subtotal = CDbl(subtotal) - CDbl(rs1!subtotal)
        iva = CDbl(iva) - CDbl(rs1!iva)
        total = CDbl(total) - CDbl(rs1!total)
        End If
    End If
    
    
    
    rs1.MoveNext
    Wend
    End If
    rs1.Close


    strSql = "SELECT * FROM FACTURASX WHERE CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    If Not rs1.EOF Then
    While Not rs1.EOF
    i = i + 1
                    
                        linea = rs1!numfact _
                        & Chr(9) & rs1!tipofact _
                        & Chr(9) & rs1!Fecha _
                        & Chr(9) & rs!nombre _
                        & Chr(9) & rs1!condventa _
                        & Chr(9) & Format(rs1!subtotal, "standard") _
                        & Chr(9) & Format(rs1!iva, "standard") _
                        & Chr(9) & Format(rs1!total, "standard") _
                        & Chr(9) & rs1!Anulada _
                        & Chr(9) & rs1!NotaCredito
        
        MSHFlexGrid1.AddItem linea, i
        
        
    
    If rs1!Anulada = 0 Then
        If rs1!NotaCredito = 0 Then
        subtotal = CDbl(subtotal) + CDbl(rs1!subtotal)
        iva = CDbl(iva) + CDbl(rs1!iva)
        total = CDbl(total) + CDbl(rs1!total)
        Else
        subtotal = CDbl(subtotal) - CDbl(rs1!subtotal)
        iva = CDbl(iva) - CDbl(rs1!iva)
        total = CDbl(total) - CDbl(rs1!total)
        End If
    End If
    
    
    
    rs1.MoveNext
    Wend
    End If
    rs1.Close


If total > 0 Then

i = i + 1
                
                linea = "" _
                & Chr(9) & "" _
                & Chr(9) & "" _
                & Chr(9) & "" _
                & Chr(9) & "TOTAL" _
                & Chr(9) & Format(subtotal, "standard") _
                & Chr(9) & Format(iva, "standard") _
                & Chr(9) & Format(total, "standard") _


MSHFlexGrid1.AddItem linea, i


GlobalSubtotal = CDbl(GlobalSubtotal) + CDbl(subtotal)
GlobalIva = CDbl(GlobalIva) + CDbl(iva)
Globaltotal = CDbl(Globaltotal) + CDbl(total)


End If


rs.MoveNext
Wend
End If


lblVentas = "Total de ventas sin IVA. $ " & Format(GlobalSubtotal, "standard")
lblCosto = "Total periodo IVA $ " & Format(GlobalIva, "standard")
lblGanancia = "Total Ventas + IVA: $ " & Format(Globaltotal, "standard")


AutoFlex Me.MSHFlexGrid1
FlexRayado MSHFlexGrid1, &HFFFFFF, &H8000000F


CerrarBase
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
' AHORA MUESTRO LA TABLA
titulos = " NumFact|Tipo|Fecha|Cliente|Cond Venta|Subtotal|IVA|Total|Anulada|NotaCredito"

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
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
    
    End With




Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"



Dim total As Double
Dim subtotal As Double
Dim iva As Double
Dim Globaltotal As Double
Dim GlobalIva As Double
Dim GlobalSubtotal As Double



AbrirBase


strSql = "Select * From Clientes Order BY ID ASC"

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText



Select Case comboCriterio.Text
Case "CodCliente"
    If txtBuscar <> "" Then
    rs.Filter = "ID =" & Val(txtBuscar)
    End If
Case "Nombre"
    If txtBuscar <> "" Then
    rs.Filter = "Nombre LIKE '*" + txtBuscar + "*'"
    End If
'Case "Cliente"
'    If txtBuscar <> "" Then
'    rs.Filter = "CodCliente LIKE '*" + txtBuscar + "*'"
'    End If
'
'Case "Talle"
'    If txtBuscar <> "" Then
'    rs.Filter = "Talle LIKE '*" + txtBuscar + "*'"
'    End If

Case Else
rs.Filter = ""
End Select





If Not rs.EOF Then
While Not rs.EOF

iva = 0
subtotal = 0
total = 0


    Dim rs1 As New Recordset
    strSql = "SELECT * FROM FACTURASA WHERE CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
   
    If Not rs1.EOF Then
    While Not rs1.EOF
    i = i + 1
                    
                    linea = rs1!numfact _
                    & Chr(9) & rs1!tipofact _
                    & Chr(9) & rs1!Fecha _
                    & Chr(9) & rs!nombre _
                    & Chr(9) & rs1!condventa _
                    & Chr(9) & Format(rs1!subtotal, "standard") _
                    & Chr(9) & Format(rs1!iva, "standard") _
                    & Chr(9) & Format(rs1!total, "standard") _
                    & Chr(9) & rs1!Anulada _
                    & Chr(9) & rs1!NotaCredito
    
    
    MSHFlexGrid1.AddItem linea, i
    
    
    
    
    If rs1!Anulada = 0 Then
        If rs1!NotaCredito = 0 Then
        subtotal = CDbl(subtotal) + CDbl(rs1!subtotal)
        iva = CDbl(iva) + CDbl(rs1!iva)
        total = CDbl(total) + CDbl(rs1!total)
        Else
        subtotal = CDbl(subtotal) - CDbl(rs1!subtotal)
        iva = CDbl(iva) - CDbl(rs1!iva)
        total = CDbl(total) - CDbl(rs1!total)
        End If
    End If
    
    rs1.MoveNext
    Wend
    End If
    rs1.Close



    strSql = "SELECT * FROM FACTURASB WHERE CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    If Not rs1.EOF Then
    While Not rs1.EOF
    i = i + 1
                    
                    linea = rs1!numfact _
                    & Chr(9) & rs1!tipofact _
                    & Chr(9) & rs1!Fecha _
                    & Chr(9) & rs!nombre _
                    & Chr(9) & rs1!condventa _
                    & Chr(9) & Format(rs1!subtotal, "standard") _
                    & Chr(9) & Format(rs1!iva, "standard") _
                    & Chr(9) & Format(rs1!total, "standard") _
                    & Chr(9) & rs1!Anulada _
                    & Chr(9) & rs1!NotaCredito
    
    
    MSHFlexGrid1.AddItem linea, i
    
    
    If rs1!Anulada = 0 Then
        If rs1!NotaCredito = 0 Then
        subtotal = CDbl(subtotal) + CDbl(rs1!subtotal)
        iva = CDbl(iva) + CDbl(rs1!iva)
        total = CDbl(total) + CDbl(rs1!total)
        Else
        subtotal = CDbl(subtotal) - CDbl(rs1!subtotal)
        iva = CDbl(iva) - CDbl(rs1!iva)
        total = CDbl(total) - CDbl(rs1!total)
        End If
    End If
    
    
    
    
    rs1.MoveNext
    Wend
    End If
    rs1.Close


    strSql = "SELECT * FROM FACTURASC WHERE CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    If Not rs1.EOF Then
    While Not rs1.EOF
    i = i + 1
                    
                        linea = rs1!numfact _
                        & Chr(9) & rs1!tipofact _
                        & Chr(9) & rs1!Fecha _
                        & Chr(9) & rs!nombre _
                        & Chr(9) & rs1!condventa _
                        & Chr(9) & Format(rs1!subtotal, "standard") _
                        & Chr(9) & Format(rs1!iva, "standard") _
                        & Chr(9) & Format(rs1!total, "standard") _
                        & Chr(9) & rs1!Anulada _
                        & Chr(9) & rs1!NotaCredito
    
        
        MSHFlexGrid1.AddItem linea, i
        
        
    
    If rs1!Anulada = 0 Then
        If rs1!NotaCredito = 0 Then
        subtotal = CDbl(subtotal) + CDbl(rs1!subtotal)
        iva = CDbl(iva) + CDbl(rs1!iva)
        total = CDbl(total) + CDbl(rs1!total)
        Else
        subtotal = CDbl(subtotal) - CDbl(rs1!subtotal)
        iva = CDbl(iva) - CDbl(rs1!iva)
        total = CDbl(total) - CDbl(rs1!total)
        End If
    End If
    
    
    
    rs1.MoveNext
    Wend
    End If
    rs1.Close


    strSql = "SELECT * FROM FACTURASX WHERE CodCLIENTE=" & Val(rs!id) & " AND Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY NUMFACT ASC"
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    If Not rs1.EOF Then
    While Not rs1.EOF
    i = i + 1
                    
                        linea = rs1!numfact _
                        & Chr(9) & rs1!tipofact _
                        & Chr(9) & rs1!Fecha _
                        & Chr(9) & rs!nombre _
                        & Chr(9) & rs1!condventa _
                        & Chr(9) & Format(rs1!subtotal, "standard") _
                        & Chr(9) & Format(rs1!iva, "standard") _
                        & Chr(9) & Format(rs1!total, "standard") _
                        & Chr(9) & rs1!Anulada _
                        & Chr(9) & rs1!NotaCredito
        
        MSHFlexGrid1.AddItem linea, i
        
        
    
    If rs1!Anulada = 0 Then
        If rs1!NotaCredito = 0 Then
        subtotal = CDbl(subtotal) + CDbl(rs1!subtotal)
        iva = CDbl(iva) + CDbl(rs1!iva)
        total = CDbl(total) + CDbl(rs1!total)
        Else
        subtotal = CDbl(subtotal) - CDbl(rs1!subtotal)
        iva = CDbl(iva) - CDbl(rs1!iva)
        total = CDbl(total) - CDbl(rs1!total)
        End If
    End If
    
    
    
    rs1.MoveNext
    Wend
    End If
    rs1.Close


If total > 0 Then

i = i + 1
                
                linea = "" _
                & Chr(9) & "" _
                & Chr(9) & "" _
                & Chr(9) & "" _
                & Chr(9) & "TOTAL" _
                & Chr(9) & Format(subtotal, "standard") _
                & Chr(9) & Format(iva, "standard") _
                & Chr(9) & Format(total, "standard") _


MSHFlexGrid1.AddItem linea, i


GlobalSubtotal = CDbl(GlobalSubtotal) + CDbl(subtotal)
GlobalIva = CDbl(GlobalIva) + CDbl(iva)
Globaltotal = CDbl(Globaltotal) + CDbl(total)


End If


rs.MoveNext
Wend
End If


lblVentas = "Total de ventas sin IVA. $ " & Format(GlobalSubtotal, "standard")
lblCosto = "Total periodo IVA $ " & Format(GlobalIva, "standard")
lblGanancia = "Total Ventas + IVA: $ " & Format(Globaltotal, "standard")


AutoFlex Me.MSHFlexGrid1
FlexRayado MSHFlexGrid1, &HFFFFFF, &H8000000F


CerrarBase
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



