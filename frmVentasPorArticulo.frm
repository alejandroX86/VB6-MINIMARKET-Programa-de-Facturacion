VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVentasPorArticulo 
   BackColor       =   &H0000FFFF&
   Caption         =   "Informe De Ventas por Producto"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmVentasPorArticulo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FFFF&
      Height          =   765
      Left            =   6165
      TabIndex        =   11
      Top             =   45
      Width           =   6930
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   5790
         TabIndex        =   14
         Top             =   270
         Width           =   945
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
         ItemData        =   "frmVentasPorArticulo.frx":57E2
         Left            =   3540
         List            =   "frmVentasPorArticulo.frx":57F2
         TabIndex        =   13
         Text            =   "Descripcion"
         Top             =   270
         Width           =   2115
      End
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
         TabIndex        =   12
         Top             =   270
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Buscar por:"
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
         Height          =   195
         Index           =   12
         Left            =   2460
         TabIndex        =   15
         Top             =   330
         Width           =   1095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6165
      Left            =   0
      TabIndex        =   10
      Top             =   855
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   10874
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
      Left            =   5085
      TabIndex        =   9
      Top             =   7470
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
      Left            =   6480
      TabIndex        =   8
      Top             =   7470
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   435
      Left            =   4425
      TabIndex        =   2
      Top             =   270
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
      Left            =   1365
      TabIndex        =   1
      Top             =   270
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
      Left            =   7920
      TabIndex        =   0
      Top             =   7470
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2745
      Top             =   7065
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
      Left            =   6570
      TabIndex        =   7
      Top             =   7185
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblVentas 
      BackColor       =   &H0000FFFF&
      Caption         =   "lblVentas"
      Height          =   255
      Left            =   105
      TabIndex        =   6
      Top             =   7185
      Width           =   3375
   End
   Begin VB.Label lblCosto 
      BackColor       =   &H0000FFFF&
      Caption         =   "lblCosto"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   7200
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Peridodo Hasta:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3195
      TabIndex        =   4
      Top             =   390
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Peridodo Desde:"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   390
      Width           =   1275
   End
End
Attribute VB_Name = "frmVentasPorArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dtp1 = Date
dtp2 = Date
VerVentasPorArticulo
End Sub

'imprimir
Private Sub cmdImprimir_Click()
ImprimirFlex MSHFlexGrid1, "Listado de Ventas por Producto"
End Sub
'close
Private Sub Command1_Click()
Unload Me
End Sub

'exportar
Private Sub Command2_Click()
FlexGrid_To_Excel Me.MSHFlexGrid1, MSHFlexGrid1.Rows, MSHFlexGrid1.Cols, "Ventas por Articulo"
End Sub

Private Sub dtp1_Change()
VerVentasPorArticulo
End Sub
Private Sub dtp2_Change()
VerVentasPorArticulo
End Sub



Private Sub VerVentasPorArticulo()
Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"

AbrirBase


'Primero Vaciamos la tabla VentasPorArticulo
strSql2 = "Select * From VentasPorArticulo"
Dim rsDel As New Recordset
rsDel.Open strSql2, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsDel.EOF Then
While Not rsDel.EOF
rsDel.Delete
rsDel.MoveNext
Wend
End If

'LLENAMOS LA TABLA DE ARTICULOS
strSql1 = "Select * From Articulos"
strSql2 = "Select * From VentasPorArticulo"
Dim rs1 As New Recordset
Dim rs2 As New Recordset
rs1.Open strSql1, DB, adOpenKeyset, adLockOptimistic, adCmdText
rs2.Open strSql2, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rs1.EOF Then
While Not rs1.EOF
rs2.AddNew
rs2!id = rs1!id
rs2!Descripcion = rs1!Descripcion
rs2!Marca = rs1!Marca & " "
rs2!Talle = rs1!Talle & " "
rs2!Precio = rs1!Precio
rs2!PrecioProv = rs1!PrecioProv
rs2.Update
rs1.MoveNext
Wend
End If


Dim varTMP As Variant
varTMP = 0


' Luego Cantidad, P_UNITARIO y P_NETO DE FACTURAS A

strSql = "Select ANulada,NotaCredito, Fecha, CodArticulo, Cantidad, P_UNITARIO, P_NETO FROM FacturasA INNER JOIN DetalleFacturasA ON FacturasA.NumFact=DetalleFacturasA.Numfact WHERE Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY FECHA ASC"
Dim rsFA As New Recordset
rsFA.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rsFA.EOF Then
While Not rsFA.EOF

strSql3 = "Select * From VentasPorArticulo Where ID=" & Val(rsFA!codarticulo)
Dim rsVPA As New Recordset
rsVPA.Open strSql3, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsVPA.EOF Then

If rsFA!Anulada = 0 Then

    If rsFA!NotaCredito = 0 Then
    
    rsVPA!Cantidad = Val(rsVPA!Cantidad) + Val(rsFA!Cantidad)
    rsVPA!P_NETO = CDbl(rsVPA!P_NETO) + CDbl(rsFA!P_NETO)
    rsVPA.Update
    
    Else
    rsVPA!Cantidad = Val(rsVPA!Cantidad) - Val(rsFA!Cantidad)
    rsVPA!P_NETO = CDbl(rsVPA!P_NETO) - CDbl(rsFA!P_NETO)
    rsVPA.Update
    End If


End If

End If
rsVPA.Close

rsFA.MoveNext
Wend
End If
rsFA.Close

' Luego Cantidad, P_UNITARIO y P_NETO DE FACTURAS B

strSql = "Select ANulada,NotaCredito, Fecha, CodArticulo, Cantidad, P_UNITARIO, P_NETO FROM FacturasB INNER JOIN DetalleFacturasB ON FacturasB.NumFact=DetalleFacturasB.Numfact WHERE Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY FECHA ASC"
Dim rsFB As New Recordset
rsFB.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rsFB.EOF Then
While Not rsFB.EOF
strSql3 = "Select * From VentasPorArticulo Where ID=" & Val(rsFB!codarticulo)
Dim rsVPB As New Recordset
rsVPB.Open strSql3, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsVPB.EOF Then


If rsFB!Anulada = 0 Then


If rsFB!NotaCredito = 0 Then

rsVPB!Cantidad = Val(rsVPB!Cantidad) + Val(rsFB!Cantidad)
rsVPB!P_NETO = CDbl(rsVPB!P_NETO) + CDbl(rsFB!P_NETO)


rsVPB.Update

Else
rsVPB!Cantidad = Val(rsVPB!Cantidad) - Val(rsFB!Cantidad)
rsVPB!P_NETO = CDbl(rsVPB!P_NETO) - CDbl(rsFB!P_NETO)

rsVPB.Update
End If


End If



End If
rsVPB.Close
rsFB.MoveNext
Wend
End If
rsFB.Close

' Luego Cantidad, P_UNITARIO y P_NETO DE FACTURAS C

strSql = "Select ANulada,NotaCredito, Fecha, CodArticulo, Cantidad, P_UNITARIO, P_NETO FROM FacturasC INNER JOIN DetalleFacturasC ON FacturasC.NumFact=DetalleFacturasC.Numfact WHERE Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY FECHA ASC"
Dim rsFC As New Recordset
rsFC.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rsFC.EOF Then
While Not rsFC.EOF
strSql3 = "Select * From VentasPorArticulo Where ID=" & Val(rsFC!codarticulo)
Dim rsVPC As New Recordset
rsVPC.Open strSql3, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsVPC.EOF Then




If ncFC!Anulada = 0 Then

If rsFC!NotaCredito = 0 Then

rsVPC!Cantidad = Val(rsVPC!Cantidad) + Val(rsFC!Cantidad)
rsVPC!P_NETO = CDbl(rsVPC!P_NETO) + CDbl(rsFC!P_NETO)
rsVPC.Update
Else
rsVPC!Cantidad = Val(rsVPC!Cantidad) - Val(rsFC!Cantidad)
rsVPC!P_NETO = CDbl(rsVPC!P_NETO) - CDbl(rsFC!P_NETO)
rsVPC.Update

End If


End If


End If
rsVPC.Close
rsFC.MoveNext
Wend
End If
rsFC.Close


' Luego Cantidad, P_UNITARIO y P_NETO DE FacturaX

strSql = "Select ANulada,NotaCredito, Fecha, CodArticulo, Cantidad, P_UNITARIO, P_NETO FROM FACTURASX INNER JOIN DETALLEFACTURASX ON FacturasX.NumFact=DetalleFacturasX.Numfact WHERE FacturasX.Fecha BETWEEN " & strDesde & " AND " & strHasta & " ORDER BY FECHA ASC"
Dim rsOP As New Recordset
rsOP.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rsOP.EOF Then
While Not rsOP.EOF
strSql3 = "Select * From VentasPorArticulo Where ID=" & Val(rsOP!codarticulo)
Dim rsVPOP As New Recordset
rsVPOP.Open strSql3, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsVPOP.EOF Then



If rsOP!Anulada = 0 Then

If rsOP!NotaCredito = 0 Then
rsVPOP!Cantidad = Val(rsVPOP!Cantidad) + Val(rsOP!Cantidad)
rsVPOP!P_NETO = CDbl(rsVPOP!P_NETO) + CDbl(rsOP!P_NETO)
rsVPOP.Update
Else
rsVPOP!Cantidad = Val(rsVPOP!Cantidad) - Val(rsOP!Cantidad)
rsVPOP!P_NETO = CDbl(rsVPOP!P_NETO) - CDbl(rsOP!P_NETO)
rsVPOP.Update

End If

End If


End If
rsVPOP.Close
rsOP.MoveNext
Wend
End If
rsOP.Close


'Ahora Elimino los Articulos Sin venta

strSql7 = "Select * From VentasPorArticulo Where Cantidad=0"
Dim rs7 As New Recordset
rs7.Open strSql7, DB, adOpenKeyset, adLockOptimistic, adCmdText
While Not rs7.EOF
rs7.Delete
rs7.MoveNext
Wend



' AHORA MUESTRO LA TABLA


titulos = " CodArt.|DESCRIPCION|Color|Talle|Cantidad|P_NETO"

 MSHFlexGrid1.Clear
    With MSHFlexGrid1
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
        .ColAlignmentFixed(0) = 3
        .ColAlignmentFixed(1) = 3
        .ColAlignmentFixed(2) = 3
        .ColAlignmentFixed(3) = 3
        .ColAlignmentFixed(4) = 3
        .ColAlignmentFixed(5) = 3
        '.ColAlignmentFixed(6) = 3
        '.ColAlignmentFixed(7) = 3
        '.ColAlignmentFixed(8) = 3
    End With



strSql = "Select * FROM VentasPorArticulo"
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

While Not rs.EOF
                i = i + 1
                
                linea = rs!id _
                & Chr(9) & rs!Descripcion _
                & Chr(9) & rs!Marca _
                & Chr(9) & rs!Talle _
                & Chr(9) & Format(rs!Cantidad, "0###") _
                & Chr(9) & Format(rs!P_NETO, "standard")
                
MSHFlexGrid1.AddItem linea, i

TotVent = CDbl(TotVent) + CDbl(rs!P_NETO)
'totComp = CDbl(totComp) + CDbl(rs!CMV)
'totUt = CDbl(totUt) + CDbl(rs!Ganancia)
rs.MoveNext
Wend

lblVentas = "TOTAL VENTAS: $ " & Format(TotVent, "#,###.#0")
'lblCosto = "COSTO MERCADERIA: $ " & Format(totComp, "#,###.#0")
'lblGanancia = "GANANCIA PERIODO: $ " & Format(totUt, "#,###.#0")


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
'AHORA MUESTRO LA TABLA
titulos = " CodArt.|DESCRIPCION|Color|Talle|Cantidad|P_NETO"

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
        .ColAlignmentFixed(0) = 3
        .ColAlignmentFixed(1) = 3
        .ColAlignmentFixed(2) = 3
        .ColAlignmentFixed(3) = 3
        .ColAlignmentFixed(4) = 3
        .ColAlignmentFixed(5) = 3
        '.ColAlignmentFixed(6) = 3
        '.ColAlignmentFixed(7) = 3
        '.ColAlignmentFixed(8) = 3
    End With


Dim strDesde As String
Dim strHasta As String
strDesde = "#" & Format(dtp1, "mm/dd/yyyy") & "#"
strHasta = "#" & Format(dtp2, "mm/dd/yyyy") & "#"


AbrirBase

strSql = "Select * FROM VentasPorArticulo"
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

Select Case comboCriterio.Text
Case "Codigo"
    If txtBuscar <> "" Then
    rs.Filter = "ID =" & Val(txtBuscar)
    End If
Case "Descripcion"
    If txtBuscar <> "" Then
    rs.Filter = "Descripcion LIKE '*" + txtBuscar + "*'"
    End If
Case "Marca"
    If txtBuscar <> "" Then
    rs.Filter = "Marca LIKE '*" + txtBuscar + "*'"
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
                
                linea = rs!id _
                & Chr(9) & rs!Descripcion _
                & Chr(9) & rs!Marca _
                & Chr(9) & rs!Talle _
                & Chr(9) & Format(rs!Cantidad, "0###") _
                & Chr(9) & Format(rs!P_NETO, "standard")
                
MSHFlexGrid1.AddItem linea, i

TotVent = CDbl(TotVent) + CDbl(rs!P_NETO)
'totComp = CDbl(totComp) + CDbl(rs!CMV)
'totUt = CDbl(totUt) + CDbl(rs!Ganancia)
rs.MoveNext
Wend

lblVentas = "TOTAL VENTAS: $ " & Format(TotVent, "#,###.#0")
'lblCosto = "COSTO MERCADERIA: $ " & Format(totComp, "#,###.#0")
'lblGanancia = "GANANCIA PERIODO: $ " & Format(totUt, "#,###.#0")

AutoFlex Me.MSHFlexGrid1
FlexRayado MSHFlexGrid1, &HFFFFFF, &H8000000F


CerrarBase
End Sub

