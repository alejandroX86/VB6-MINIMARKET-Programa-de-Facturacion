VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBuscarProveedor 
   BackColor       =   &H00000000&
   Caption         =   "Consultar Proveedor"
   ClientHeight    =   5370
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmBuscarProveedor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9810
      TabIndex        =   14
      Top             =   4995
      Width           =   1230
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ordenar por Nº de Proveedor:"
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
      Height          =   1035
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Width           =   5835
      Begin VB.CommandButton cmdOrdenar 
         Caption         =   "Ordenar"
         Height          =   315
         Left            =   4500
         TabIndex        =   10
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txtHasta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   2
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtDesde 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   660
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FFFF&
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   1860
         TabIndex        =   9
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FFFF&
         Caption         =   "Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Buscar Proveedor:"
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
      Height          =   1035
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   6315
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
         TabIndex        =   0
         Top             =   360
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
         ItemData        =   "frmBuscarProveedor.frx":0442
         Left            =   3240
         List            =   "frmBuscarProveedor.frx":044F
         TabIndex        =   5
         Text            =   "Nombre"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Buscar por:"
         Height          =   375
         Index           =   12
         Left            =   2400
         TabIndex        =   6
         Top             =   420
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmBuscarProveedor.frx":0469
      Height          =   3675
      Left            =   60
      TabIndex        =   11
      Top             =   1215
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   6482
      _Version        =   393216
      FixedCols       =   0
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      Caption         =   "Total Registros:"
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
      Left            =   60
      TabIndex        =   13
      Top             =   4995
      Width           =   1395
   End
   Begin VB.Label lblTotalRef 
      BackColor       =   &H0000FFFF&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   12
      Top             =   4935
      Width           =   1395
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuRefrescarDatos 
         Caption         =   "&Refrescar Datos"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
   End
   Begin VB.Menu mnuOrdenar 
      Caption         =   "&Ordenar datos"
      Begin VB.Menu mnuHistoriaClinica 
         Caption         =   "&Nº Cliente"
         Begin VB.Menu HistoriaClinicaAscendente 
            Caption         =   "Orden Ascendente"
         End
         Begin VB.Menu HistoriaClinicaDescendente 
            Caption         =   "Oden Descendente"
         End
      End
      Begin VB.Menu mnuApellido 
         Caption         =   "&Apellido"
         Begin VB.Menu ApellidoAscendente 
            Caption         =   "Orden Ascendente"
         End
         Begin VB.Menu ApellidoDescendente 
            Caption         =   "Orden Descendente"
         End
      End
      Begin VB.Menu mnuDNI 
         Caption         =   "&DNI"
         Begin VB.Menu DNIAscendente 
            Caption         =   "Orden Ascendente"
         End
         Begin VB.Menu DNIDescendente 
            Caption         =   "Orden Desdendente"
         End
      End
   End
End
Attribute VB_Name = "frmBuscarProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
MostrarTodos
End Sub

Private Sub LimpioFlex()
titulos = " Cod.Cliente.|Nombre Completo|Domicilio|Teléfono|Nº Cuit|E-mail"
lblTotalRef = "#"
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
        .ColWidth(1) = 3000
        .ColWidth(2) = 2500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 2000
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With
End Sub

Private Sub MostrarTodos()

LimpioFlex
On Error GoTo Errorbuscar
AbrirBase
Dim rsRef As New Recordset
Dim strSql As String

strSql = "SELECT * FROM Proveedores"

rsRef.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    Do While Not rsRef.EOF

        i = i + 1
       
        ' Y TODO POR ESTA LINEA DE MIERDA
        'VENCIMIENTO / DESCRIPCION / IMPORTE DE CUOTA / FECHA DE PAGO
        linea = rsRef!id _
                & Chr(9) & rsRef!nombre _
                & Chr(9) & rsRef!domicilio _
                & Chr(9) & rsRef!telefono _
                & Chr(9) & rsRef!cuit _
                & Chr(9) & rsRef!Email
                
    MSHFlexGrid1.AddItem linea, i
    rsRef.MoveNext
    Loop
    lblTotalRef = rsRef.RecordCount

CerrarBase

Errorbuscar:
    If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Number & " - " & Err.Description
    MsgBox "La busqueda no arrojó resultados"
    Resume Next
    End If



End Sub

Private Sub mnuRefrescarDatos_Click()
MostrarTodos
End Sub

'######################## MENUES #############################33


Private Sub mnuSalir_Click()
Unload Me
End Sub

'################PROCEDIMIENTO BUSCAR ####################

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
LimpioFlex
AbrirBase

Dim strSql As String
strSql = "Select * From Proveedores"

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText



Select Case comboCriterio.Text

Case "Nombre"
    If txtBuscar <> "" Then
    rs.Filter = "Nombre LIKE '*" + txtBuscar + "*'"
    Else
    rs.Filter = ""
    End If
Case "Codigo"
    If txtBuscar <> "" Then
    rs.Filter = "ID =" & Val(txtBuscar)
    Else
    rs.Filter = ""
    End If
Case "Cuit"
    If txtBuscar <> "" Then
    rs.Filter = "Cuit LIKE '*" + txtBuscar + "*'"
    Else
    rs.Filter = ""
    End If
Case Else
End Select

HacerLinea

CerrarBase
End Sub

Private Sub mnuCerrar_Click()
Me.Hide
End Sub


Private Sub cmdBuscar_Click()
AceptarBuscar txtBuscar.Text
End Sub


''''''''''''''''''''"ORDENAR POR RANGOS"###############################

Private Sub txtDesde_Change()
FiltrarHistoria
End Sub

Private Sub txtHasta_Change()
FiltrarHistoria
End Sub


Private Sub FiltrarHistoria()

LimpioFlex
AbrirBase
Dim strSql As String
strSql = "Select * From Proveedores Where ID BETWEEN " & Val(txtDesde) & " AND " & Val(txtHasta) & " ORDER BY ID Asc"

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

HacerLinea

CerrarBase

End Sub



Private Sub dtp1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
End If
End Sub


Private Sub txtHistoriaDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
End If
End Sub

Private Sub txtHistoriaHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtBuscar.SetFocus
'SendKeys ("{tab}")
KeyAscii = 0
End If
End Sub

Private Sub mnuImprimir_Click()
    If MsgBox("¿Imprimir Listado?", vbYesNo + vbInformation, "Impresión") = vbYes Then
    'IniciarImpresion
    End If
End Sub


Private Sub MSHFlexGrid1_DblClick()
On Error Resume Next
Dim nrolista As Long
MSHFlexGrid1.Col = 0
nrolista = MSHFlexGrid1.Text
'frmClientes.txtCodDist.SetFocus
Dim rs1 As New Recordset

strSql = "Select * From Proveedores Where ID=" & Val(nrolista)

AbrirBase
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rs.EOF Then
'traigo campos
frmProveedores.txtCodDist = Val(nrolista)
frmProveedores.txtFechaAlta = rs1!FechaAlta
frmProveedores.txtNombre = rs1!nombre
frmProveedores.txtDomicilio = rs1!domicilio
frmProveedores.txtLocalidad = rs1!Localidad
frmProveedores.txtCodPost = rs1!CodPost
frmProveedores.txtTelefono = rs1!telefono
frmProveedores.ComboCategIva = rs1!categiva
frmProveedores.txtEmail = rs1!Email
frmProveedores.txtCuit = rs1!cuit
'Habilito desabilito
'frmProveedores.txtCodDist.Enabled = False
'frmProveedores.txtFechaAlta.Enabled = True
'frmProveedores.txtNombre.Enabled = True
'frmProveedores.txtDomicilio.Enabled = True
'frmProveedores.txtLocalidad.Enabled = True
'frmProveedores.txtCodPost.Enabled = True
'frmProveedores.txtTelefono.Enabled = True
'frmProveedores.ComboCategIva.Enabled = True
'frmProveedores.txtEmail.Enabled = True
'frmProveedores.txtCuit.Enabled = True
'frmProveedores.txtNombre.SetFocus

frmProveedores.cmdBuscar.Visible = False
frmProveedores.CmdContinuar.Visible = True
frmProveedores.CmdActualizar.Visible = False
frmProveedores.cmdCancelar.Visible = False
frmProveedores.CmdModificar.Visible = True
frmProveedores.CmdEliminar.Visible = True
'frmProveedores.cmdImprimir.Visible = True
frmProveedores.Command1.Visible = True
frmProveedores.CmdContinuar.SetFocus
VerCuentaCorriente

End If
CerrarBase
Unload Me
End Sub


Private Sub VerCuentaCorriente()

titulos = "ID|Nº FACT.|FECHA|DESCRIPCIÓN|DEBE|HABER"

 frmClientes.MSHFlexGrid1.Clear
    With frmClientes.MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .ColWidth(0) = 0
        .ColWidth(1) = 1000
        .ColWidth(2) = 1200
        .ColWidth(3) = 2500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .ColAlignmentFixed = 3
    End With


Dim rsCons As New Recordset
strSql = "Select * From CuentasCorrientes Where Codcliente=" & Val(frmClientes.txtCodDist)
rsCons.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

TotDebe = "0"
TotHaber = "0"
While Not rsCons.EOF
                i = i + 1
                
                linea = rsCons!id _
                & Chr(9) & rsCons!numfact _
                & Chr(9) & rsCons!Fecha _
                & Chr(9) & rsCons!Descripcion _
                & Chr(9) & Format(rsCons!Debe, "fixed") _
                & Chr(9) & Format(rsCons!Haber, "fixed")

frmClientes.MSHFlexGrid1.AddItem linea, i
TotDebe = TotDebe + rsCons!Debe
TotHaber = TotHaber + rsCons!Haber
rsCons.MoveNext
Wend
frmClientes.lblTotalRef = rsCons.RecordCount
frmClientes.lblTotal = "Saldo Cta/Cte: $ " & Format(TotDebe - TotHaber, "#,###.#0")

End Sub



Private Sub HacerLinea()

    While Not rs.EOF
        i = i + 1
              linea = rs!id _
                & Chr(9) & rs!nombre _
                & Chr(9) & rs!domicilio _
                & Chr(9) & rs!telefono _
                & Chr(9) & rs!cuit _
                & Chr(9) & rs!Email
        
    MSHFlexGrid1.AddItem linea, i
    rs.MoveNext
    Wend
    lblTotalRef = rs.RecordCount

End Sub


