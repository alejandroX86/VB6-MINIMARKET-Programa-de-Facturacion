VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBuscarClienteF 
   BackColor       =   &H00000000&
   Caption         =   "Consultar Cliente"
   ClientHeight    =   5745
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11205
   Icon            =   "frmBuscarClienteF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   11205
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
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      Top             =   5220
      Width           =   1995
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ordenar por Nº de Cliente:"
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
      Width           =   4755
      Begin VB.CommandButton cmdOrdenar 
         Caption         =   "Ordenar"
         Height          =   315
         Left            =   3780
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
         Top             =   360
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
      Caption         =   "Buscar Cliente:"
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
         ItemData        =   "frmBuscarClienteF.frx":0442
         Left            =   3240
         List            =   "frmBuscarClienteF.frx":044F
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
      Bindings        =   "frmBuscarClienteF.frx":0469
      Height          =   3900
      Left            =   60
      TabIndex        =   11
      Top             =   1275
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6879
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
      Top             =   5310
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
      Top             =   5250
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
Attribute VB_Name = "frmBuscarClienteF"
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

strSql = "SELECT * FROM Clientes"

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
strSql = "Select * From Clientes"

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
strSql = "Select * From Clientes Where ID BETWEEN " & Val(txtDesde) & " AND " & Val(txtHasta) & " ORDER BY ID Asc"

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

strSql = "Select * From Clientes Where ID=" & Val(nrolista)

AbrirBase
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rs.EOF Then
'traigo campos
If frmFacturacion.Visible = True Then
frmFacturacion.txtCodCliente = Val(nrolista)
frmFacturacion.txtNombre = rs1!nombre
frmFacturacion.txtDomicilio = rs1!domicilio
frmFacturacion.txtTelefono = rs1!telefono
frmFacturacion.ComboCategIva = rs1!categiva
frmFacturacion.txtCuit = rs1!cuit
'Habilito desabilito
frmFacturacion.txtCodCliente.Enabled = False
frmFacturacion.txtNombre.Enabled = True
frmFacturacion.txtDomicilio.Enabled = True
frmFacturacion.txtTelefono.Enabled = True
frmFacturacion.ComboCategIva.Enabled = True
frmFacturacion.txtCuit.Enabled = True
frmFacturacion.txtNombre.SetFocus
frmFacturacion.txtCodarticulo.Enabled = True

ElseIf frmOrdenPedido.Visible = True Then
'traigo campos
frmOrdenPedido.txtCodCliente = Val(nrolista)
frmOrdenPedido.txtNombre = rs1!nombre
frmOrdenPedido.txtDomicilio = rs1!domicilio
frmOrdenPedido.txtTelefono = rs1!telefono
frmOrdenPedido.ComboCategIva = rs1!categiva
frmOrdenPedido.txtCuit = rs1!cuit
'Habilito desabilito
frmOrdenPedido.txtCodCliente.Enabled = False
frmOrdenPedido.txtNombre.Enabled = True
frmOrdenPedido.txtDomicilio.Enabled = True
frmOrdenPedido.txtTelefono.Enabled = True
frmOrdenPedido.ComboCategIva.Enabled = True
frmOrdenPedido.txtCuit.Enabled = True
frmOrdenPedido.txtNombre.SetFocus
frmOrdenPedido.txtCodarticulo.Enabled = True
ElseIf frmPresupuestos.Visible = True Then
'traigo campos
frmPresupuestos.txtCodCliente = Val(nrolista)
frmPresupuestos.txtNombre = rs1!nombre
frmPresupuestos.txtDomicilio = rs1!domicilio
frmPresupuestos.txtTelefono = rs1!telefono
frmPresupuestos.ComboCategIva = rs1!categiva
frmPresupuestos.txtCuit = rs1!cuit
'Habilito desabilito
frmPresupuestos.txtCodCliente.Enabled = False
frmPresupuestos.txtNombre.Enabled = True
frmPresupuestos.txtDomicilio.Enabled = True
frmPresupuestos.txtTelefono.Enabled = True
frmPresupuestos.ComboCategIva.Enabled = True
frmPresupuestos.txtCuit.Enabled = True
frmPresupuestos.txtNombre.SetFocus
frmPresupuestos.txtCodarticulo.Enabled = True

End If




End If
CerrarBase
Unload Me
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
