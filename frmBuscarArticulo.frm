VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBuscarArticulo 
   BackColor       =   &H00000000&
   Caption         =   "ConsultarArticulos"
   ClientHeight    =   5415
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7665
   Icon            =   "frmBuscarArticulo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7665
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
      Left            =   5580
      TabIndex        =   8
      Top             =   4905
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Buscar Articulo"
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
      TabIndex        =   1
      Top             =   120
      Width           =   7605
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
         ItemData        =   "frmBuscarArticulo.frx":0442
         Left            =   3540
         List            =   "frmBuscarArticulo.frx":0452
         TabIndex        =   3
         Text            =   "Descripcion"
         Top             =   360
         Width           =   2115
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar por:"
         Height          =   195
         Index           =   12
         Left            =   2700
         TabIndex        =   4
         Top             =   420
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmBuscarArticulo.frx":0479
      Height          =   3675
      Left            =   0
      TabIndex        =   5
      Top             =   1185
      Width           =   7575
      _ExtentX        =   13361
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
      Left            =   15
      TabIndex        =   7
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
      Left            =   1455
      TabIndex        =   6
      Top             =   4935
      Width           =   1395
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
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
         Caption         =   "&Codigo de Articulo"
         Begin VB.Menu HistoriaClinicaAscendente 
            Caption         =   "Orden Ascendente"
         End
         Begin VB.Menu HistoriaClinicaDescendente 
            Caption         =   "Oden Descendente"
         End
      End
      Begin VB.Menu mnuApellido 
         Caption         =   "&Descripcion"
         Begin VB.Menu ApellidoAscendente 
            Caption         =   "Orden Ascendente"
         End
         Begin VB.Menu ApellidoDescendente 
            Caption         =   "Orden Descendente"
         End
      End
   End
End
Attribute VB_Name = "frmBuscarArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
MostrarTodos
End Sub

Private Sub LimpioFlex()
titulos = " Codigo.|Descripción|Color|Talle|Precio"
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
End Sub
Private Sub MostrarTodos()
LimpioFlex
AbrirBase
Dim strSql As String
strSql = "SELECT * FROM Articulos"
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    Do While Not rs.EOF
        i = i + 1
        linea = rs!id _
                & Chr(9) & rs!Descripcion _
                & Chr(9) & rs!Marca _
                & Chr(9) & rs!Talle _
                & Chr(9) & Format(rs!Precio, "Fixed")

    MSHFlexGrid1.AddItem linea, i
    rs.MoveNext
    Loop
    lblTotalRef = rs.RecordCount

CerrarBase
AutoFlex MSHFlexGrid1
End Sub


'######################## MENUES #############################33

Private Sub mnuSalir_Click()
Unload Me
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
LimpioFlex
AbrirBase
Dim rs As New Recordset
Dim strSql As String
strSql = "Select * From Articulos"

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

    Do While Not rs.EOF
        i = i + 1
        linea = rs!id _
                & Chr(9) & rs!Descripcion _
                & Chr(9) & rs!Marca _
                & Chr(9) & rs!Talle _
                & Chr(9) & Format(rs!Precio, "Fixed")

    MSHFlexGrid1.AddItem linea, i
    rs.MoveNext
    Loop
    lblTotalRef = rs.RecordCount

CerrarBase
AutoFlex Me.MSHFlexGrid1
End Sub

Private Sub mnuCerrar_Click()
Unload Me
End Sub


Private Sub mnuImprimir_Click()
    If MsgBox("¿Imprimir Listado?", vbYesNo + vbInformation, "Impresión") = vbYes Then
    'IniciarImpresion
    End If
End Sub

Private Sub MSHFlexGrid1_DblClick()

If frmFacturacion.Visible = True Then

MSHFlexGrid1.Col = 0
frmFacturacion.txtCodarticulo.Text = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 1
v = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 2
v = v & " " & MSHFlexGrid1.Text
MSHFlexGrid1.Col = 3
v = v & " " & MSHFlexGrid1.Text
frmFacturacion.txtDescripcion.Text = v
MSHFlexGrid1.Col = 4
frmFacturacion.txtPrecio.Text = Replace(Format(MSHFlexGrid1.Text, "fixed"), ",", ".")
frmFacturacion.txtCantidad.Enabled = True
frmFacturacion.txtCantidad.SetFocus

ElseIf frmOrdenPedido.Visible = True Then

MSHFlexGrid1.Col = 0
frmOrdenPedido.txtCodarticulo.Text = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 1
v = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 2
v = v & " " & MSHFlexGrid1.Text
MSHFlexGrid1.Col = 3
v = v & " " & MSHFlexGrid1.Text
frmOrdenPedido.txtDescripcion.Text = v
MSHFlexGrid1.Col = 4
frmOrdenPedido.txtPrecio.Text = Replace(Format(MSHFlexGrid1.Text, "fixed"), ",", ".")
frmOrdenPedido.txtCantidad.Enabled = True
frmOrdenPedido.txtCantidad.SetFocus

ElseIf frmPresupuestos.Visible = True Then
MSHFlexGrid1.Col = 0
frmPresupuestos.txtCodarticulo.Text = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 1
v = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 2
v = v & " " & MSHFlexGrid1.Text
MSHFlexGrid1.Col = 3
v = v & " " & MSHFlexGrid1.Text
frmPresupuestos.txtDescripcion.Text = v
MSHFlexGrid1.Col = 4
frmPresupuestos.txtPrecio.Text = Replace(Format(MSHFlexGrid1.Text, "fixed"), ",", ".")
frmPresupuestos.txtCantidad.Enabled = True
frmPresupuestos.txtCantidad.SetFocus



End If

Unload Me
End Sub
