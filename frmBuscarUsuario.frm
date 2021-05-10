VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBuscarUsuario 
   BackColor       =   &H00000000&
   Caption         =   "Buscar Usuario..."
   ClientHeight    =   5745
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8550
   Icon            =   "frmBuscarUsuario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8550
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
      Left            =   6180
      TabIndex        =   8
      Top             =   5280
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
         Left            =   3240
         TabIndex        =   3
         Text            =   "Nombre"
         Top             =   360
         Width           =   2115
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   5400
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Buscar por:"
         Height          =   375
         Index           =   12
         Left            =   2400
         TabIndex        =   4
         Top             =   420
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3675
      Left            =   60
      TabIndex        =   5
      Top             =   1560
      Width           =   8355
      _ExtentX        =   14737
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
      TabIndex        =   7
      Top             =   1260
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
      TabIndex        =   6
      Top             =   1200
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
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmBuscarUsuario"
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
titulos = " Codigo.|Nombre|Usuario|Password|Permiso"
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
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 0
        .ColWidth(4) = 3000
'        .ColWidth(5) = 1000
'        .ColWidth(6) = 1500
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

strSql = "SELECT * FROM Usuarios"

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

HacerLinea

CerrarBase




End Sub


'######################## MENUES #############################33






Private Sub mnuSalir_Click()
Unload Me
End Sub

Private Sub mnuVer_Click()

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
strSql = "Select * From Usuarios"

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText



Select Case comboCriterio.Text
Case "Nombre"
    If txtBuscar <> "" Then
    rs.Filter = "Nombre LIKE '*" + txtBuscar + "*'"
    Else
    rs.Filter = ""
    End If

Case "Usuario"
    If txtBuscar <> "" Then
    rs.Filter = "Usuario LIKE '*" + txtBuscar + "*'"
    Else
    rs.Filter = ""
    End If

'Case "Telefono"
'    If txtBuscar <> "" Then
'    rs.Filter = "Telefono LIKE '*" + txtBuscar + "*'"
'    Else
'    rs.Filter = ""
'    End If
'
'
'Case "Email"
'    If txtBuscar <> "" Then
'    rs.Filter = "Email LIKE '*" + txtBuscar + "*'"
'    Else
'    rs.Filter = ""
'    End If
Case Else
rs.Filter = ""
End Select

HacerLinea

CerrarBase
End Sub

Private Sub mnuCerrar_Click()
Unload Me
End Sub


Private Sub cmdBuscar_Click()
AceptarBuscar txtBuscar.Text
End Sub




Private Sub MSHFlexGrid1_DblClick()
On Error Resume Next
MSHFlexGrid1.Col = 0
frmUsuarios.txtID.Text = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 1
frmUsuarios.txtNombre.Text = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 2
frmUsuarios.txtUsuario.Text = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 3
frmUsuarios.txtClave.Text = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 4
frmUsuarios.ComboPermiso.Text = MSHFlexGrid1.Text

frmUsuarios.txtID.Enabled = True
frmUsuarios.txtID.SetFocus
Unload Me

End Sub

Private Sub HacerLinea()


    Do While Not rs.EOF
        i = i + 1
        linea = rs!id _
                & Chr(9) & rs!nombre _
                & Chr(9) & rs!Usuario _
                & Chr(9) & rs!Clave _
                & Chr(9) & rs!Permiso _

    MSHFlexGrid1.AddItem linea, i
    rs.MoveNext
    Loop
    lblTotalRef = rs.RecordCount


End Sub



