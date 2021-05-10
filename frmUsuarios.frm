VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUsuarios 
   BackColor       =   &H00000000&
   Caption         =   "Modulo ABM de Usuarios del Sistema"
   ClientHeight    =   4635
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6330
   Icon            =   "frmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   300
      TabIndex        =   15
      Top             =   840
      Width           =   5655
      Begin VB.ComboBox ComboPermiso 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmUsuarios.frx":0442
         Left            =   1380
         List            =   "frmUsuarios.frx":044C
         TabIndex        =   4
         Text            =   "Administrador del Sistema"
         Top             =   1920
         Width           =   3435
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtUsuario 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtClave 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1380
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton CmdContinuar 
         Caption         =   "C&ontinuar"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   2700
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Image Image4 
         Height          =   600
         Left            =   4920
         Picture         =   "frmUsuarios.frx":0480
         Top             =   1740
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre            :"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Código            :"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario            :"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Clave               :"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Permiso           :"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   120
         X2              =   5460
         Y1              =   2700
         Y2              =   2700
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   13
      Top             =   4335
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "21/09/10"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:21"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1058
      ButtonWidth     =   2249
      ButtonHeight    =   953
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo Usuario"
            Key             =   "ToolNuevo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar Usuario"
            Key             =   "ToolGuardar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar Usuario"
            Key             =   "ToolEliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar Todo"
            Key             =   "ToolCancelar"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":0604
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":0A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":0EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":1300
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":1754
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":1BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":1FFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":2450
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":28A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6300
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNuevo 
         Caption         =   "&Nuevo..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Imprimir Ficha..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "&Guardar"
         Shortcut        =   ^S
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuBuscar 
         Caption         =   "&Buscar..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuCancelar 
         Caption         =   "&Cancelar..."
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuPromedios 
         Caption         =   "&Listado..."
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAyudaContextual 
         Caption         =   "&Ayuda Contextual..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu qq 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCredits 
         Caption         =   "&Sobre el Programa..."
         Shortcut        =   ^J
      End
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Left = 0
Me.Top = 0
Me.Height = 5350
Me.Width = 6500
End Sub

Private Sub cmdBuscar_Click()
frmBuscarUsuario.Show
End Sub
Private Sub cmdactualizar_click()
Guardar
End Sub

Private Sub cmdCancelar_Click()
DeshabilitarEdicion
End Sub

Private Sub cmdcontinuar_click()
DeshabilitarEdicion
End Sub

Private Sub cmdeliminar_click()
Eliminar
End Sub

Private Sub cmdmodificar_click()
HabilitarEdicion
txtNombre.SetFocus
End Sub

Private Sub CmdNuevo_Click()
NuevoRegistro
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


'&&&&&&&&&&&&&&&&&& MENUES &&&&&&&&&&&&&&&&&&&&&&&


Private Sub mnuCredits_Click()
'frmAbout.Show
End Sub

Private Sub mnuGuardar_Click()
Guardar
End Sub

Private Sub mnuNuevo_Click()
NuevoRegistro
End Sub
Private Sub mnuImprimir_Click()
'Imprimir
End Sub

Private Sub mnuSalir_Click()
Salir
End Sub

Private Sub mnuAyudaContextual_Click()
MsgBox "El pack de Ayuda está en construcción", vbInformation
End Sub

Private Sub mnuPromedios_Click()
frmBuscarUsuario.Show
End Sub
Private Sub mnuBuscar_Click()
frmBuscarUsuario.Show
End Sub


'&&&&&&&&&&&&&&&&&& Barra de Herramientas &&&&&&&&&&&&&&&&&

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "ToolNuevo"
NuevoRegistro
Case "ToolGuardar"
Guardar
Case "ToolEliminar"
Eliminar
Case "ToolCancelar"
DeshabilitarEdicion
Case "ToolBuscar"
'frmBuscarSocio.Show
Case "ToolPrint"
'ImprimirFicha
Case "ToolPromedios"
'frmBuscarPaciente.Show
End Select

End Sub

'&&&&&&&&&&&&&&&&&&&& Botones &&&&&&&&&&&&&&&&&&&&&&&


    
Private Sub Salir()
If MsgBox("¿Está seguro de salir de la aplicación?", vbYesNoCancel, "Salir") = vbYes Then
Unload Me
End If
End Sub



' ################ EVENTOS #############################



' ################ NUEVO REGISTRO #############################


Private Sub NuevoRegistro()
Dim strSql As String
Dim cont As Integer
DeshabilitarEdicion
AbrirBase

strSql = "SELECT * FROM Usuarios"
        
        rs.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
        If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        cont = Val(rs!id)
        txtID = cont + 1
        
        Else
        txtID = "1"
        End If
CerrarBase
HabilitarEdicion
PreparoTemplate
End Sub

Private Sub PreparoTemplate()
'TbLoc = "Río Cuarto"
'TbCodPos = "5800"
'TbNac = "Arg."
txtNombre.SetFocus
End Sub

' ################ ELIMINAR #############################



Private Sub Eliminar()
If txtID <> "" Then
    If MsgBox("ATENCIÓN: ¿Realmente Elimina este Registro?", vbYesNo + vbExclamation, "Eliminar") = vbYes Then
    EliminarRegistro
    End If
End If
End Sub
Private Sub EliminarRegistro()
'On Error GoTo Error_Eliminar
AbrirBase
'db.BeginTrans
EliminarCaso
'EliminarCuotasRelacionadas
'Error_Eliminar:
'    If Err.Number <> 0 Then
'    db.RollbackTrans
'    CerrarBase
'    MsgBox "Error: " & Err.Number & " - " & Err.Description & " - " & Err.Source, vbCritical
'    Else
'    db.CommitTrans
    CerrarBase
    DeshabilitarEdicion
'    End If
End Sub

Private Sub EliminarCaso()
Dim strSql As String
strSql = "SELECT * FROM Usuarios WHERE ID =" & Val(txtID)
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not (rs.BOF And rs.EOF) Then
        rs.Delete
        rs.Update
        End If

End Sub


' ################ ACEPTAR #############################

Private Sub TxtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(txtID) Then
        KeyAscii = 0
        Aceptar txtID.Text
    End If
End Sub


Private Sub Aceptar(ByVal txtID As Variant)
'On Error GoTo Error_Guardar

AbrirBase
'DB.BeginTrans
AceptarRegistro
'Error_Guardar:
'    If Err.Number <> 0 Then
'    DB.RollbackTrans
'    CerrarBase
'    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
'    Else
'    DB.CommitTrans
    CerrarBase
    'HabilitarEdicion
'    End If

End Sub


Private Sub AceptarRegistro()
Dim strSql As String
strSql = "SELECT * FROM Usuarios WHERE ID =" & Val(txtID)

rs.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
           
  ' Si existe
If Not (rs.BOF And rs.EOF) Then
          
TraigoDatos
BotoneraExploracion
Else
HabilitarNuevo
HabilitarEdicion
PreparoTemplate
End If

End Sub



Private Sub TraigoDatos()
  On Error Resume Next
  
    'Traigo textos
    
    txtID.Text = RTrim(rs!id)
    
    
    txtNombre.Text = RTrim(rs!nombre)
    
        txtUsuario.Text = RTrim(rs!Usuario)
    
        txtClave.Text = RTrim(rs!Clave)
    
        ComboPermiso.Text = RTrim(rs!Permiso)
    
   
    

End Sub


' ################ G U A R D A R #############################
'##############################################################
Private Sub Guardar()

If IsNumeric(txtID) And txtID <> "" Then


        GuardarTodo

End If
End Sub


Private Sub GuardarTodo()
'On Error GoTo Error_Guardar

AbrirBase
'DB.BeginTrans
ActualizarRegistro
'Error_Guardar:
'    If Err.Number <> 0 Then
'    DB.RollbackTrans
'    MsgBox "Error: " & Err.Number & " - " & Err.Description
'    CerrarBase
'    Else
'   DB.CommitTrans
    CerrarBase
    DeshabilitarEdicion
 '   End If


End Sub


Private Sub ActualizarRegistro()

strSql = "SELECT * FROM Usuarios WHERE ID =" & Val(txtID)

rs.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
        'Si no existe que lo agregue
        If rs.BOF And rs.EOF Then
        rs.AddNew
        GuardarRegistros
        'Verifico_existencia_de_cuotas
        Else
        GuardarRegistros
        'Verifico_existencia_de_cuotas
        End If

End Sub


Private Sub GuardarRegistros()


    With rs
        !id = Val(txtID)
        !nombre = UCase(txtNombre.Text)
        !Usuario = txtUsuario.Text
        !Clave = txtClave.Text
        !Permiso = ComboPermiso.Text
        .Update
    End With

End Sub


'################# Habilitar / Deshabilitar ################
Private Sub HabilitarNuevo()
    
    Dim saveNroSoc As Integer
    
    saveNroSoc = txtID
    'Primero vacío los controles
        For i = 1 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is TextBox Then
                Me.Controls(i).Text = ""
            End If
    Next i
    txtID = saveNroSoc
End Sub



Private Sub HabilitarEdicion()
    ' Habilito todos los textBox


    For i = 1 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is TextBox Then
                Me.Controls(i).Enabled = True
            End If
    Next i
    
    
    'Habilito todos los Combobox
    
    For i = 1 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is ComboBox Then
                Me.Controls(i).Enabled = True
            End If
    Next i
    
    ' Deshabilito el campo clave
    txtID.Enabled = False
    'Y pongo el Foco en el Campo Nombre
    txtNombre.SetFocus
BotoneraEdicion
End Sub


Private Sub DeshabilitarEdicion()
    
    
    
    
    
    ' Primero vacío los controles TextBox
    For i = 1 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is TextBox Then
                Me.Controls(i).Text = ""
            End If
    Next i
    ' Después deshabilito los controles TextBox
    For i = 1 To Me.Controls.Count - 1
        If TypeOf Me.Controls(i) Is TextBox Then
            Me.Controls(i).Enabled = False
        End If
    Next i
    

    'Tambien desabilito los controles combobox
    
        For i = 1 To Me.Controls.Count - 1
        If TypeOf Me.Controls(i) Is ComboBox Then
            Me.Controls(i).Enabled = False
        End If
        Next i
    
'        For i = 1 To Me.Controls.Count - 1
'        If TypeOf Me.Controls(i) Is ComboBox Then
'            Me.Controls(i).Visible = False
'        End If
'        Next i

    
    
    ' Después habilito el campo de número de
    ' historia clínica y le coloco el foco
    txtID.Enabled = True
    txtID.SetFocus
   BotoneraNeutra
End Sub




' ####################### KeyPress  ##########################

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtNombre = "" Then
DeshabilitarEdicion
KeyAscii = 0
End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtUsuario = "" Then
txtNombre.SetFocus
KeyAscii = 0
End If
End Sub


Private Sub txtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtClave = "" Then
txtUsuario.SetFocus
KeyAscii = 0
End If
End Sub



Private Sub ComboPermiso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And ComboPermiso = "" Then
txtClave.SetFocus
KeyAscii = 0
End If
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'TODO ESTO ES PARA ANIMAR UN COMBO
Private Sub tbtipdoc_gotfocus()
    
    CbTipDoc.Height = 1150
    CbTipDoc.Visible = True
    CbTipDoc.Text = TbTipDoc.Text
    CbTipDoc.SetFocus
End Sub

Private Sub cbtipdoc_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cbtipdoc_dblclick
    End If
End Sub

Private Sub cbtipdoc_dblclick()
       TbTipDoc.Text = CbTipDoc.Text
       'QueTipDoc = CbTipDoc.ListIndex + 1
       CbTipDoc.Visible = False
       'noenter = False
       TbNroDoc.SetFocus
End Sub
Private Sub cbtipdoc_lostfocus()
    CbTipDoc.Visible = False
End Sub





' POSICION DE BOTONES CON RESPECTO AL MOMENTO DE EDICION

Private Sub BotoneraNeutra()

cmdBuscar.Visible = True

CmdContinuar.Visible = False

CmdActualizar.Visible = False

cmdCancelar.Visible = False

CmdModificar.Visible = False

CmdEliminar.Visible = False

'CmdImprimir.Visible = False

End Sub

Private Sub BotoneraExploracion()

cmdBuscar.Visible = False

CmdContinuar.Visible = True

CmdActualizar.Visible = False



cmdCancelar.Visible = False

CmdModificar.Visible = True


CmdEliminar.Visible = True

'CmdImprimir.Visible = True

CmdContinuar.SetFocus

End Sub

Private Sub BotoneraEdicion()

cmdBuscar.Visible = False

CmdContinuar.Visible = False

CmdActualizar.Visible = True


cmdCancelar.Visible = True

CmdModificar.Visible = False


CmdEliminar.Visible = True

'CmdImprimir.Visible = True


CmdActualizar.SetFocus

End Sub






