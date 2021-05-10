VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmClientes 
   BackColor       =   &H00000000&
   Caption         =   "ABM de Clientes"
   ClientHeight    =   7905
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10395
   Icon            =   "frmClientes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6795
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   10095
      Begin VB.CommandButton Command2 
         Caption         =   "Buscar"
         Height          =   270
         Left            =   3600
         TabIndex        =   41
         Top             =   585
         Width           =   915
      End
      Begin VB.CommandButton cmdInputPago 
         Caption         =   "Imputar Pago"
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
         TabIndex        =   40
         Top             =   6180
         Width           =   1575
      End
      Begin VB.OptionButton OptPresupuestos 
         Caption         =   "Ver Presupuestos"
         Height          =   255
         Left            =   4950
         TabIndex        =   38
         Top             =   315
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.OptionButton optCtaCte 
         Caption         =   "Ver Cuenta Corriente"
         Height          =   255
         Left            =   270
         TabIndex        =   37
         Top             =   2610
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.OptionButton optConsumos 
         Caption         =   "Ver Consumos"
         Height          =   255
         Left            =   2430
         TabIndex        =   36
         Top             =   2610
         Width           =   1470
      End
      Begin VB.ComboBox ComboCategIva 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmClientes.frx":0442
         Left            =   7500
         List            =   "frmClientes.frx":0458
         TabIndex        =   9
         Text            =   "IVA Consumidor Final"
         Top             =   1860
         Width           =   2475
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar Lista de Precios para este cliente"
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
         Left            =   4095
         TabIndex        =   34
         Top             =   2475
         Width           =   5790
      End
      Begin VB.TextBox txtCuit 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7500
         MaxLength       =   15
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtEmail 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1860
         Width           =   3375
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Nuevo..."
         Height          =   270
         Left            =   2520
         TabIndex        =   31
         Top             =   585
         Width           =   915
      End
      Begin VB.TextBox txtFechaAlta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7500
         MaxLength       =   10
         TabIndex        =   1
         Top             =   540
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6180
         Width           =   1035
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   6180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   6180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   2700
         TabIndex        =   14
         Top             =   6180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   6180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdContinuar 
         Caption         =   "C&ontinuar"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   6180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   6180
         Width           =   1095
      End
      Begin VB.TextBox txtTelefono 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7500
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtCodPost 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7500
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox txtCodDist 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   0
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox txtLocalidad 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtDomicilio 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1260
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Top             =   960
         Width           =   3375
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Bindings        =   "frmClientes.frx":04DC
         Height          =   3075
         Left            =   240
         TabIndex        =   35
         Top             =   2940
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5424
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   255
         ForeColorFixed  =   16777215
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
      Begin VB.Label lblTotal 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   5820
         TabIndex        =   39
         Top             =   6300
         Width           =   2895
      End
      Begin VB.Label lblTotalRef 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label Label8 
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
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label Label17 
         Caption         =   "CUIT Nº:"
         Height          =   255
         Left            =   6780
         TabIndex        =   30
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Email:"
         Height          =   255
         Left            =   780
         TabIndex        =   29
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha Alta:"
         Height          =   255
         Left            =   6660
         TabIndex        =   28
         Top             =   600
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         X1              =   120
         X2              =   9900
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label11 
         Caption         =   "Categ:"
         Height          =   255
         Left            =   6840
         TabIndex        =   27
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "Teléfono:"
         Height          =   255
         Left            =   6720
         TabIndex        =   26
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Localidad:"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   1620
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Cod. Postal:"
         Height          =   255
         Left            =   6540
         TabIndex        =   24
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Domicilio:"
         Height          =   255
         Left            =   540
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   660
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   21
         Top             =   1020
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   17
      Top             =   7605
      Width           =   10395
      _ExtentX        =   18336
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
            TextSave        =   "10:46"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   1058
      ButtonWidth     =   2249
      ButtonHeight    =   953
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo Cliente"
            Key             =   "ToolNuevo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar Cliente"
            Key             =   "ToolGuardar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Ficha"
            Key             =   "ToolPrint"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Baja Cliente"
            Key             =   "ToolEliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar Edición"
            Key             =   "ToolCancelar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar..."
            Key             =   "ToolBuscar"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   7080
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
            Picture         =   "frmClientes.frx":04F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientes.frx":0946
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientes.frx":0D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientes.frx":11EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientes.frx":1642
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientes.frx":1A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientes.frx":1EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientes.frx":233E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientes.frx":2792
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10380
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNuevo 
         Caption         =   "&Nuevo..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Imprimir Ficha"
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
         Caption         =   "&Cancelar Edición"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuPromedios 
         Caption         =   "&Listado de Clientes"
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
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstSocios As New Recordset
Dim RstCuotas As New Recordset
Dim rsVaciarTabImpSoc As New Recordset
Dim RstVarios As New Recordset
Dim RstImpSocios As New Recordset
Dim rsVerif As New Recordset
Dim TotCount As Long
Dim TotVent As Double


Private Sub Command1_Click()
If txtCodDist <> "" Then
frmPreciosClientes.Show
frmPreciosClientes.lblCodProveedor = txtCodDist
End If
End Sub



Private Sub Command2_Click()
frmBuscarCliente.Show

End Sub

Private Sub Form_Load()

'Me.Left = 0
'Me.Top = 0
'Me.Height = 9100
'Me.Width = 10400

If IsNumeric(txtCodDist) Then
KeyAscii = 0
Aceptar txtCodDist.Text
End If

'LimpioFlex

End Sub



Private Sub cmdBuscar_Click()
frmBuscarCliente.Show
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
txtFechaAlta.SetFocus
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
MsgBox "Esto no está listo aún", vbInformation
'ImprimirFicha
End Sub

Private Sub mnuSalir_Click()
Salir
End Sub

Private Sub mnuAyudaContextual_Click()
MsgBox "El pack de Ayuda está en construcción", vbInformation
End Sub

Private Sub mnuPromedios_Click()
frmBuscarCliente.Show
End Sub
Private Sub mnuBuscar_Click()
frmBuscarCliente.Show
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
frmBuscarCliente.Show
Case "ToolPrint"
ImprimirFicha
Case "ToolPromedios"
frmBuscarCliente.Show
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
Dim cont As Long
DeshabilitarEdicion
AbrirBase

strSql = "SELECT * FROM Clientes"
        
        rs.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
        If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        cont = CVar(rs!id)
        txtCodDist = cont + 1
        
        Else
        txtCodDist = "1"
        End If
CerrarBase
HabilitarEdicion
PreparoTemplate
End Sub

Private Sub PreparoTemplate()
txtLocalidad = ""
txtCodPost = ""
txtFechaAlta = Date
txtFechaAlta.SetFocus
End Sub

' ################ ELIMINAR #############################



Private Sub Eliminar()
If txtCodDist <> "" Then
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
strSql = "SELECT * FROM Clientes WHERE ID =" & Val(txtCodDist)
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not (rs.BOF And rs.EOF) Then
        rs.Delete
        rs.Update
        End If

End Sub



' ################ ACEPTAR #############################

Private Sub txtCodDist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(txtCodDist) Then
        KeyAscii = 0
        Aceptar txtCodDist.Text
    End If
End Sub


Private Sub Aceptar(ByVal txtCodDist As Variant)
'On Error GoTo Error_Guardar

AbrirBase
'DB.BeginTrans
AceptarRegistro
'LimpioFlex
'VerConsumos
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
VerOpcion
End Sub


Private Sub AceptarRegistro()
Dim strSql As String
strSql = "SELECT * FROM Clientes WHERE ID =" & Val(txtCodDist)

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
  
    'Traigo textos
    
    'Campo Clave
    txtCodDist.Text = RTrim(rs!id)

    
        'Fecha Alta

    If IsNull(rs!FechaAlta) Then
    txtFechaAlta.Text = Empty
    Else
    txtFechaAlta.Text = RTrim(rs!FechaAlta)
    End If
    'Codigo Patrocinador
    
    
    'Campo NombreSol
    If IsNull(rs!nombre) Then
    txtNombre.Text = Empty
    Else
    txtNombre.Text = RTrim(rs!nombre)
    End If
    
   If IsNull(rs!cuit) Then
    txtCuit.Text = Empty
    Else
        txtCuit.Text = RTrim(rs!cuit)
    End If
    
    
    If IsNull(rs!domicilio) Then
    txtDomicilio.Text = Empty
    Else
        txtDomicilio.Text = RTrim(rs!domicilio)
    End If
    
    If IsNull(rs!CodPost) Then
    txtCodPost.Text = Empty
    Else
        txtCodPost.Text = RTrim(rs!CodPost)
    End If
    
    If IsNull(rs!Localidad) Then
    txtLocalidad.Text = Empty
    Else
        txtLocalidad.Text = RTrim(rs!Localidad)
    End If

    If IsNull(rs!telefono) Then
    txtTelefono.Text = Empty
    Else
        txtTelefono.Text = RTrim(rs!telefono)
    End If
    
    'Combo Estado Civil
    
    If IsNull(rs!Email) Then
    txtEmail.Text = Empty
    Else
    txtEmail.Text = RTrim(rs!Email)
    End If
    
    
    
   
    
    If IsNull(rs!categiva) Then
    ComboCategIva.Text = ""
    Else
        ComboCategIva.Text = RTrim(rs!categiva)
    End If
    

End Sub











' ################ G U A R D A R #############################
'##############################################################
Private Sub Guardar()

If IsNumeric(txtCodDist) And txtCodDist <> "" Then


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

strSql = "SELECT * FROM Clientes WHERE ID =" & Val(txtCodDist)

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


    If IsNull(txtCodDist.Text) Or txtCodDist.Text = "" Then
    rs!id = 0
    Else
    rs!id = CLng(txtCodDist.Text)
    End If


    If IsNull(txtFechaAlta.Text) Or txtFechaAlta.Text = "" Then
    rs!FechaAlta = Date
    Else
    rs!FechaAlta = txtFechaAlta.Text
    End If



    If IsNull(txtNombre.Text) Or txtNombre.Text = "" Then
    rs!nombre = "NO ESPECIFICADO"
    Else
        rs!nombre = txtNombre.Text
    End If



   If IsNull(txtCuit.Text) Or txtCuit.Text = "" Then
        rs!cuit = 0
    Else
        rs!cuit = txtCuit.Text
    End If



    If IsNull(txtDomicilio.Text) Or txtDomicilio.Text = "" Then
    rs!domicilio = "NO ESPECIFICADO"
    Else
    rs!domicilio = txtDomicilio.Text
    End If

    
    If IsNull(txtCodPost.Text) Or txtCodPost.Text = "" Then
    rs!CodPost = 0
    Else
    rs!CodPost = txtCodPost.Text
    End If

    If IsNull(txtLocalidad.Text) Or txtLocalidad.Text = "" Then
    rs!Localidad = "NO ESPECIFICADO"
    Else
    rs!Localidad = txtLocalidad.Text
    End If
    
    
    If IsNull(txtTelefono.Text) Or txtTelefono.Text = "" Then
        rs!telefono = Empty
    Else
        rs!telefono = txtTelefono.Text
    End If
   
    If IsNull(txtEmail.Text) Or txtEmail.Text = "" Then
        rs!Email = Empty
    Else
        rs!Email = txtEmail.Text
    End If
    
    If IsNull(ComboCategIva.Text) Or ComboCategIva.Text = "" Then
        rs!categiva = Empty
    Else
        rs!categiva = ComboCategIva.Text
    End If



    rs.Update

End Sub



'################# Habilitar / Deshabilitar ################
Private Sub HabilitarNuevo()
    
    'Dim saveNroDist As Long
    
    saveNroDist = CLng(txtCodDist)
    'Primero vacío los controles
        For i = 1 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is TextBox Then
                Me.Controls(i).Text = ""
            End If
    Next i
    txtCodDist = CVar(saveNroDist)
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
  txtCodDist.Enabled = False
    'Y pongo el Foco en el Campo Nombre
   ' txtNombre.SetFocus
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
'          Me.Controls(i).Visible = False
'        End If
'        Next i

    
    
    ' Después habilito el campo de número de
    ' historia clínica y le coloco el foco
    txtCodDist.Enabled = True
    txtCodDist.SetFocus
   BotoneraNeutra
   MSHFlexGrid1.Clear
End Sub



'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        'If noenter = False Then
'           SendKeys "{tab}"
'           'KeyAscii = 0
'        'End If
'    End If
'End Sub




' ####################### KeyPress  ##########################
Private Sub txtFechaAlta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtFechaAlta = "" Then
DeshabilitarEdicion
KeyAscii = 0
End If
End Sub
Private Sub txtFechaAlta_GotFocus()
txtFechaAlta.SelLength = Len(txtFechaAlta.Text)
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtNombre = "" Then
txtFechaAlta.SetFocus
KeyAscii = 0
End If
End Sub
Private Sub txtNombre_GotFocus()
txtNombre.SelLength = Len(txtNombre.Text)
End Sub

Private Sub txtCuit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtCuit = "" Then
txtNombre.SetFocus
KeyAscii = 0
End If
End Sub
Private Sub txtCuit_GotFocus()
txtCuit.SelLength = Len(txtCuit.Text)
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtDomicilio = "" Then
txtCuit.SetFocus
KeyAscii = 0
End If
End Sub
Private Sub txtDomicilio_GotFocus()
txtDomicilio.SelLength = Len(txtDomicilio.Text)
End Sub

Private Sub txtCodPost_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtCodPost = "" Then
txtDomicilio.SetFocus
KeyAscii = 0
End If
End Sub
Private Sub txtCodPost_GotFocus()
 txtCodPost.SelLength = Len(txtCodPost.Text)
End Sub


Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtLocalidad = "" Then
txtCodPost.SetFocus
KeyAscii = 0
End If
End Sub

Private Sub txtLocalidad_GotFocus()
 txtLocalidad.SelLength = Len(txtLocalidad.Text)
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtTelefono = "" Then
txtLocalidad.SetFocus
KeyAscii = 0
End If
End Sub
Private Sub txtTelefono_GotFocus()
 txtTelefono.SelLength = Len(txtTelefono.Text)
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtEmail = "" Then
txtLocalidad.SetFocus
KeyAscii = 0
End If
End Sub

Private Sub txtEmail_GotFocus()
 txtEmail.SelLength = Len(txtEmail.Text)
End Sub


Private Sub ComboCategIva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And ComboCategIva = "" Then
txtEmail.SetFocus
KeyAscii = 0
End If
End Sub
Private Sub ComboCategIva_GotFocus()
 ComboCategIva.SelLength = Len(ComboCategIva.Text)
End Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'TODO ESTO ES PARA ANIMAR EL COMBO ESTADO CIVIL

Private Sub txtEstCiv_gotfocus()
    CbEstCiv.Height = 1150
    CbEstCiv.Visible = True
    CbEstCiv.Text = txtEstCiv.Text
    CbEstCiv.SetFocus
    'noenter = True
End Sub

Private Sub cbestciv_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cbestciv_dblclick
    End If
End Sub

Private Sub cbestciv_dblclick()
    txtEstCiv.Text = CbEstCiv.Text
    'QueEstCiv = CbEstCiv.ListIndex
    CbEstCiv.Visible = False
    'noenter = False
    txtFecNac.SetFocus
End Sub

Private Sub cbestciv_lostfocus()
    CbEstCiv.Visible = False
End Sub




' POSICION DE BOTONES CON RESPECTO AL MOMENTO DE EDICION

Private Sub BotoneraNeutra()

cmdBuscar.Visible = True

CmdContinuar.Visible = False

CmdActualizar.Visible = False

cmdCancelar.Visible = False

CmdModificar.Visible = False

CmdEliminar.Visible = False


Command1.Visible = False


End Sub

Private Sub BotoneraExploracion()

cmdBuscar.Visible = False
CmdContinuar.Visible = True
CmdActualizar.Visible = False
cmdCancelar.Visible = False
CmdModificar.Visible = True
CmdEliminar.Visible = True
'CmdImprimir.Visible = True
Command1.Visible = True
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

Command1.Visible = True

CmdActualizar.SetFocus

End Sub

'########PROCEDIMEITNO IMPRIMIR IMPRIMIR ###################33

Private Sub cmdImprimir_Click()
'ImprimirFicha
End Sub


Private Sub optConsumos_Click()
VerOpcion
End Sub

Private Sub optCtaCte_Click()
VerOpcion
End Sub

Private Sub OptPresupuestos_Click()
VerOpcion
End Sub


Private Sub VerOpcion()
AbrirBase
If optConsumos.Value = True Then
VerConsumos
ElseIf optCtaCte.Value = True Then
VerCuentaCorriente
ElseIf OptPresupuestosValue = True Then
VerPresupuestos
End If
CerrarBase
End Sub

Private Sub VerConsumos()

titulos = "TIPO|Nº FACT.|FECHA|COND. VENTA|TOTAL"

 MSHFlexGrid1.Clear
    With MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .ColWidth(0) = 500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1100
        .ColWidth(3) = 2000
        .ColWidth(4) = 1000
        .ColWidth(5) = 0
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .ColAlignmentFixed = 3
    End With


TotVent = 0
TotCount = 0

TraerConsumoXComprobante "FACTURASA"
TraerConsumoXComprobante "FACTURASB"
TraerConsumoXComprobante "FACTURASC"
TraerConsumoXComprobante "FACTURASX"


lblTotalRef = TotCount
lblTotal = "TOTAL VENTAS: $ " & Format(TotVent, "#,###.#0")
End Sub

Private Sub TraerConsumoXComprobante(tipo As Variant)
Dim rsCons As New Recordset
strSql = "Select * From " & tipo & " Where Codcliente=" & Val(txtCodDist)
rsCons.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
While Not rsCons.EOF
                i = i + 1
                
                linea = rsCons!tipofact _
                & Chr(9) & rsCons!numfact _
                & Chr(9) & rsCons!Fecha _
                & Chr(9) & rsCons!condventa _
                & Chr(9) & Format(rsCons!total, "#,###.#0")

MSHFlexGrid1.AddItem linea, i
TotVent = TotVent + rsCons!total
rsCons.MoveNext
Wend
TotCount = TotCount + rsCons.RecordCount
rsCons.Close

End Sub


Private Sub VerCuentaCorriente()

titulos = "ID|Nº FACT.|FECHA|DESCRIPCIÓN|DEBE|HABER"

 MSHFlexGrid1.Clear
    With MSHFlexGrid1
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
strSql = "Select * From CuentasCorrientes Where Codcliente=" & Val(txtCodDist)
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

MSHFlexGrid1.AddItem linea, i
TotDebe = TotDebe + rsCons!Debe
TotHaber = TotHaber + rsCons!Haber
rsCons.MoveNext
Wend
lblTotalRef = rsCons.RecordCount



total = TotDebe - TotHaber

IIf (total < 0), lblTotal.ForeColor = &HFF&, lblTotal.ForeColor = &HFF0000

lblTotal = "Saldo Cta/Cte " & IIf((total < 0), "Negativo", "Positivo") & " $ " & Format(total, "standard")

End Sub



Private Sub VerPresupuestos()

titulos = " Nº Pres|Fecha|VENDEDOR|CLIENTE|TOTAL"

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
        .ColWidth(1) = 1000
        .ColWidth(2) = 0
        .ColWidth(3) = 2000
        .ColWidth(4) = 1500
        .ColWidth(5) = 0
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .ColAlignmentFixed = 3
    End With



Dim rsCons As New Recordset
Dim strsConsql As String
strsConsql = "SELECT * FROM Presupuestos WHERE CodCliente=" & Val(txtCodDist)
rsCons.Open strsConsql, DB, adOpenKeyset, adLockOptimistic, adCmdText

'Dim totvent As Double
TotVent = "0"

While Not rsCons.EOF
                i = i + 1
                
                'com = 0
                'com = (rsCons!TotalVenta * Val(txtComision) / 100)
                linea = rsCons!NumPres _
                & Chr(9) & rsCons!Fecha _
                & Chr(9) & rsCons!CodUsuario _
                & Chr(9) & rsCons!CodCliente _
                & Chr(9) & Format(rsCons!total, "#,###.#0") _

MSHFlexGrid1.AddItem linea, i
TotVent = TotVent + rsCons!total
'totComp = totComp + rsCons!TotalCompra
'totUt = totUt + rsCons!Ganancia
rsCons.MoveNext
Wend
lblTotalRef = rsCons.RecordCount
lblTotal = "TOTAL Presup: $ " & Format(TotVent, "#,###.#0")
'lblCosto = "COSTO MERCADERIA: $ " & Format(totComp, "#,###.#0")
'lblGanancia = "GANANCIA: $ " & Format(totUt, "#,###.#0")
End Sub


Private Sub cmdInputPago_Click()
MSHFlexGrid1.Col = 0
If IsNumeric(MSHFlexGrid1.Text) Then
MSHFlexGrid1.Col = 5
If Me.optCtaCte.Value = True And Val(MSHFlexGrid1.Text) > 0 Then
frmInputPago.Show
frmInputPago.lblID = Val(txtCodDist)
frmInputPago.lblNombre = txtNombre
MSHFlexGrid1.Col = 1
frmInputPago.lblNumFact = Val(MSHFlexGrid1.Text)
MSHFlexGrid1.Col = 2
frmInputPago.lblFecha = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 5
frmInputPago.lblTotal = MSHFlexGrid1.Text
frmInputPago.lblSaldo = lblTotal
frmInputPago.dtp1 = Date
frmInputPago.txtImporte = ""
End If
End If

End Sub



Private Sub MSHFlexGrid1_DblClick()
If Me.optCtaCte.Value = True Then
    MSHFlexGrid1.Col = 0
    If IsNumeric(MSHFlexGrid1.Text) Then
        If MsgBox("¿Elimina este Item?", vbExclamation + vbYesNoCancel) = vbYes Then
        EliminarItem
        End If
    End If

ElseIf Me.optConsumos.Value = True Then

    MSHFlexGrid1.Col = 0
    If IsNumeric(MSHFlexGrid1.Text) Then
        If MsgBox("¿Elimina este Item?", vbExclamation + vbYesNoCancel) = vbYes Then
        EliminarVenta
        End If
    End If


ElseIf Me.OptPresupuestos.Value = True Then

    MSHFlexGrid1.Col = 0
    If IsNumeric(MSHFlexGrid1.Text) Then
        If MsgBox("¿Elimina este Item?", vbExclamation + vbYesNoCancel) = vbYes Then
        EliminarPresupuesto
        End If
    End If

End If
End Sub

Private Sub EliminarItem()
Dim itemselecto As Integer
Dim strSql As String
AbrirBase
MSHFlexGrid1.Col = 0
itemselecto = Val(MSHFlexGrid1.Text)

Dim rsCarrito As New Recordset
strSql = "SELECT * FROM CuentasCorrientes Where ID=" & Val(itemselecto)
rsCarrito.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rsCarrito.BOF And rsCarrito.EOF) Then
rsCarrito.Delete
rsCarrito.Update
End If

CerrarBase
VerOpcion
End Sub


Private Sub EliminarVenta()
Dim rsCarrito As New Recordset
Dim itemselecto As Integer
Dim strSql As String
AbrirBase

MSHFlexGrid1.Col = 0
itemselecto = MSHFlexGrid1.Text

Dim rsActualizacionStock As New Recordset
'Dim strSql As String

MSHFlexGrid1.Col = 2
cobselecto = MSHFlexGrid1.Text


v = "select * from FACTURASX where numfact=" & Val(itemselecto)
rs.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rs.BOF And rs.EOF) Then
'While Not rs.EOF
'rs!CondVenta = "ANULADO"
rs.Delete
rs.Update
'Wend
End If

Dim rs1 As New Recordset
v = "select * from DETALLEFACTURASX where numfact=" & Val(itemselecto)
rs1.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rs1.BOF And rs1.EOF) Then
While Not rs1.EOF
rs1.Delete
rs1.MoveNext
Wend
End If

Dim rs2 As New Recordset
v = "select * from CuentasCorrientes where NumFact=" & Val(itemselecto)
rs2.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rs2.BOF And rs2.EOF) Then
While Not rs2.EOF
rs2.Delete
rs2.MoveNext
Wend
End If
CerrarBase
VerOpcion
End Sub

Private Sub EliminarPresupuesto()

AbrirBase
MSHFlexGrid1.Col = 0
itemselecto = MSHFlexGrid1.Text
Dim rs2 As New Recordset
v = "select * from Presupuestos where NumPres=" & Val(itemselecto)
rs2.Open (v), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rs2.BOF And rs2.EOF) Then
While Not rs2.EOF
rs2.Delete
rs2.MoveNext
Wend
End If
CerrarBase
VerOpcion
End Sub

Private Sub ImprimirFicha()

If IsNumeric(txtCodDist) Then
AbrirBase
ImprimirSeleccion
CerrarBase
Else
MsgBox "Debe seleccionar al menos un Cliente para poder imprimir", vbExclamation
End If
End Sub

Private Sub ImprimirSeleccion()

    
On Error Resume Next
    
    ContRay = 0
    ContLin = 0
    espacioCelda = 0.4
    
    ImpTitulos
    

    
If Option2.Value = True Then
PrintCtaCte
ElseIf Option1.Value = True Then
PrintFacturas
ElseIf Option3.Value = True Then
PrintPresupuestos
End If
End Sub

Private Sub ImpTitulos()


If Option2.Value = True Then
Titulo = "Resumen de Cta.Cte. de  " & txtNombre & " " & "(" & txtCodDist & ")"
ElseIf Option1.Value = True Then
Titulo = "Resumen de Facturas de  " & txtNombre & " " & "(" & txtCodDist & ")"
ElseIf Option3.Value = True Then
Titulo = "Resumen de Presupuestos de  " & txtNombre & " " & "(" & txtCodDist & ")"
End If
    'Imprime el Template de la hoja
Printer.ScaleMode = vbCentimeters
Printer.PaperSize = 9
Printer.Orientation = 1
Printer.FontName = "ARIAL"
Printer.DrawStyle = 0
Printer.Font.Size = 14


'Printer.PaintPicture frmLogo.Picture1.Picture, 1, 1

'Printer.CurrentX = 1
'Printer.CurrentY = 1
''Printer.Print "DYVCOM S.R.L."
Printer.Font.Size = 8

Printer.CurrentX = 18.8
Printer.CurrentY = 1
Printer.Print "Página " & Printer.Page


Printer.CurrentX = 18.5
Printer.CurrentY = 1.5
Printer.Print Date '& " " & Format(Time, "hh:nn am/pm")

Printer.Font.Size = 14

x = 21 / 2
Printer.CurrentX = x - (Printer.TextWidth(Titulo)) + (Printer.TextWidth(Titulo)) / 2
Printer.CurrentY = 1.5
Printer.Print Titulo

Printer.Font.Size = 8


' DETALLE DEL CLIENTE

Printer.CurrentY = 2.5
Printer.CurrentX = 1

Printer.Print "Cliente: " & txtNombre & " " & "(" & txtCodDist & ")";

'Printer.CurrentX = 12
'Printer.Print "CUIT Nº: " & txtCuit;

Printer.CurrentX = 15
Printer.Print "Fecha de Alta: " & txtFechaAlta

Printer.CurrentY = 3
Printer.CurrentX = 1
Printer.Print "Domicilio: " & txtDomicilio & " " & "(" & txtCodPost & ")" & " " & "-" & txtLocalidad & "-";
Printer.CurrentX = 15
Printer.Print "Telefono: " & txtTelefono

Printer.CurrentY = 3.5
Printer.CurrentX = 1
Printer.Print "E-mail: " & txtEmail;

Printer.CurrentX = 15
Printer.Print "Categ IVA: " & ComboCategIva.Text


Printer.Line (1, (Printer.CurrentY + 0.8))-(20, (Printer.CurrentY + 0.8))


Printer.CurrentY = Printer.CurrentY + espacioCelda


If Option2.Value = True Then


Printer.CurrentX = 1
Printer.Print "Nº Factura";

Printer.CurrentX = 2.8
Printer.Print "Fecha";

Printer.CurrentX = 4.5
Printer.Print "DESCRIPCION";

Printer.CurrentX = 14
Printer.Print "DEBE";

Printer.CurrentX = 16
Printer.Print "HABER"


ElseIf Option1.Value = True Then


Printer.CurrentX = 1
Printer.Print "Nº Fact";

Printer.CurrentX = 2.8
Printer.Print "Fecha";

Printer.CurrentX = 4.5
Printer.Print "COND. VENTA";

Printer.CurrentX = 16
Printer.Print "TOTAL"

ElseIf Option3.Value = True Then

Printer.CurrentX = 1
Printer.Print "Nº Pres.";

Printer.CurrentX = 2.8
Printer.Print "Fecha";

Printer.CurrentX = 4.5
Printer.Print "COD. CLIENTE";

Printer.CurrentX = 16
Printer.Print "TOTAL"


End If
Printer.Line (1, (Printer.CurrentY + espacioCelda))-(20, (Printer.CurrentY + espacioCelda))

End Sub

Private Sub PrintCtaCte()
strSql = "Select * From CuentasCorrientes Where CodCliente=" & Val(txtCodDist)
Dim rs1 As New Recordset
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText


Do While Not rs1.EOF
            
If ContRay > 0 Then
Printer.CurrentY = Printer.CurrentY + espacioCelda
End If
                     
Dim CadenaCod As String
Dim LongitudCod As Long
CadenaCod = Format(CadenaCod, "0####")
LongitudCod = Len(CadenaCod)

Printer.CurrentX = 2 + (Val(LongitudCod)) - (Printer.TextWidth(Format(rs1!numfact, "0####")))
Printer.Print Format(rs1!numfact, "0####");


       
       Printer.CurrentX = 2.8
       Printer.Print Left(rs1!Fecha, 22);
       
       
       
       
    Printer.CurrentX = 4.5
       Printer.Print Left(rs1!Descripcion, 30);

       
'       Printer.CurrentX = 11.5
'       Printer.Print Left(RstSocios!Telefono, 22);
'
'       Printer.CurrentX = 13.7
'       Printer.Print Left(RstSocios!Localidad, 22);
'
'        Printer.CurrentX = 17.5
'        Printer.Print Left(RstSocios!Obs, 22);




Dim Cadena2 As String
Dim Longitud2 As Long
Cadena2 = Format(Cadena2, "#,###.#0")
Longitud2 = Len(Cadena2)

Printer.CurrentX = 15 + (Val(Longitud2)) - (Printer.TextWidth(Format(rs1!Debe, "#,###.#0")))
Printer.Print Format(rs1!Debe, "#,###.#0");


Dim Cadena3 As String
Dim Longitud3 As Long
Cadena3 = Format(Cadena3, "#,###.#0")
Longitud3 = Len(Cadena3)

Printer.CurrentX = 17 + (Val(Longitud3)) - (Printer.TextWidth(Format(rs1!Haber, "#,###.#0")))
Printer.Print Format(rs1!Haber, "#,###.#0")

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
       rs1.MoveNext
   Loop

Printer.Line (1, Printer.CurrentY + espacioCelda)-(20, Printer.CurrentY + espacioCelda)
Printer.Font.Size = 12
Printer.CurrentY = Printer.CurrentY + espacioCelda
Printer.CurrentX = 14
Printer.Print lblTotal
Printer.EndDoc

End Sub


Private Sub PrintFacturas()

strSql = "Select * From FACTURASX Where Codcliente=" & Val(txtCodDist)
Dim rs2 As New Recordset
rs2.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

Do While Not rs2.EOF
            
If ContRay > 0 Then
Printer.CurrentY = Printer.CurrentY + espacioCelda
End If
                     
Dim CadenaCod As String
Dim LongitudCod As Long
CadenaCod = Format(CadenaCod, "0####")
LongitudCod = Len(CadenaCod)

Printer.CurrentX = 2 + (Val(LongitudCod)) - (Printer.TextWidth(Format(rs2!numfact, "0####")))
Printer.Print Format(rs2!numfact, "0####");


       
       Printer.CurrentX = 2.8
       Printer.Print Left(rs2!Fecha, 22);
       
       
       
       
    Printer.CurrentX = 4.5
       Printer.Print Left(rs2!condventa, 30);

       
'       Printer.CurrentX = 11.5
'       Printer.Print Left(RstSocios!Telefono, 22);
'
'       Printer.CurrentX = 13.7
'       Printer.Print Left(RstSocios!Localidad, 22);
'
'        Printer.CurrentX = 17.5
'        Printer.Print Left(RstSocios!Obs, 22);




'Dim Cadena2 As String
'Dim Longitud2 As Long
'Cadena2 = Format(Cadena2, "#,###.#0")
'Longitud2 = Len(Cadena2)
'
'Printer.CurrentX = 15 + (Val(Longitud2)) - (Printer.TextWidth(Format(rs2!Debe, "#,###.#0")))
'Printer.Print Format(rs2!Debe, "#,###.#0");


Dim Cadena3 As String
Dim Longitud3 As Long
Cadena3 = Format(Cadena3, "#,###.#0")
Longitud3 = Len(Cadena3)

Printer.CurrentX = 17 + (Val(Longitud3)) - (Printer.TextWidth(Format(rs2!total, "#,###.#0")))
Printer.Print Format(rs2!total, "#,###.#0")

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
       rs2.MoveNext
   Loop

Printer.Line (1, Printer.CurrentY + espacioCelda)-(20, Printer.CurrentY + espacioCelda)
Printer.Font.Size = 12
Printer.CurrentY = Printer.CurrentY + espacioCelda
Printer.CurrentX = 14
Printer.Print lblTotal
Printer.EndDoc

End Sub

Private Sub PrintPresupuestos()

strSql = "SELECT * FROM Presupuestos WHERE CodCliente=" & Val(txtCodDist)

Dim rs3 As New Recordset
rs3.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

Do While Not rs3.EOF
            
If ContRay > 0 Then
Printer.CurrentY = Printer.CurrentY + espacioCelda
End If
                     
Dim CadenaCod As String
Dim LongitudCod As Long
CadenaCod = Format(CadenaCod, "0####")
LongitudCod = Len(CadenaCod)

Printer.CurrentX = 2 + (Val(LongitudCod)) - (Printer.TextWidth(Format(rs3!NumPres, "0####")))
Printer.Print Format(rs3!NumPres, "0####");
     
Printer.CurrentX = 2.8
Printer.Print Left(rs3!Fecha, 22);
       
Printer.CurrentX = 4.5
Printer.Print Left(rs3!CodCliente, 30);

       
'       Printer.CurrentX = 11.5
'       Printer.Print Left(RstSocios!Telefono, 22);
'
'       Printer.CurrentX = 13.7
'       Printer.Print Left(RstSocios!Localidad, 22);
'
'        Printer.CurrentX = 17.5
'        Printer.Print Left(RstSocios!Obs, 22);




'Dim Cadena2 As String
'Dim Longitud2 As Long
'Cadena2 = Format(Cadena2, "#,###.#0")
'Longitud2 = Len(Cadena2)
'
'Printer.CurrentX = 15 + (Val(Longitud2)) - (Printer.TextWidth(Format(rs3!Debe, "#,###.#0")))
'Printer.Print Format(rs3!Debe, "#,###.#0");


Dim Cadena3 As String
Dim Longitud3 As Long
Cadena3 = Format(Cadena3, "#,###.#0")
Longitud3 = Len(Cadena3)

Printer.CurrentX = 17 + (Val(Longitud3)) - (Printer.TextWidth(Format(rs3!total, "#,###.#0")))
Printer.Print Format(rs3!total, "#,###.#0")

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
       rs3.MoveNext
   Loop

Printer.Line (1, Printer.CurrentY + espacioCelda)-(20, Printer.CurrentY + espacioCelda)
Printer.Font.Size = 12
Printer.CurrentY = Printer.CurrentY + espacioCelda
Printer.CurrentX = 14
Printer.Print lblTotal
Printer.EndDoc




End Sub
