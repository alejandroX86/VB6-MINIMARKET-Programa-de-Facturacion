VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpresa 
   BackColor       =   &H00000000&
   Caption         =   "Datos de La Empresa"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7125
   Icon            =   "frmEmpresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos de la Empresa"
      Height          =   4395
      Left            =   180
      TabIndex        =   8
      Top             =   180
      Width           =   6735
      Begin MSComCtl2.DTPicker dtpFechaVto 
         Height          =   375
         Left            =   2400
         TabIndex        =   22
         Top             =   3840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   141164545
         CurrentDate     =   40420
      End
      Begin MSComCtl2.DTPicker dtpInicioActividades 
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   3000
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   141164545
         CurrentDate     =   40420
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   0
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtSucursal 
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtCuit 
         Height          =   285
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtIngresosBrutos 
         Height          =   285
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2640
         Width           =   2535
      End
      Begin VB.ComboBox ComboCategIva 
         Height          =   315
         ItemData        =   "frmEmpresa.frx":57E2
         Left            =   2400
         List            =   "frmEmpresa.frx":57F2
         TabIndex        =   4
         Text            =   "IVA Responsable Inscripto"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton cmdGuardarEmpresa 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   5220
         TabIndex        =   9
         Top             =   3780
         Width           =   1215
      End
      Begin VB.TextBox txtCai 
         Height          =   285
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   7
         Top             =   3420
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre de la Empresa:"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Dirección y Código Postal:"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Sucursal:"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Teléfono:"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Categ. Iva:"
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Cuit Nº:"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Ingresos Brutos Nº:"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Inicio de Actividades:"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   3060
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "C.A.I. Nº:"
         Height          =   255
         Left            =   1380
         TabIndex        =   11
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha Vto:"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   3840
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
Dim emp As String
emp = "SELECT * FROM EMPRESA WHERE IdEmpresa = 1"
AbrirBase
rsEmpresa.Open (emp), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rsEmpresa.BOF And rsEmpresa.EOF) Then
txtNombre = rsEmpresa!nombre
txtDireccion = rsEmpresa!Direccion
txtSucursal = rsEmpresa!Sucursal
txtTelefono = rsEmpresa!telefono
ComboCategIva = rsEmpresa!categiva
txtCuit = rsEmpresa!cuit
txtIngresosBrutos = rsEmpresa!IngresosBrutos
dtpInicioActividades = rsEmpresa!InicioActividades
txtCai = rsEmpresa!Cai
dtpFechaVto = rsEmpresa!FechaVto
End If
CerrarBase
End Sub
Private Sub cmdGuardarEmpresa_Click()
Dim emp As String
emp = "SELECT * FROM EMPRESA WHERE IdEmpresa = 1"
AbrirBase
rsEmpresa.Open (emp), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rsEmpresa.BOF And rsEmpresa.EOF) Then
rsEmpresa!nombre = txtNombre
rsEmpresa!Direccion = txtDireccion
rsEmpresa!Sucursal = txtSucursal
rsEmpresa!telefono = txtTelefono
rsEmpresa!categiva = ComboCategIva
rsEmpresa!cuit = txtCuit
rsEmpresa!IngresosBrutos = txtIngresosBrutos
rsEmpresa!InicioActividades = dtpInicioActividades
rsEmpresa!Cai = txtCai
rsEmpresa!FechaVto = dtpFechaVto
rsEmpresa.Update
MsgBox "Datos de Empresa Exitosamente Registrados"
End If
CerrarBase
End Sub
