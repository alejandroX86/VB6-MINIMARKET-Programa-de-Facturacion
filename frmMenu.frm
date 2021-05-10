VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00000000&
   Caption         =   "  MINIMARKET    avrmicrorobot@gmail.com   Omar Alejandro  Bazar"
   ClientHeight    =   7890
   ClientLeft      =   885
   ClientTop       =   435
   ClientWidth     =   10245
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Caption         =   "Cambiar Clave"
      Height          =   300
      Left            =   3510
      TabIndex        =   16
      Top             =   8550
      Width           =   3195
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Listados:"
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
      Height          =   2265
      Left            =   6840
      TabIndex        =   21
      Top             =   5520
      Width           =   3365
      Begin VB.CommandButton Command24 
         BackColor       =   &H0000FFFF&
         Caption         =   "Ventas en cta.cte."
         Height          =   300
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1320
         Width           =   1485
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H0000FFFF&
         Caption         =   "Ventas X Cliente"
         Height          =   300
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   1485
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0000FFFF&
         Caption         =   "Listado de Clientes"
         Height          =   300
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1485
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0000FFFF&
         Caption         =   "Vtas X Factura"
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   1485
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H0000FFFF&
         Caption         =   "Vtas X Producto"
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   1485
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H0000FFFF&
         Caption         =   "Valuación Stock"
         Height          =   300
         Left            =   1665
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Control:"
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
      Height          =   2265
      Left            =   0
      TabIndex        =   20
      Top             =   5520
      Width           =   3465
      Begin VB.CommandButton Command26 
         BackColor       =   &H0000FF00&
         Caption         =   "Asignar Alicuotas"
         Height          =   330
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdImpuestos 
         BackColor       =   &H0000FF00&
         Caption         =   "Impuestos"
         Height          =   330
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H0000FF00&
         Caption         =   "Cordenadas"
         Height          =   330
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H0000FF00&
         Caption         =   "Impresoras"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H0000FF00&
         Caption         =   "Proveedores"
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1485
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "Articulos"
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1485
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "Clientes"
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   1485
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Caption         =   "Empresa"
         Height          =   300
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   600
         Width           =   1485
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0000FF00&
         Caption         =   "Vendedores"
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   1485
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0000FF00&
         Caption         =   "Usuarios"
         Height          =   300
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Comprobantes:"
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
      Height          =   2265
      Left            =   3480
      TabIndex        =   19
      Top             =   5520
      Width           =   3345
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFF00&
         Caption         =   "Codigos de Barra"
         Height          =   285
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Presupuestos"
         Height          =   300
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   1485
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFF00&
         Caption         =   "Ventas"
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1485
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFF00&
         Caption         =   "Pedidos"
         Height          =   300
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Cerrar Sesión"
      Height          =   300
      Left            =   90
      TabIndex        =   15
      Top             =   8550
      Width           =   3240
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   300
      Left            =   6885
      TabIndex        =   17
      Top             =   8550
      Width           =   3330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Portafolio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   7440
      TabIndex        =   28
      Top             =   4680
      Width           =   2745
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "avrmicrorobot@gmail.com"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   465
      Left            =   6000
      TabIndex        =   27
      Top             =   5040
      Width           =   4185
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "lblUsuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   360
      Left            =   180
      TabIndex        =   18
      Top             =   5670
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   6240
      Left            =   0
      Picture         =   "frmMenu.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Power By = visualbasicseis@yahoo.com.ar

Private Sub cmdImportarExcel_Click()
frmImportarExcel.Show
End Sub

Private Sub cmdImpuestos_Click()
ListarObjeto "Impuestos"
End Sub

Private Sub Command19_Click()

End Sub

Private Sub Command20_Click()

End Sub

Private Sub Command26_Click()
frmListConectObject.Show
End Sub

Private Sub Command25_Click()
ListarObjeto "Articulos"
End Sub

Private Sub ListarObjeto(varObject As String)
ListObject frmListObject.flex01, varObject, frmListObject.Image1, frmListObject.Image2
frmListObject.Caption = varObject
frmListObject.lblObject = varObject
frmListObject.Width = frmListObject.flex01.Width + 300
frmListObject.Show
End Sub

Private Sub Command1_Click()
frmPresupuestos.Show
End Sub

Private Sub Command10_Click()
frmVentasPorComprob.Show
End Sub

Private Sub Command11_Click()
frmUsuarios.Show
End Sub

Private Sub Command12_Click()
frmVentasPorArticulo.Show
End Sub

Private Sub Command13_Click()
frmCambioClave.Show
End Sub

Private Sub Command14_Click()
frmLogin.Show
Unload Me
End Sub

Private Sub Command15_Click()
frmValuacion.Show
End Sub

Private Sub Command16_Click()
frmBarcode.Show
End Sub

Private Sub Command17_Click()
frmProveedores.Show
End Sub

Private Sub Command2_Click()
If lblUsuario = "Usuario del Sistema" Then
frmArticulosUser.Show
Else
frmArticulos.Show
End If
End Sub

Private Sub Command21_Click()
frmVentasPorCliente.Show

End Sub

Private Sub Command22_Click()
frmConfigFacturacion.Show

End Sub

Private Sub Command23_Click()
frmCordPrint.Show
End Sub

Private Sub Command24_Click()
frmVentasPorCtaCte.Show
End Sub

Private Sub Command3_Click()
frmClientes.Show
End Sub

Private Sub Command4_Click()
frmEmpresa.Show
End Sub

Private Sub Command5_Click()
frmListadoClientes.Show
End Sub

Private Sub Command6_Click()
frmVendedores.Show

'ListarObjeto "Vendedores"

End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()
frmFacturacion.Show
End Sub

Private Sub Command9_Click()
frmOrdenPedido.Show
End Sub


