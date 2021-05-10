VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPresupuestos 
   BackColor       =   &H00000000&
   Caption         =   "PRESUPUESTOS"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   Icon            =   "frmPresupuestos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Opciones de Impresión:"
      Height          =   795
      Left            =   0
      TabIndex        =   53
      Top             =   5220
      Width           =   5655
      Begin VB.ComboBox ComboCopias 
         Height          =   315
         ItemData        =   "frmPresupuestos.frx":0442
         Left            =   3900
         List            =   "frmPresupuestos.frx":0455
         TabIndex        =   55
         Text            =   "1"
         Top             =   300
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Imprimir Sin Membrete"
         Height          =   315
         Left            =   180
         TabIndex        =   54
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Copias:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   56
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Artículo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   5760
      TabIndex        =   32
      Top             =   660
      Width           =   6435
      Begin VB.TextBox txtTalle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4455
         MaxLength       =   50
         TabIndex        =   58
         Top             =   180
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtMarca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   57
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Limpiar"
         Height          =   255
         Left            =   4800
         TabIndex        =   51
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregarItem 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   2640
         TabIndex        =   50
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminarItem 
         Caption         =   "Eliminar"
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarArticulo 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarArticulo 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtCodArticulo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0"
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox txtCantidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         MaxLength       =   5
         TabIndex        =   12
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtPrecio 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   840
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Bindings        =   "frmPresupuestos.frx":0468
         Height          =   3435
         Left            =   120
         TabIndex        =   52
         Top             =   1560
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   6059
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12632064
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   3600
         TabIndex        =   47
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio:"
         Height          =   255
         Left            =   480
         TabIndex        =   46
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Descripción:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CodArt:"
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Subtotal Sin IVA:"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Total:"
         Height          =   255
         Left            =   4320
         TabIndex        =   42
         Top             =   5340
         Width           =   495
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0E0FF&
         Caption         =   "IVA 21%"
         Height          =   255
         Left            =   2160
         TabIndex        =   41
         Top             =   5280
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Encabezado de Presupuesto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   23
      Top             =   660
      Width           =   5655
      Begin VB.TextBox txtHora 
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNumFact 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nº Presupuesto:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   17
      Top             =   2820
      Width           =   5655
      Begin VB.CommandButton cmdAceptarCliente 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   600
         TabIndex        =   49
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardarCliente 
         Caption         =   "Guardar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelarCliente 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3480
         TabIndex        =   29
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtCodCliente 
         Height          =   285
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   6
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtDomicilio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   7
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtTelefono 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox ComboCategIva 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPresupuestos.frx":047E
         Left            =   1320
         List            =   "frmPresupuestos.frx":0494
         TabIndex        =   9
         Text            =   "IVA Consumidor Final"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Categ. IVA:"
         Height          =   255
         Left            =   480
         TabIndex        =   48
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Señor/es:"
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ID:"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cuit:"
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Tel:"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Vendedor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   16
      Top             =   1740
      Width           =   5655
      Begin VB.CommandButton cmdAceptarUsuario 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelarUsuario 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4320
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtUsuario 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtCodUsuario 
         Height          =   285
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cod. Vendedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   6465
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "15/09/10"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "09:40"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1005
      ButtonWidth     =   2963
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo Presupuesto"
            Key             =   "ToolNuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "ToolPrint"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar Presupuesto"
            Key             =   "ToolCancelar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo Cliente"
            Key             =   "ToolNuevoCliente"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar Cliente"
            Key             =   "ToolBuscarCliente"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "--"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar Artículo"
            Key             =   "ToolBuscarArticulo"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestos.frx":0518
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestos.frx":062C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestos.frx":0740
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestos.frx":0854
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestos.frx":0968
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestos.frx":0A7C
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
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNuevaFactura 
         Caption         =   "&Nuevo Presupuesto"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuCancelarFactura 
         Caption         =   "&Cancelar Presupuesto"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu A 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuBuscar 
      Caption         =   "&Buscar"
      Begin VB.Menu mnuBuscarCliente 
         Caption         =   "&Buscar Cliente"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuBuscarArticulo 
         Caption         =   "&Buscar Artículo"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuBuscarFactura 
         Caption         =   "&Buscar Factura"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu Arqueo 
         Caption         =   "&Arqueo de Caja"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmPresupuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsClientes As New Recordset
Dim rsPresupuestos As New Recordset
Dim rsUsuarios As New Recordset
Dim rsArticulos As New Recordset
Dim rsTotales As New Recordset
Dim rsCarrito As New Recordset
Dim rsEmpresa As New Recordset
Dim rsPrintItem As New Recordset


'#########################  SECCION ABRIR/CERRAR BASE  ###########################


Private Sub Arqueo_Click()
frmArqueo.Show
End Sub

Private Sub mnuBuscarFactura_Click()
frmBuscarFactura.Show
End Sub

'&&&&&&&&&&&&&&&&&& MENUES &&&&&&&&&&&&&&&&&&&&&&&

Private Sub mnuNuevaFactura_Click()
NuevaFactura
End Sub

Private Sub mnuCancelarFactura_Click()
CancelarFactura
End Sub
Private Sub mnuImprimir_Click()
ImprimirFactura
End Sub
Private Sub mnuSalir_Click()
frmFacturacion.Hide
End Sub
Private Sub mnuBuscarArticulo_Click()
frmBuscarArticulo.Show
End Sub
Private Sub mnuBuscarCliente_Click()
frmBuscarCliente.Show
End Sub



'&&&&&&&&&&&&&&&&&& Barra de Herramientas &&&&&&&&&&&&&&&&&

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
Case "ToolNuevo"
NuevaFactura
Case "ToolPrint"
ImprimirFactura
Case "ToolCancelar"
CancelarFactura
Case "ToolNuevoCliente"
DeshabilitarCliente
HabilitarCliente
NuevoCliente
Case "ToolBuscarCliente"
frmBuscarClienteF.Show
Case "ToolBuscarArticulo"
frmBuscarArticulo.Show
End Select
End Sub



'#########################  SECCION ENCABEZADO FACTURAS   ###########################



Private Sub txtNumFact_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        NuevaFactura
    End If
End Sub


Private Sub NuevaFactura()
VaciarCarrito
AveriguarNumero
DeshabilitarArticulo
DeshabilitarCliente
DeshabilitarUsuario
End Sub

Private Sub CancelarFactura()
VaciarCarrito
DeshabilitarArticulo
DeshabilitarCliente
DeshabilitarUsuario
txtFecha = ""
txtHora = ""
txtNumFact = ""
txtNumFact.SetFocus
End Sub

'#########################  NUEVA   ###########################


Private Sub AveriguarNumero()
Dim strSql As String
AbrirBase
        strSql = "SELECT * FROM Presupuestos"
        rsPresupuestos.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        If Not (rsPresupuestos.BOF And rsPresupuestos.EOF) Then
        rsPresupuestos.MoveLast
        txtNumFact = rsPresupuestos!NumPres
        txtNumFact = Val(txtNumFact) + 1
        txtFecha = Date
        txtHora = Time
        Else
        txtNumFact = Val(txtNumFact) + 1
        txtFecha = Date
        txtHora = Time
        End If
        
        
CerrarBase
End Sub


'#########################  SECCION USUARIO   ###########################


Private Sub txtCodUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCodUsuario.Text <> "" Then
        KeyAscii = 0
        AceptarUsuario txtCodUsuario.Text
    ElseIf KeyAscii = 8 And txtCodUsuario = "" Then
    txtNumFact.SetFocus
    KeyAscii = 0

    End If
End Sub


Private Sub AceptarUsuario(ByVal txtCodUsuario As Variant)

Dim strSql As String

AbrirBase

    If IsNumeric(txtCodUsuario) Then
         strSql = "SELECT * FROM Vendedores Where ID=" & Val(txtCodUsuario)

        rsUsuarios.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText

            If Not (rsUsuarios.BOF And rsUsuarios.EOF) Then
            ' Si existe
            CargarUsuario
            HabilitarUsuario

            Else
            MsgBox "Código de vendedor incorrecto"
            End If

    Else

        MsgBox "Código de vendedor incorrecto"

    End If

CerrarBase

End Sub

Private Sub CargarUsuario()

    txtUsuario.Text = rsUsuarios!nombre

End Sub

Private Sub HabilitarUsuario()
    ' Deshabilito el campo clave
    txtCodUsuario.Enabled = False
    ' Habilito todos los demás campos
    txtUsuario.Enabled = True
    'También habilito los botones correspondientes
    cmdCancelarUsuario.Enabled = True
    cmdAceptarUsuario.Enabled = False
    txtCodCliente.SetFocus

End Sub

Private Sub DeshabilitarUsuario()

    txtCodUsuario = ""
    txtUsuario = ""

    ' Hablilito el campo clave
    txtCodUsuario.Enabled = True
    ' Deshabilito todos los demás campos
    txtUsuario.Enabled = False

    'También habilito los botones correspondientes
    cmdCancelarUsuario.Enabled = False
    cmdAceptarUsuario.Enabled = True
    txtCodUsuario.SetFocus
End Sub


Private Sub cmdAceptarUsuario_Click()
AceptarUsuario txtCodUsuario.Text
End Sub



Private Sub cmdCancelarUsuario_Click()
DeshabilitarUsuario
End Sub
















'#########################  SECCION CLIENTE  ###########################





Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCodCliente.Text <> "" Then
        KeyAscii = 0
        AceptarCliente txtCodCliente.Text
    ElseIf KeyAscii = 8 And txtCodCliente = "" Then
    DeshabilitarUsuario
    KeyAscii = 0
    End If
End Sub

Private Sub NuevoCliente()
Dim strSql As String
AbrirBase
        strSql = "SELECT * FROM Clientes"
        rsClientes.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not (rsClientes.BOF And rsClientes.EOF) Then
        rsClientes.MoveLast
        txtCodCliente = rsClientes!id
        txtCodCliente = Val(txtCodCliente) + 1
        Else
        txtCodCliente = "1"
        End If
        
        
CerrarBase
End Sub
Private Sub AceptarCliente(ByVal txtCodCliente As Variant)

Dim strSql As String

AbrirBase

    If IsNumeric(txtCodCliente) Then
         strSql = "SELECT * FROM Clientes " & _
                 "WHERE ID =" & txtCodCliente
        
        rsClientes.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
            
            ' Si existe
            If Not (rsClientes.BOF And rsClientes.EOF) Then
            CargarCliente
            'cmdEliminar.Enabled = True
            HabilitarCliente
            Else
            MsgBox "Código de cliente incorrecto"
            
            End If

    Else
        
        MsgBox "Código de cliente incorrecto"
     
    End If
       CerrarBase
End Sub



Private Sub CargarCliente()
On Error Resume Next


    txtNombre = rsClientes!nombre
    txtDomicilio = rsClientes!domicilio
    txtTelefono = rsClientes!telefono
    'ComboCategIva = rsClientes!CategIva
    txtCuit = rsClientes!cuit
    
End Sub

Private Sub HabilitarCliente()
    ' Deshabilito el campo clave
    txtCodCliente.Enabled = False
    ' Habilito todos los demás campos
    
    txtNombre.Enabled = True
    txtDomicilio.Enabled = True
    txtTelefono.Enabled = True
    ComboCategIva.Enabled = True
    txtCuit.Enabled = True
    txtCodarticulo.Enabled = True
    'habilito los botones cancelar y guardar
    cmdGuardarCliente.Enabled = True
    cmdCancelarCliente.Enabled = True
    
    ' Y deshabilito el boton aceptar
    cmdAceptarCliente.Enabled = False
    
    txtNombre.SetFocus
    
End Sub


Private Sub DeshabilitarCliente()

'Primero Vacio los controles

txtCodCliente = ""
txtNombre = ""
txtDomicilio = ""
txtTelefono = ""
ComboCategIva = "IVA Consumidor Final"
txtCuit = ""

'Luego los deshabilito
txtNombre.Enabled = False
txtDomicilio.Enabled = False
txtTelefono.Enabled = False
ComboCategIva.Enabled = False
txtCuit.Enabled = False
txtCodarticulo.Enabled = False

'pero habilito el campo clave
txtCodCliente.Enabled = True
'Tambien desabilito los botones cancelar y guardar
cmdGuardarCliente.Enabled = False
cmdCancelarCliente.Enabled = False

'Y habilito el boton Aceptar
cmdAceptarCliente.Enabled = True
txtCodCliente.SetFocus

End Sub

Private Sub GuardarCliente()
Dim strSql As String

AbrirBase
strSql = "SELECT * FROM Clientes " & _
                 "WHERE ID =" & Val(txtCodCliente)
        
rsClientes.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
        
    'SI EXISTE
    If Not (rsClientes.BOF And rsClientes.EOF) Then
    rsClientes!nombre = txtNombre
    rsClientes!domicilio = txtDomicilio
    rsClientes!telefono = txtTelefono
    rsClientes!categiva = ComboCategIva
    rsClientes!cuit = txtCuit
    rsClientes.Update
    Else
    rsClientes.AddNew
    rsClientes!id = txtCodCliente
    rsClientes!nombre = txtNombre
    rsClientes!domicilio = txtDomicilio
    rsClientes!telefono = txtTelefono
    rsClientes!categiva = ComboCategIva
    rsClientes!cuit = txtCuit
    rsClientes.Update
    End If
CerrarBase

End Sub



Private Sub cmdAceptarCliente_Click()
AceptarCliente txtCodCliente.Text
End Sub

Private Sub cmdGuardarCliente_Click()
GuardarCliente
End Sub
Private Sub cmdCancelarCliente_Click()
DeshabilitarCliente
End Sub









'#########################  SECCION ARTICULO  ###########################


Private Sub txtCodArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCodarticulo.Text <> "" Then
    
    If InStr(1, txtCodarticulo, "%", vbTextCompare) > 0 Then
Dim x As String
x = txtCodarticulo
pp = InStr(1, x, "%", vbTextCompare)
txtCodarticulo = Replace(Mid(x, 1, pp), "%", "")
txtMarca = Mid(Replace(Mid(x, pp, Len(x)), "%", ""), 1, 2)
txtTalle = Mid(Replace(Mid(x, pp, Len(x)), "%", ""), 3, Len(x))
End If
        
    
    
    
        KeyAscii = 0
        AceptarPrecioCliente txtCodarticulo.Text
    ElseIf KeyAscii = 8 And txtCodarticulo = "" Then
    txtCuit.SetFocus
    KeyAscii = 0
    
    End If
End Sub

Private Sub AceptarPrecioCliente(ByVal txtCodarticulo As Variant)
Dim rsPreciosProv As New Recordset
Dim strSql As String

AbrirBase

    If IsNumeric(txtCodarticulo) Then
         strSql = "SELECT * FROM PreciosClientes " & _
                 "WHERE CodProveedor =" & Val(txtCodCliente) & _
                    " AND CodArticulo =" & Val(txtCodarticulo)

        rsPreciosProv.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
            
            If Not (rsPreciosProv.BOF And rsPreciosProv.EOF) Then
            ' Si existe
            txtDescripcion.Text = rsPreciosProv!Descripcion & " " & txtMarca & " " & txtTalle
            
            
            'txtPrecio.Text = Str(rsPreciosProv!Precio)
            
            
                'If Option1.Value = True Then
                txtPrecio.Text = Replace(Format(rsPreciosProv!Precio, "fixed"), ",", ".")
                'ElseIf Option2.Value = True Then
                'txtPrecio.Text = Replace(Format(rsPreciosProv!PrecioDocena, "fixed"), ",", ".")
                'End If
            
            
            HabilitarArticulo
            Else
            AceptarArticulo txtCodarticulo
            End If
    Else
        
        'MsgBox "Número de Artículo Incorrecto"
     
    End If

CerrarBase


End Sub

Private Sub AceptarArticulo(ByVal txtCodarticulo As Variant)
Dim strSql As String

'AbrirBase
    If IsNumeric(txtCodarticulo) Then
         strSql = "SELECT * FROM Articulos " & _
                 "WHERE ID =" & Val(txtCodarticulo)
        
        rsArticulos.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
            
            If Not (rsArticulos.BOF And rsArticulos.EOF) Then
            ' Si existe
            CargarArticulo
            HabilitarArticulo
            'cmdEliminar.Enabled = True
            Else
            MsgBox "Número de Artículo Incorrecto"
            End If
    Else
        MsgBox "Número de Artículo Incorrecto"
    End If
'CerrarBase
End Sub

Private Sub CargarArticulo()


'    lblCodInt = rsArticulos!CodInt
    txtDescripcion.Text = rsArticulos!Descripcion & " " & txtMarca & " " & txtTalle
    
    
    'If Option1.Value = True Then
    txtPrecio.Text = Replace(Format(rsArticulos!Precio, "fixed"), ",", ".")
    'ElseIf Option2.Value = True Then
    'txtPrecio.Text = Replace(Format(rsArticulos!PrecioDocena, "fixed"), ",", ".")
    'End If
    
End Sub

Private Sub HabilitarArticulo()
    ' Deshabilito el campo clave
    txtCodarticulo.Enabled = False
    ' Habilito todos los demás campos
    txtCantidad.Enabled = True
    txtPrecio.Enabled = True
   
    txtCantidad.SetFocus
    
End Sub

Private Sub DeshabilitarArticulo()
    'Primero Vacío los campos
    
    txtCodarticulo = ""
    txtDescripcion = ""
    txtPrecio = ""
    txtCantidad = ""
        txtMarca = ""
txtTalle = ""

    ' Deshabilito el campo clave
    txtCantidad.Enabled = False
    ' Habilito todos los demás campos
    txtCodarticulo.Enabled = True
    'Option1.Value = True
    txtCodarticulo.SetFocus
    
End Sub


Private Sub cmdAceptarArticulo_Click()
AceptarArticulo txtCodarticulo.Text
End Sub

Private Sub cmdCancelarArticulo_Click()
DeshabilitarArticulo
End Sub







'#########################  SECCION Cantidad  ###########################

'Private Sub Option1_Click()
'AceptarPrecioCliente txtCodArticulo.Text
'txtCantidad = "1"
'
'End Sub
'Private Sub Option2_Click()
'AceptarPrecioCliente txtCodArticulo.Text
'txtCantidad = "12"
'End Sub


Private Sub txtCantidad_GotFocus()
txtCantidad.SelLength = Len(txtCantidad)
End Sub



Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCantidad.Text <> "" Then
        KeyAscii = 0
        AceptarCantidad txtCantidad.Text
    
    ElseIf KeyAscii = 8 And txtCantidad = "" Then
    DeshabilitarArticulo
    KeyAscii = 0
    End If
End Sub

Private Sub AceptarCantidad(ByVal txtCantidad As Variant)

If IsNumeric(txtCantidad) Then
AgregarItem
Else
MsgBox "Número incorrecto"
End If

End Sub



Private Sub AgregarItem()
AbrirBase
Dim strSql As String
rsCarrito.Open ("Carrito"), DB, adOpenKeyset, adLockOptimistic, adCmdTable
rsCarrito.AddNew
rsCarrito!codarticulo = txtCodarticulo
rsCarrito!Descripcion = txtDescripcion
rsCarrito!Cantidad = Val(txtCantidad)

'If Option1.Value = True Then

rsCarrito!P_Unitario = Format(Val(txtPrecio), "#,###.#0")
rsCarrito!P_NETO = Format(Val(txtPrecio) * Val(txtCantidad), "#,###.#0")

'ElseIf Option2.Value = True Then
'rsCarrito!P_UNITARIO = Format(Val(txtPrecio) / Val(txtCantidad), "#,###.#0")
'rsCarrito!P_NETO = Format(Val(txtPrecio), "#,###.#0")

'End If
rsCarrito.Update
CalcularTotales
LimpioFlex
ActualizoFlex
CargarNuevoArticulo
CerrarBase
'Option1.Value = True
txtCantidad = ""

End Sub




Private Sub CalcularTotales()
Dim R As Double
Dim strSql As String
strSql = "SELECT * FROM Carrito"
rsTotales.Open ("Carrito"), DB, adOpenKeyset, adLockOptimistic, adCmdTable

While Not rsTotales.EOF
R = R + rsTotales!P_NETO
rsTotales.MoveNext
Wend
txtTotal = R
txtIva = txtTotal * 0.21
txtSubTotal = txtTotal - txtIva
End Sub


Private Sub CargarNuevoArticulo()
txtCodarticulo.Enabled = True
        txtMarca = ""
txtTalle = ""

lblCodInt = ""
txtCodarticulo = ""
txtDescripcion = ""
txtPrecio = ""
txtCantidad = ""
txtCodarticulo.SetFocus
End Sub


Private Sub ActualizoFlex()

    'me FIJO EN LA TABLA COUTAS SI EL SOCIO TIENE CUOTAS
    Dim RstCtas As New Recordset
    LineaSQL = "select * from Carrito"
    RstCtas.Open LineaSQL, DB, adOpenKeyset, adLockOptimistic, adCmdText
    
    CtasOpen = True
    
    'SI NO TIENE Cuotas, Es porque está dado de baja
    If RstCtas.BOF = True And RstCtas.EOF = True Then
        
    'ACA ES DONDE DEBO GENERAR LAS CUOTAS
    MsgBox "No hay articulos!", vbExclamation
    Exit Sub
    'GeneroCuotas
        
    'VeoCuotas
    End If
    
    'SI SI TIENE LAS MANDO AL FLEXGRID UNA POR UNA
    i = 0
    yaesta = 0
    Do While Not RstCtas.EOF
        
        i = i + 1
       
        ' Y TODO POR ESTA LINEA DE MIERDA
        'VENCIMIENTO / DESCRIPCION / IMPORTE DE CUOTA / FECHA DE PAGO
        linea = Format(RstCtas!codarticulo, "0###") _
                & Chr(9) & RstCtas!Descripcion _
                & Chr(9) & Format(RstCtas!Cantidad, "0###") _
                & Chr(9) & Format(RstCtas!P_Unitario, "#,###.#0") _
                & Chr(9) & Format(RstCtas!P_NETO, "#,###.#0")
        
        
    MSHFlexGrid1.AddItem linea, i
    RstCtas.MoveNext
    Loop
    
    
    

End Sub




Private Sub VaciarCarrito()
Dim strSql As String
AbrirBase
strSql = "SELECT * FROM Carrito"
rsCarrito.Open ("Carrito"), DB, adOpenKeyset, adLockOptimistic, adCmdTable
While (Not rsCarrito.EOF)
rsCarrito.Delete
rsCarrito.MoveNext
Wend
CalcularTotales
CerrarBase
LimpioFlex
'Refrescar
CargarNuevoArticulo
End Sub

Private Sub cmdAgregarItem_Click()
'AceptarCantidad txtCantidad.Text
End Sub

Private Sub cmdEliminarItem_Click()
EliminarItem

End Sub
Private Sub Command1_Click()
VaciarCarrito
End Sub


Private Sub LimpioFlex()
titulos = " COD.|DESCRIPCION |CANT.|P.UNITARIO|P.NETO"

 MSHFlexGrid1.Clear
    With MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .ColWidth(0) = 600
        .ColWidth(1) = 2500
        .ColWidth(2) = 600
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With
End Sub


Private Sub MSHFlexGrid1_DblClick()
MSHFlexGrid1.Col = 0
If IsNumeric(MSHFlexGrid1.Text) Then
If MsgBox("¿Elimina este Item?", vbExclamation + vbYesNoCancel) = vbYes Then
EliminarItem
End If
End If
End Sub

Private Sub EliminarItem()
Dim itemselecto As Integer
Dim strSql As String
AbrirBase
MSHFlexGrid1.Col = 0
itemselecto = Val(MSHFlexGrid1.Text)

strSql = "SELECT * FROM Carrito Where CodArticulo=" & Val(itemselecto)
rsCarrito.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not (rsCarrito.BOF And rsCarrito.EOF) Then
rsCarrito.MoveLast
rsCarrito.Delete
rsCarrito.Update
End If
CalcularTotales
'Refrescar
LimpioFlex
ActualizoFlex
CargarNuevoArticulo
CerrarBase
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  SECCION GUARDAR CAMBIOS E IMPRIMIR   @@@@@@@@@@@@@@@@@@@





Private Sub ImprimirFactura()

If txtTotal <> 0 And txtNumFact <> "" And txtCodCliente <> "" Then

GuardarTodo

End If


End Sub
Private Sub GuardarTodo()
'On Error GoTo Error_Guardar

AbrirBase
'DB.BeginTrans
GuardarDatosDeCliente
GuardarDatosDePresupuesto
IniciarImpresion

'Error_Guardar:
'    If Err.Number <> 0 Then
'    DB.RollbackTrans
'    MsgBox "Error: " & Err.Number & " - " & Err.Description
'    CerrarBase
'    Else
'    DB.CommitTrans
    CerrarBase
    PresupuestoOK
'    End If

End Sub



Private Sub GuardarDatosDeCliente()
Dim strSql As String
strSql = "SELECT * FROM Clientes " & _
                 "WHERE ID =" & txtCodCliente
        
rsClientes.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
        
    'SI EXISTE QUE LO ACTUALIZE
    If Not (rsClientes.BOF And rsClientes.EOF) Then
    rsClientes!id = txtCodCliente
    rsClientes!nombre = txtNombre
    rsClientes!domicilio = txtDomicilio
    rsClientes!telefono = txtTelefono
    'rsClientes!CategIva = ComboCategIva
    rsClientes!cuit = txtCuit
    rsClientes.Update
    Else  'SI NO EXISTE QUE LO AGREGUE
    rsClientes.AddNew
    rsClientes!id = txtCodCliente
    rsClientes!nombre = txtNombre
    rsClientes!domicilio = txtDomicilio
    rsClientes!telefono = txtTelefono
    'rsClientes!CategIva = ComboCategIva
    rsClientes!cuit = txtCuit
    rsClientes.Update
    End If

End Sub



Private Sub GuardarDatosDePresupuesto()
Dim n As Integer
Dim v As String



rsPresupuestos.Open ("Presupuestos"), DB, adOpenKeyset, adLockOptimistic, adCmdTable
      
            If Not (rsPresupuestos.BOF And rsPresupuestos.EOF) Then
            rsPresupuestos.MoveLast
            n = Val(rsPresupuestos!NumPres) + 1
            rsPresupuestos.AddNew
            rsPresupuestos!NumPres = n
            rsPresupuestos!Fecha = Date
            rsPresupuestos!Hora = Time
            rsPresupuestos!CodUsuario = Val(txtCodUsuario)
            rsPresupuestos!CodCliente = txtCodCliente
            rsPresupuestos!total = txtTotal
            rsPresupuestos.Update
            Else
            n = 1
            rsPresupuestos.AddNew
            rsPresupuestos!NumPres = n
            rsPresupuestos!Fecha = Date
            rsPresupuestos!Hora = Time
            rsPresupuestos!CodUsuario = Val(txtCodUsuario)
            rsPresupuestos!CodCliente = txtCodCliente
            rsPresupuestos!total = txtTotal
            rsPresupuestos.Update
            End If
            
            
End Sub






Private Sub PresupuestoOK()
MsgBox "Presupuesto Imprimiendose", vbInformation
txtNumFact = ""
txtNumFact.SetFocus
End Sub




'€€€€€€€€€€€€€€€€€€€€€€€€  I M P R E S I O N €€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Private Sub IniciarImpresion()
For i = 1 To ComboCopias.Text
ImprimirTitulos
ImprimirDatosCliente
ImprimirCondicionesDeVenta
ImprimirTituloCampos
ImprimirDetalle
ImprimirTotales
Printer.NewPage
Printer.EndDoc
Next i
End Sub

Private Sub ImprimirTitulos()





'###########ENCABEZADO DE PRESUPUESTO ########################



Dim strSql As String
Dim x As String
strSql = "Select * From Empresa Where IdEmpresa = 1"
rsEmpresa.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText


'Tipo de Papel A4
Printer.PaperSize = 9
'Impresion en blanco y negro
'Printer.ColorMode = 1
        
Printer.ScaleMode = vbCentimeters



If Check1.Value = 0 Then

'Printer.PaintPicture frmLogo.Picture1.Picture, 1, 1

End If

Printer.Font.Size = 14

Printer.CurrentX = 10
Printer.CurrentY = 1
Printer.Print "X"


Printer.CurrentX = 15
Printer.CurrentY = 1.3
'Printer.Print "Original Cliente / Copia Vendedor"
Printer.Print "PRESUPUESTO"

Printer.Font.Size = 9
Printer.CurrentX = 15
Printer.CurrentY = 2
Printer.Print "Presup. Nº: " & Format(rsPresupuestos!NumPres, "0001-########")


'BORDES HORIZONTALES
Printer.Line (1, 1)-(20, 1)
Printer.Line (1, 29)-(20, 29)
'BORDES VERTICALES
Printer.Line (1, 1)-(1, 29)
Printer.Line (20, 1)-(20, 29)


Printer.Font.Size = 9

Printer.CurrentX = 15
Printer.CurrentY = 2.5
Printer.Print "Página " & Printer.Page

Printer.CurrentX = 15
Printer.CurrentY = 3
Printer.Print "Fecha: " & rsPresupuestos!Fecha & " " & Format(rsPresupuestos!Hora, "hh:nn am/pm")

'SEGUNDO RENGLON

If Check1.Value = 0 Then


Printer.CurrentX = 2
Printer.CurrentY = 4
Printer.Print rsEmpresa!Direccion


Printer.CurrentX = 15
Printer.CurrentY = 4
Printer.Print "C.U.I.T. Nº: " & rsEmpresa!cuit


'TERCER RENGLON

Printer.CurrentX = 2
Printer.CurrentY = 4.4
Printer.Print "TEL: " & rsEmpresa!telefono & " " & "Sucursal: " & rsEmpresa!Sucursal

Printer.CurrentX = 15
Printer.CurrentY = 4.4
Printer.Print "Ing.Brutos: " & rsEmpresa!IngresosBrutos


' CUARTO RENGLON

Printer.CurrentX = 2
Printer.CurrentY = 4.8
Printer.Print rsEmpresa!categiva

Printer.CurrentX = 15
Printer.CurrentY = 4.8
Printer.Print "Inicio Act.: " & rsEmpresa!InicioActividades

End If

rsEmpresa.Close



Printer.Line (1, 5.5)-(20, 5.5)
'###########  FIN ENCABEZADO DE PRESUPUESTO ########################

End Sub

Private Sub ImprimirDatosCliente()

Printer.Font.Size = 10

Printer.CurrentX = 2
Printer.CurrentY = 6
Printer.Print "Señor/es: " & rsClientes!nombre & " (" & rsClientes!id & ")";

Printer.CurrentX = 15
Printer.Print "C.U.I.T. Nº: " & rsClientes!cuit

Printer.CurrentX = 2
Printer.CurrentY = Printer.CurrentY + 0.1
Printer.Print "Dirección: " & rsClientes!domicilio;

Printer.CurrentX = 15
Printer.Print "Teléfono: " & rsClientes!telefono

Printer.Line (1, Printer.CurrentY + 0.1)-(20, Printer.CurrentY + 0.1)

End Sub

Private Sub ImprimirCondicionesDeVenta()

Printer.Font.Size = 10

Printer.CurrentX = 2
Printer.CurrentY = Printer.CurrentY + 0.1
Printer.Print "Condiciones de Venta: Contado Efectivo";

Printer.CurrentX = 15
Printer.Print "Vendedor: " & Val(txtCodUsuario)

Printer.Line (1, Printer.CurrentY + 0.1)-(20, Printer.CurrentY + 0.1)

End Sub

Private Sub ImprimirTituloCampos()

Printer.CurrentY = Printer.CurrentY + 0.1

Printer.CurrentX = 2
Printer.Print "ART.";
Printer.CurrentX = 3.5
Printer.Print "DESCRIPCION";
Printer.CurrentX = 11
Printer.Print "CANTIDAD";
Printer.CurrentX = 14
Printer.Print "P.UNITARIO";
Printer.CurrentX = 17
Printer.Print "SUBTOTAL"

Printer.Line (1, Printer.CurrentY + 0.1)-(20, Printer.CurrentY + 0.1)

End Sub

Private Sub ImprimirDetalle()

x = "Select * From Carrito"
rsPrintItem.Open (x), DB, adOpenKeyset, adLockOptimistic, adCmdText

While (Not rsPrintItem.EOF)

Printer.CurrentY = Printer.CurrentY + 0.1

Printer.CurrentX = 2
Printer.Print Format(rsPrintItem!codarticulo, "0###");

Printer.CurrentX = 3.5
Printer.Print rsPrintItem!Descripcion;

Printer.CurrentX = 11.5
Printer.Print Format(rsPrintItem!Cantidad, "0###");

Dim CadenaU As String
Dim LongitudU As Long
CadenaU = Format(CadenaU, "#,###.#0")
LongitudU = Len(CadenaU)
Printer.CurrentX = 15.5 + (Val(LongitudU)) - (Printer.TextWidth(Format(rsPrintItem!P_Unitario, "#,###.#0")))
Printer.Print Format(rsPrintItem!P_Unitario, "#,###.#0");


Dim CadenaN As String
Dim LongitudN As Long
CadenaN = Format(CadenaN, "#,###.#0")
LongitudN = Len(CadenaN)
Printer.CurrentX = 18.5 + (Val(LongitudN)) - (Printer.TextWidth(Format(rsPrintItem!P_NETO, "#,###.#0")))
Printer.Print Format(rsPrintItem!P_NETO, "#,###.#0")


       If Printer.CurrentY > 28 Then
            Printer.NewPage
            ImprimirTitulos
            ImprimirDatosCliente
            ImprimirCondicionesDeVenta
            ImprimirTituloCampos
       End If
       
rsPrintItem.MoveNext
Wend

rsPrintItem.Close

End Sub
Private Sub ImprimirTotales()


Printer.Font.Size = 14

Printer.Line (1, 25)-(20, 25)

Printer.CurrentX = 19 - Printer.TextWidth("SUBTOTAL $: " & Format(rsPresupuestos!total, "#,###.#0"))
Printer.CurrentY = 25.5
Printer.Print "SUBTOTAL $: " & Format(rsPresupuestos!total, "#,###.#0")

Printer.Line (1, 26)-(20, 26)

Printer.CurrentX = 19 - Printer.TextWidth("TOTAL $: " & Format(rsPresupuestos!total, "#,###.#0"))
Printer.CurrentY = 26.5
Printer.Print "TOTAL $: " & Format(rsPresupuestos!total, "#,###.#0")

Printer.Line (1, 27)-(20, 27)

Printer.Font.Size = 10
Printer.CurrentX = 2
Printer.CurrentY = 27.5
Printer.Print "FIRMA:......................................................"

Printer.Line (1, 28)-(20, 28)
End Sub

'#########################  F I N   D E   I M P R E S I O N  ##################










' ####################### KeyPress  ##########################



Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 Then
DeshabilitarCliente

KeyAscii = 0
End If
End Sub
Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 Then
txtNombre.SetFocus
KeyAscii = 0

End If
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 Then
txtDomicilio.SetFocus
KeyAscii = 0

End If
End Sub

Private Sub ComboCategIva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 Then
txtTelefono.SetFocus
KeyAscii = 0
End If
End Sub

Private Sub txtCuit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 Then
ComboCategIva.SetFocus
KeyAscii = 0
End If
End Sub

Private Sub ComboTipoFact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 Then
'DeshabilitarUsuario
KeyAscii = 0
End If
End Sub

Private Sub ComboCondVenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 Then
ComboTipoFact.SetFocus
KeyAscii = 0
End If
End Sub




Private Sub txtTotal_Change()
txtTotal = Format(txtTotal, "#,###.#0")
End Sub

Private Sub txtIva_Change()
txtIva = Format(txtIva, "#,###.#0")
End Sub


Private Sub txtSubtotal_Change()
txtSubTotal = Format(txtSubTotal, "#,###.#0")
End Sub




