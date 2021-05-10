VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFacturacion 
   BackColor       =   &H00000000&
   Caption         =   "FACTURAS"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   11880
   Icon            =   "frmFacturacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFF00&
      Caption         =   "RECARGO / DESCUENTO %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5220
      TabIndex        =   60
      Top             =   6840
      Width           =   7275
      Begin VB.CommandButton cmdAplicarRecargoSeleccion 
         Caption         =   "SELECCION"
         Height          =   315
         Left            =   3780
         TabIndex        =   64
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtPorcentaje 
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
         Left            =   2460
         MaxLength       =   12
         TabIndex        =   62
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAplicarRecargo 
         Caption         =   "GENERAL"
         Height          =   315
         Left            =   5160
         TabIndex        =   61
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         Caption         =   "RECARGO / DESCUENTO % ="
         Height          =   255
         Left            =   60
         TabIndex        =   63
         Top             =   300
         Width           =   2355
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Opciones de Impresión:"
      Height          =   795
      Left            =   0
      TabIndex        =   56
      Top             =   6840
      Width           =   5055
      Begin VB.CheckBox chkFiscal 
         BackColor       =   &H00FFFF00&
         Caption         =   "Impresora Fiscal"
         Height          =   330
         Left            =   1680
         TabIndex        =   72
         Top             =   315
         Width           =   1500
      End
      Begin VB.ComboBox ComboCopias 
         Height          =   315
         ItemData        =   "frmFacturacion.frx":0442
         Left            =   4170
         List            =   "frmFacturacion.frx":0455
         TabIndex        =   58
         Text            =   "1"
         Top             =   300
         Width           =   525
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Imprimir Remito"
         Height          =   315
         Left            =   120
         TabIndex        =   57
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFF00&
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
         Left            =   3510
         TabIndex        =   59
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
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
      Height          =   6075
      Left            =   5220
      TabIndex        =   36
      Top             =   780
      Width           =   7275
      Begin VB.Frame FrameIVA 
         BackColor       =   &H00FFFF00&
         Caption         =   "IVA:"
         Height          =   780
         Left            =   135
         TabIndex        =   67
         Top             =   4950
         Width           =   4245
         Begin VB.TextBox txtIva 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   69
            Text            =   "0"
            Top             =   315
            Width           =   1215
         End
         Begin VB.TextBox txtSubTotal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   68
            Text            =   "0"
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFF00&
            Caption         =   "IVA"
            Height          =   255
            Left            =   2565
            TabIndex        =   71
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFF00&
            Caption         =   "Subtotal:"
            Height          =   240
            Left            =   135
            TabIndex        =   70
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.TextBox txtMarca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   66
         Top             =   855
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtTalle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4050
         MaxLength       =   50
         TabIndex        =   65
         Top             =   855
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ver Utilidad"
         Height          =   255
         Left            =   4860
         TabIndex        =   53
         Top             =   5580
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Limpiar"
         Height          =   255
         Left            =   4800
         TabIndex        =   50
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregarItem 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   2640
         TabIndex        =   49
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminarItem 
         Caption         =   "Eliminar"
         Height          =   255
         Left            =   3720
         TabIndex        =   41
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarArticulo 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   1560
         TabIndex        =   40
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarArticulo 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtCodArticulo 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   10
         Top             =   360
         Width           =   1095
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
         Left            =   4860
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0"
         Top             =   5220
         Width           =   1515
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   4440
         MaxLength       =   5
         TabIndex        =   11
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtPrecio 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Bindings        =   "frmFacturacion.frx":0468
         Height          =   3375
         Left            =   120
         TabIndex        =   51
         Top             =   1560
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   5953
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   8388608
         ForeColorFixed  =   16777215
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Begin VB.Label lblCodInt 
         BackColor       =   &H00FFFF00&
         Caption         =   "*"
         Height          =   195
         Left            =   2400
         TabIndex        =   52
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   3600
         TabIndex        =   46
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         Caption         =   "Precio:"
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF00&
         Caption         =   "Descripción:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF00&
         Caption         =   "CodArt:"
         Height          =   255
         Left            =   480
         TabIndex        =   43
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         Caption         =   "Total:"
         Height          =   255
         Left            =   4440
         TabIndex        =   42
         Top             =   5280
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Detalle de Factura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   0
      TabIndex        =   28
      Top             =   780
      Width           =   5115
      Begin VB.ComboBox ComboTipoFact 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   675
         ItemData        =   "frmFacturacion.frx":047E
         Left            =   480
         List            =   "frmFacturacion.frx":048B
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "C"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtHora 
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
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
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
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtNumFact 
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
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   30
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Nº FACT:"
         Height          =   255
         Left            =   1560
         TabIndex        =   29
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   22
      Top             =   2040
      Width           =   5115
      Begin VB.CommandButton cmdAceptarCliente 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   600
         TabIndex        =   48
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardarCliente 
         Caption         =   "Guardar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelarCliente 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3480
         TabIndex        =   33
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtCodCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtDomicilio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtTelefono 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox ComboCategIva 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmFacturacion.frx":0498
         Left            =   1320
         List            =   "frmFacturacion.frx":04AE
         TabIndex        =   5
         Text            =   "IVA Consumidor Final"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         Caption         =   "Categ. IVA:"
         Height          =   255
         Left            =   480
         TabIndex        =   47
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Señor/es:"
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         Caption         =   "ID:"
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cuit:"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         Caption         =   "Tel:"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Condiciones de Venta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   20
      Top             =   5520
      Width           =   5115
      Begin VB.ComboBox ComboCondVenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         ItemData        =   "frmFacturacion.frx":0532
         Left            =   2040
         List            =   "frmFacturacion.frx":053F
         TabIndex        =   9
         Text            =   "Contado/Efectivo"
         Top             =   240
         Width           =   2715
      End
      Begin MSComCtl2.DTPicker dtpVto 
         Height          =   375
         Left            =   2040
         TabIndex        =   73
         Top             =   840
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   16384
         CalendarTitleForeColor=   16384
         Format          =   49217537
         CurrentDate     =   38402
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         Caption         =   "Vencimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   55
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         Caption         =   "CondVenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   21
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFF00&
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
      Height          =   1035
      Left            =   0
      TabIndex        =   19
      Top             =   4440
      Width           =   5115
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cobra Comisión"
         Height          =   315
         Left            =   3360
         TabIndex        =   54
         Top             =   660
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptarUsuario 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   960
         TabIndex        =   35
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelarUsuario 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtUsuario 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   2025
      End
      Begin VB.TextBox txtCodUsuario 
         Height          =   285
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cod. Vendedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   7785
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "21/09/10"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "21:39"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1111
      ButtonWidth     =   2355
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nueva Factura"
            Key             =   "ToolNuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Factura"
            Key             =   "ToolPrint"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar Factura"
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
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar Factura"
            Key             =   "ToolBuscarFactura"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10860
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturacion.frx":057B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturacion.frx":068F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturacion.frx":07A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturacion.frx":08B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturacion.frx":09CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturacion.frx":0ADF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturacion.frx":115B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   18
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
         Caption         =   "&Nueva Factura"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuCancelarFactura 
         Caption         =   "&Cancelar Factura"
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
      Begin VB.Menu mnuBuscarDistribuidor 
         Caption         =   "&Buscar Vendedor"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuBuscarArticulo 
         Caption         =   "&Buscar Artículo"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuBuscarFactura 
         Caption         =   "&Buscar Orden de Pedido"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Funciones de Impresora Fiscal"
      Begin VB.Menu mnuEstadoPFiscal 
         Caption         =   "Estado de la Impresora"
      End
      Begin VB.Menu mnuProxNum 
         Caption         =   "Ver Proximo numero de factura"
      End
      Begin VB.Menu mnuReporteX 
         Caption         =   "Imprimir Cierre X cambio de cajero"
      End
      Begin VB.Menu ImprimirCierreZ 
         Caption         =   "Imprimir Cierre de Jornada Fiscal Z"
      End
   End
End
Attribute VB_Name = "frmFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Funciones() As String
Dim AyudaFunciones() As String
Dim Parametros() As String

Private Sub Form_Load()

Check1.Value = IIf(ReadINI("PRINTREM", "VALUE") = "", 1, ReadINI("PRINTREM", "VALUE"))
Me.chkFiscal.Value = IIf(ReadINI("PRINTFISCAL", "VALUE") = "", 1, ReadINI("PRINTFISCAL", "VALUE"))
ComboCopias.Text = IIf(ReadINI("NUMPAGE", "VALUE") = "", 1, ReadINI("NUMPAGE", "VALUE"))
WriteIni "PRINTREM", "VALUE", Check1.Value
WriteIni "PRINTFISCAL", "VALUE", Me.chkFiscal.Value
WriteIni "NUMPAGE", "VALUE", ComboCopias.Text

Dim i As Integer
Dim nombre As String
Dim Ayuda As String
ReDim Funciones(1)
ReDim AyudaFunciones(1)


If ComboTipoFact.Text = "C" Then
FrameIVA.Visible = False
End If

End Sub


'#########################  SECCION CHECKS  ###########################
Private Sub Check1_Click()
WriteIni "PRINTREM", "VALUE", Check1.Value
End Sub

Private Sub chkFiscal_Click()
WriteIni "PRINTFISCAL", "VALUE", Me.chkFiscal.Value
End Sub
Private Sub ComboCopias_Click()
WriteIni "NUMPAGE", "VALUE", ComboCopias.Text
End Sub

Private Sub Command2_Click()
frmVerUtilidad.Show
End Sub




Private Sub mnuBuscarFactura_Click()
'frmBuscarFactura.Show
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


'&&&&&&&&&&&&&&&&&& MENU FISCAl &&&&&&&&&&&&&&&&&&&&&&&

Private Sub mnuEstadoPFiscal_Click()

End Sub

Private Sub mnuSalir_Click()
Unload Me
End Sub
Private Sub mnuBuscarCliente_Click()
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
NuevoCliente
Case "ToolBuscarCliente"
frmBuscarClienteF.Show
Case "ToolBuscarArticulo"
frmBuscarArticulo.Show
Case "ToolBuscarFactura"
frmGanancia.Show
End Select
End Sub

'#########################  SECCION ENCABEZADO FACTURAS   ###########################


Private Sub ComboTipoFact_Click()
NuevaFactura
End Sub
Private Sub ComboTipoFact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
NuevaFactura
ElseIf KeyAscii = 8 Then
CancelarFactura
KeyAscii = 0
End If
End Sub

Private Sub NuevaFactura()
ClearControles
CargarNuevaFactura
HabilitarFactura
End Sub
Private Sub CargarNuevaFactura()
txtNumFact = NuevoNumeroFactura(ComboTipoFact.Text)
txtFecha = CDate(Date)
txtHora = Time
dtpVto = CDate(Date)
ComboCondVenta.Text = "Contado/Efectivo"
txtCodCliente.SetFocus
End Sub

Private Sub HabilitarFactura()
If ComboTipoFact.Text = "A" Or ComboTipoFact.Text = "B" Then
FrameIVA.Visible = True
Else
FrameIVA.Visible = False
End If
End Sub

Private Sub CancelarFactura()
ClearControles
End Sub

Private Sub ClearControles()
VaciarCarrito
CalcularTotales ComboTipoFact, txtSubTotal, txtIva, txtTotal
RefreshGrid MSHFlexGrid1, "Carrito"
DeshabilitarArticulo
DeshabilitarCliente
DeshabilitarUsuario
deshabilitarFactura
End Sub

Private Sub deshabilitarFactura()
txtFecha = ""
txtHora = ""
txtNumFact = ""
ComboTipoFact.SetFocus
End Sub

Private Sub txtCodCliente_LostFocus()
If Val(txtCodCliente) = "1" Then
txtCodUsuario.SetFocus
End If
End Sub

'#########################  SECCION USUARIO   ###########################


Private Sub txtCodUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(txtCodUsuario) Then
    KeyAscii = 0
    ObtenerUsuario txtCodUsuario, txtUsuario, ComboCondVenta
    ElseIf KeyAscii = 8 And txtCodUsuario = "" Then
    txtCuit.SetFocus
    KeyAscii = 0
    End If
End Sub


Private Sub DeshabilitarUsuario()
On Error Resume Next
txtCodUsuario = ""
txtUsuario = ""
txtCodUsuario.Enabled = True
txtUsuario.Enabled = False
cmdCancelarUsuario.Enabled = False
cmdAceptarUsuario.Enabled = True
txtCodUsuario.SetFocus
End Sub


'#########################  SECCION CLIENTE  ###########################

Private Sub NuevoCliente()
Dim strSql As String
AbrirBase
Dim rsClientes As New Recordset
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

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(txtCodCliente) Then
    KeyAscii = 0
    ObtenerCliente txtCodCliente, txtNombre, txtDomicilio, txtTelefono, ComboCategIva, txtCuit
    ElseIf KeyAscii = 8 And txtCodCliente = "" Then
    ComboTipoFact.SetFocus
    KeyAscii = 0
    End If
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
'pero habilito el campo clave
txtCodCliente.Enabled = True
txtCodCliente.SetFocus
End Sub


Private Sub cmdAceptarCliente_Click()
ObtenerCliente txtCodCliente, txtNombre, txtDomicilio, txtTelefono, ComboCategIva, txtCuit
End Sub

Private Sub cmdGuardarCliente_Click()
GuardarCliente txtCodCliente, txtNombre, txtDomicilio, txtTelefono, ComboCategIva, txtCuit
End Sub
Private Sub cmdCancelarCliente_Click()
DeshabilitarCliente
End Sub

'#########################  SECCION ARTICULO  ###########################

Private Sub txtCodArticulo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And IsNumeric(txtCodarticulo) Then

If InStr(1, txtCodarticulo, "%", vbTextCompare) > 0 Then
Dim x As String
x = txtCodarticulo
pp = InStr(1, x, "%", vbTextCompare)
txtCodarticulo = Replace(Mid(x, 1, pp), "%", "")
txtMarca = Mid(Replace(Mid(x, pp, Len(x)), "%", ""), 1, 2)
txtTalle = Mid(Replace(Mid(x, pp, Len(x)), "%", ""), 3, Len(x))
End If

KeyAscii = 0
GetProduct txtCodarticulo, txtDescripcion, txtPrecio, txtCantidad
ElseIf KeyAscii = 8 And txtCodarticulo = "" Then
txtCuit.SetFocus
KeyAscii = 0
End If
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
txtCodarticulo.SetFocus
End Sub

Private Sub cmdAceptarArticulo_Click()
GetProduct txtCodarticulo, txtDescripcion, txtPrecio, txtCantidad
End Sub

Private Sub cmdCancelarArticulo_Click()
DeshabilitarArticulo
End Sub


'#########################  SECCION Cantidad  ###########################

Private Sub txtCantidad_GotFocus()
txtCantidad.SelLength = Len(txtCantidad)
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(txtCantidad.Text) Then
    AgregarProducto ComboTipoFact, txtCodarticulo, txtDescripcion, txtPrecio, txtCantidad
    CalcularTotales ComboTipoFact, txtSubTotal, txtIva, txtTotal
    RefreshGrid MSHFlexGrid1, "Carrito"

    ElseIf KeyAscii = 8 And txtCantidad = "" Then
    DeshabilitarArticulo
    KeyAscii = 0
    End If
End Sub



'''''###################RECARGO DESCUENTO ##############################

' ############ INDIVIDUAL

Private Sub cmdAplicarRecargoSeleccion_Click()
AplicarRecargoIndividual MSHFlexGrid1, txtPorcentaje
CalcularTotales ComboTipoFact, txtSubTotal, txtIva, txtTotal
RefreshGrid MSHFlexGrid1, "Carrito"
End Sub

' ############ GENERAL

Private Sub cmdAplicarRecargo_Click()
AplicarRecargoGeneral txtPorcentaje
CalcularTotales ComboTipoFact, txtSubTotal, txtIva, txtTotal
RefreshGrid MSHFlexGrid1, "Carrito"
End Sub


'''''################### FIN RECARGO DESCUENTO ##############################



Private Sub MSHFlexGrid1_DblClick()
If MsgBox("¿Elimina este Item?", vbExclamation + vbYesNoCancel) = vbYes Then
EliminarItem MSHFlexGrid1
CalcularTotales ComboTipoFact, txtSubTotal, txtIva, txtTotal
RefreshGrid MSHFlexGrid1, "Carrito"
End If
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  SECCION GUARDAR CAMBIOS E IMPRIMIR   @@@@@@@@@@@@@@@@@@@

Private Sub ImprimirFactura()
If txtTotal <> 0 And txtNumFact <> "" And txtCodCliente <> "" Then
GuardarTodo
End If
End Sub

Private Sub GuardarTodo()
'On Error GoTo Error_Guardar
GuardarCliente txtCodCliente, txtNombre, txtDomicilio, txtTelefono, ComboCategIva, txtCuit

AbrirBase
'DB.BeginTrans
nfact = GuardarDatosDeFactura(ComboTipoFact, txtCodUsuario, txtCodCliente, ComboCondVenta, dtpVto, txtSubTotal, txtIva, txtTotal)
nrem = GuardarDatosDeRemito(ComboTipoFact, txtCodUsuario, txtCodCliente, ComboCondVenta, dtpVto, txtSubTotal, txtIva, txtTotal)
'GuardarUtilidad nfact, ComboTipoFact, ComboCondVenta
ActualizarCuentasCorrientes ComboCondVenta, txtCodCliente, nfact, txtTotal
ActualizarStock

CerrarBase

If chkFiscal.Value = 0 Then
IniciarImpresion nfact, nrem
Else
IniciarImpresionFiscal ComboTipoFact, nfact, ComboCopias
End If

FacturaOK
End Sub
Private Sub FacturaOK()
MsgBox "Venta Exitosamente Registrada!", vbInformation
NuevaFactura
End Sub


'€€€€€€€€€€€€€€€€€€€€€€€€  I M P R E S I O N €€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Private Sub IniciarImpresion(nfact, nrem)
PrintFactura ComboTipoFact, nfact, ComboCopias
If Check1.Value = 1 Then
PrintRemito ComboTipoFact, nrem, ComboCopias
End If
End Sub
Private Sub IniciarImpresionFiscal(ComboTipoFact, nfact, ComboCopias)
Screen.MousePointer = 11
PrintFacturaFiscal CStr(ComboTipoFact), CStr(nfact), CStr(ComboCopias)
Screen.MousePointer = 1
End Sub





'@@@@@@@@@@@@@@@@@@@@@@ IMPRESORA FISCAL ##########################################################
'@@@@@@@@@@@@@@@@@@@@@@ IMPRESORA FISCAL ##########################################################
'@@@@@@@@@@@@@@@@@@@@@@ IMPRESORA FISCAL ##########################################################

Public Sub PrintFacturaFiscal(ComboTipoFact As String, nfact As String, ComboCopias As String)

On Error GoTo hell

Dim respuesta As Boolean

AbrirBase

PrinterFiscal1.SetPreference "P", "T", "P", "N"


strSql = "Select * From Facturas" & ComboTipoFact & " WHERE NumFact=" & Val(nfact)
Dim rsFacturas As New Recordset
rsFacturas.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsFacturas.EOF Then

    sql2 = "Select * From Clientes Where ID=" & Val(rsFacturas!CodCliente)
    Dim rsClientes As New Recordset
    rsClientes.Open sql2, DB, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsClientes.EOF Then
    
        
    Dim emp As String
    emp = "SELECT * FROM EMPRESA WHERE IdEmpresa = 1"
    Dim rsEmpresa As New Recordset
    rsEmpresa.Open (emp), DB, adOpenKeyset, adLockOptimistic, adCmdText
            
    numfact = rsFacturas!numfact
    tipofact = rsFacturas!tipofact
    resp_vendedor = "I" 'ObtenerPrefijoCategIVA(rsEmpresa!categiva)
    resp_comprador = "F" 'ObtenerPrefijoCategIVA(rsClientes!categiva)

    If tipofact = "A" Then
    tipodoc = "CUIT"
    resp_comprador = "I"
    Else
    resp_comprador = "F"
    tipodoc = ""
    End If
    
    NomCliente = rsClientes!nombre
    cuitCliente = QuitarGuionesCuit(" " & rsClientes!cuit)
    DomicilioCliente = "" & rsClientes!domicilio
      
      
    'A B R I R  F A C T U R A
respuesta = PrinterFiscal1.OpenInvoice("F", "C", CStr(tipofact), CStr(ComboCopias), _
"P", "12", CStr(resp_vendedor), CStr(resp_comprador), CStr(NomCliente), "", _
CStr(tipodoc), CStr(cuitCliente), "N", CStr(DomicilioCliente), "", "", "01", "01", "C")




        x = "Select * From Carrito"
        Dim rsPrintItem As New Recordset
        rsPrintItem.Open (x), DB, adOpenKeyset, adLockOptimistic, adCmdText
        While (Not rsPrintItem.EOF)
        descItem = rsPrintItem!Descripcion
        cantItem = rsPrintItem!Cantidad & "000"
        precioItem = Replace(Format(rsPrintItem!P_Unitario, "standard"), ",", "")
        precioItem = Replace(precioItem, ".", "")
        ivaitem = ObtenerImpuestoProducto(Val(rsPrintItem!codarticulo))
        ivaitem = Format((ivaitem - 1) * 100, "standard")
        ivaitem = Replace(ivaitem, ",", "")
        
        
        'E N V I A R  L O S  I T E M S  D E  A  U N O
        If respuesta Then respuesta = PrinterFiscal1.SendInvoiceItem( _
          CStr(descItem), _
        CStr(cantItem), _
        CStr(precioItem), _
        CStr(ivaitem), _
        "M", "0", "0", "", "", "", "0", "0")
        
        
        rsPrintItem.MoveNext
        Wend
        rsPrintItem.Close


        MontoDescuento = 0
        MontoRecargo = 0
        MontoPago = rsFacturas!total
        MontoRevPago = 0



'If MontoDescuento > 0 Then
'If respuesta Then respuesta = PrinterFiscal1.SendInvoicePayment("Descuento", MontoDescuento, "D")
'End If

'If MontoRecargo > 0 Then
'If respuesta Then respuesta = PrinterFiscal1.SendInvoicePayment("Recargo", MontoRecargo, "R")
'End If

' ENVIAR MONTO DEL PAGO
'If MontoPago > 0 Then
'If respuesta Then respuesta = PrinterFiscal1.SendInvoicePayment("Pago", CStr(MontoPago), "T")
'End If

'If MontoRevPago > 0 Then
'If respuesta Then respuesta = PrinterFiscal1.SendInvoicePayment("Reversión de Pago", MontoRevPago, "t")
'End If

'C E R R A R  F A C T U R A #############################################
If respuesta Then respuesta = PrinterFiscal1.CloseInvoice("F", CStr(tipofact), "")


End If
rsClientes.Close


End If
rsFacturas.Close

rsEmpresa.Close


If respuesta = False Then
devolverstock tipofact, numfact
borrarfactura tipofact, numfact
End If

CerrarBase

'MsgBox respuesta & Chr$(13) & Me.PrinterFiscal1.FiscalStatus & Chr$(13) & Me.PrinterFiscal1.PrinterStatus
Exit Sub
hell:
devolverstock tipofact, numfact
borrarfactura tipofact, numfact
CerrarBase
MsgBox "hubo un error en el proceso de impresión, la operación no se registró", vbCritical
Exit Sub
End Sub

Function AddParam(NewParam As String, Optional Reset = False) As Integer
'Agrega un parametro a la lista
Dim i As Integer
If Reset Then
    i = 0
    ReDim Parametros(i)
    Parametros(i) = NewParam
Else
    i = UBound(Parametros) + 1
    
    ReDim Preserve Parametros(i)
    Parametros(i) = NewParam
End If
End Function



























' ####################### KeyPress  ##########################

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtNombre = "" Then
DeshabilitarCliente
KeyAscii = 0
End If
End Sub
Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtDomicilio = "" Then
txtNombre.SetFocus
KeyAscii = 0
End If
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtTelefono = "" Then
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
ElseIf KeyAscii = 8 And txtCuit = "" Then
ComboCategIva.SetFocus
KeyAscii = 0
End If
End Sub

Private Sub ComboCondVenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 Then
'ComboTipoFact.SetFocus
txtCodUsuario.SetFocus
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
Private Sub dtpVto_Change()
txtCodarticulo.SetFocus
End Sub


