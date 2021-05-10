VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmArticulos 
   BackColor       =   &H00000000&
   Caption         =   "ABM Articulos"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11865
   ForeColor       =   &H00800000&
   Icon            =   "frmArticulos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir Codigo de Barras"
      Height          =   375
      Left            =   9810
      TabIndex        =   41
      Top             =   765
      Width           =   2040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Leer desde Barcode Scanner"
      Height          =   375
      Left            =   7515
      TabIndex        =   40
      Top             =   765
      Width           =   2220
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      Left            =   7560
      ScaleHeight     =   1365
      ScaleWidth      =   3780
      TabIndex        =   38
      Top             =   1170
      Width           =   3810
   End
   Begin VB.Frame FrameStock 
      BackColor       =   &H00FFFF00&
      Caption         =   "Modificar Existencias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   7560
      TabIndex        =   31
      Top             =   2745
      Visible         =   0   'False
      Width           =   3720
      Begin VB.TextBox txtSumarExistencias 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1620
         MaxLength       =   13
         TabIndex        =   34
         Top             =   270
         Width           =   1875
      End
      Begin VB.TextBox txtRestarExistencias 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1620
         MaxLength       =   13
         TabIndex        =   33
         Top             =   585
         Width           =   1875
      End
      Begin VB.CommandButton cmdAplicarStock 
         Caption         =   "Aplicar"
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
         Left            =   1620
         TabIndex        =   32
         Top             =   900
         Width           =   1890
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         Caption         =   "Sumar a Stock:"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         Caption         =   "Restar de Stock:"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   585
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FF00&
      Caption         =   "Buscar Artículo:"
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
      Height          =   795
      Left            =   60
      TabIndex        =   20
      Top             =   3300
      Width           =   6735
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
         TabIndex        =   23
         Top             =   240
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
         ItemData        =   "frmArticulos.frx":57E2
         Left            =   3420
         List            =   "frmArticulos.frx":57EF
         TabIndex        =   22
         Text            =   "Descripcion"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   5640
         TabIndex        =   21
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Buscar por:"
         Height          =   375
         Index           =   12
         Left            =   2580
         TabIndex        =   24
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Datos del Artículo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   60
      TabIndex        =   10
      Top             =   660
      Width           =   7395
      Begin VB.TextBox txtTalle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4995
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1125
         Width           =   1515
      End
      Begin VB.CommandButton cmdModificarStock 
         Caption         =   "Modificar Existencias"
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
         Left            =   5160
         TabIndex        =   37
         Top             =   2160
         Width           =   2115
      End
      Begin VB.TextBox txtMarca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1125
         Width           =   1515
      End
      Begin VB.TextBox txtStockMinimo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4995
         MaxLength       =   13
         TabIndex        =   7
         Top             =   1725
         Width           =   1515
      End
      Begin VB.TextBox txtExistencias 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1695
         MaxLength       =   13
         TabIndex        =   6
         Top             =   1725
         Width           =   1515
      End
      Begin VB.TextBox txtPrecioProv 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4995
         MaxLength       =   13
         TabIndex        =   5
         Top             =   1425
         Width           =   1515
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   2700
         TabIndex        =   14
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdContinuar 
         Caption         =   "C&ontinuar"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1710
         TabIndex        =   0
         Top             =   360
         Width           =   1530
      End
      Begin VB.TextBox txtPrecio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1695
         MaxLength       =   13
         TabIndex        =   4
         Top             =   1425
         Width           =   1515
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1695
         MaxLength       =   50
         TabIndex        =   1
         Top             =   780
         Width           =   4800
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000FFFF&
         Caption         =   "Cont:"
         Height          =   255
         Left            =   4455
         TabIndex        =   39
         Top             =   1170
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "Marca:"
         Height          =   255
         Left            =   450
         TabIndex        =   30
         Top             =   1185
         Width           =   1155
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FFFF&
         Caption         =   "Stock Mínimo  :"
         Height          =   255
         Left            =   3795
         TabIndex        =   29
         Top             =   1785
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "Existencias:"
         Height          =   255
         Left            =   540
         TabIndex        =   28
         Top             =   1785
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FFFF&
         Caption         =   "Precio Compra:"
         Height          =   255
         Left            =   3795
         TabIndex        =   27
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "Precio Unitario:"
         Height          =   255
         Left            =   540
         TabIndex        =   19
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "Descripción:"
         Height          =   255
         Left            =   495
         TabIndex        =   18
         Top             =   885
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "Codigo:"
         Height          =   300
         Left            =   540
         TabIndex        =   17
         Top             =   495
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   3375
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
            Picture         =   "frmArticulos.frx":580F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArticulos.frx":5C63
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArticulos.frx":60B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArticulos.frx":650B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArticulos.frx":695F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArticulos.frx":6DB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArticulos.frx":7207
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArticulos.frx":765B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArticulos.frx":7AAF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1058
      ButtonWidth     =   2593
      ButtonHeight    =   953
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo Articulo"
            Key             =   "ToolNuevo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar Articulo"
            Key             =   "ToolGuardar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Listado"
            Key             =   "ToolPrint"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Inventario"
            Key             =   "ToolPrintInventario"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ImprimirFicha"
            Key             =   "ToolPrintFicha"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar Articulo"
            Key             =   "ToolEliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar Edición"
            Key             =   "ToolCancelar"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   25
      Top             =   7605
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "08/05/2021"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "03:28 p.m."
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmArticulos.frx":7DCB
      Height          =   3315
      Left            =   120
      TabIndex        =   26
      Top             =   4200
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   5847
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   0
      FixedCols       =   0
      BackColorFixed  =   128
      ForeColorFixed  =   16777215
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
End
Attribute VB_Name = "frmArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############MODIFICAR EL STOCK ####################3

Private Sub cmdModificarStock_Click()
If txtID <> "" And IsNumeric(txtID) Then

If InputBox("Ingrese Codigo de seguridad para poder Modificar Stock", "Clave requerida") = "admin" Then

FrameStock.Visible = True
txtSumarExistencias.SetFocus

Else
MsgBox "La clave es incorrecta", vbCritical
End If

End If
End Sub


Private Sub cmdAplicarStock_Click()
If txtID <> "" And IsNumeric(txtID) Then
txtExistencias = Val(txtExistencias) + Val(txtSumarExistencias) - Val(txtRestarExistencias)
AgregarExistencias
txtSumarExistencias = ""
txtRestarExistencias = ""
FrameStock.Visible = False
MsgBox "Existencias Actualizadas", vbInformation
End If

End Sub

Private Sub AgregarExistencias()
AbrirBase
strSql = "SELECT * FROM Articulos WHERE ID =" & Val(txtID)
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not (rs.BOF And rs.EOF) Then
        rs!Existencias = Val(txtExistencias)
        rs.Update
        End If
CerrarBase
RefreshGrid MSHFlexGrid1, "Articulos"
End Sub


Private Sub Command1_Click()
AceptarBuscar txtBuscar.Text
End Sub

Private Sub Command2_Click()
DeshabilitarEdicion

End Sub

Private Sub Command3_Click()
frmPrintBarcode.Text1 = Me.txtID
frmPrintBarcode.txtMarca = Me.txtMarca
frmPrintBarcode.txtTalle = Me.txtTalle
frmPrintBarcode.Show 1
End Sub

Private Sub Form_Load()
'Me.Left = 0
'Me.Top = 0
'Me.Height = 5350
'Me.Width = 6000


RefreshGrid MSHFlexGrid1, "Articulos"

FechaHoy
End Sub
Private Sub FechaHoy()
dtp1 = Date

End Sub

Private Sub Form_Resize()
Me.MSHFlexGrid1.Width = Me.Width - 300
Me.MSHFlexGrid1.Height = Me.Height - Me.Height / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
'Private Sub cmdBuscar_Click()
'frmBuscarInsumo.Show
'End Sub
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
txtArticulo.SetFocus
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
frmBuscarInsumo.Show
End Sub
Private Sub mnuBuscar_Click()
frmBuscarInsumo.Show
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
ImprimirListado
Case "ToolPrintInventario"
ImprimirInventario
Case "ToolPrintFicha"
ImprimirFicha
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
Dim cont As Long
DeshabilitarEdicion
AbrirBase

strSql = "SELECT * FROM Articulos"
        
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
txtExistencias.Enabled = True
'TbLoc = "Río Cuarto"
'TbCodPos = "5800"
'TbNac = "Arg."
'dtp1 = Date
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
RefreshGrid MSHFlexGrid1, "Articulos"

End Sub

Private Sub EliminarCaso()
Dim strSql As String
strSql = "SELECT * FROM Articulos WHERE ID =" & Val(txtID)



rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not (rs.BOF And rs.EOF) Then
        rs.Delete
        rs.Update
        End If

End Sub

' ################ ACEPTAR #############################
Private Sub txtID_Change()
End Sub
Private Sub TxtID_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And txtID <> "" Then
    
If InStr(1, txtID, "%", vbTextCompare) > 0 Then
Dim x As String
x = txtID
pp = InStr(1, x, "%", vbTextCompare)
txtID = Replace(Mid(x, 1, pp), "%", "")
txtMarca = Mid(Replace(Mid(x, pp, Len(x)), "%", ""), 1, 2)
txtTalle = Mid(Replace(Mid(x, pp, Len(x)), "%", ""), 3, Len(x))
End If
    
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
    HabilitarEdicion
 '   RellenoCombo
'    End If

End Sub


Private Sub AceptarRegistro()
Dim strSql As String
strSql = "SELECT * FROM Articulos WHERE ID =" & Val(txtID)

rs.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
           
  ' Si existe
If Not (rs.BOF And rs.EOF) Then
TraigoDatos
'GeneroBarCode Val(txtID)
'BotoneraExploracion
Else
'HabilitarNuevo
HabilitarEdicion
PreparoTemplate
End If

End Sub

Private Sub Modebarcode(cadena As String)
Picture1.Cls
Picture1.ScaleMode = 3
Picture1.Height = Picture1.Height * (2.4 * 40 / Picture1.ScaleHeight)
Picture1.FontSize = 10
Call DrawBarcode(cadena, Picture1, txtMarca & " - " & txtTalle)
Clipboard.Clear
Clipboard.SetData Picture1.Image, 2
End Sub
Private Sub GeneroBarCode(cadena As String)
'ModeFree3of9 cadena
Modebarcode cadena
End Sub

Private Sub ModeFree3of9(cadena As String)

cadena = "*" & cadena & "*"

Text1 = ""
Text1.Font.Name = "Free 3 of 9"
Text1.Font.Size = 72
Text1 = cadena
Text2 = cadena
Clipboard.Clear
Clipboard.SetText cadena


End Sub

Private Sub TraigoDatos()
 On Error Resume Next
    'Traigo textos
    
txtID.Text = RTrim(rs!id)
'dtp1 = rs!FechaAlta
txtDescripcion.Text = RTrim(rs!Descripcion)
txtMarca.Text = RTrim(rs!Marca)
txtTalle.Text = RTrim(rs!Talle)
txtPrecio.Text = Replace(Format(rs!Precio, "fixed"), ",", ".")
txtPrecioProv.Text = Replace(Format(rs!PrecioProv, "fixed"), ",", ".")
txtExistencias = rs!Existencias
txtStockMinimo = rs!StockMinimo

End Sub


' ################ G U A R D A R #############################
'##############################################################
Private Sub Guardar()

If txtID <> "" Then


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


RefreshGrid MSHFlexGrid1, "Articulos"

End Sub


Private Sub ActualizarRegistro()

strSql = "SELECT * FROM Articulos WHERE ID =" & Val(txtID)

rs.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
        'Si no existe que lo agregue
        If rs.BOF And rs.EOF Then
        rs.AddNew
        GuardarRegistros
        ActualizarDescripcionEnPreciosClientes
        Else
        GuardarRegistros
        ActualizarDescripcionEnPreciosClientes
        End If

End Sub


Private Sub GuardarRegistros()
On Error Resume Next

rs!id = txtID
'rs!FechaAlta = dtp1
rs!Descripcion = IIf((txtDescripcion = ""), "---", txtDescripcion)
rs!Marca = IIf((txtMarca = ""), "---", txtMarca)
rs!Talle = IIf((txtTalle = ""), "---", txtTalle)
rs!Precio = Val(txtPrecio)
rs!PrecioProv = Val(txtPrecioProv)
rs!Existencias = Val(txtExistencias)
rs!StockMinimo = Val(txtStockMinimo)
rs.Update

End Sub

Private Sub ActualizarDescripcionEnPreciosClientes()
strSql = "Select * From PreciosClientes Where CodArticulo=" & Val(txtID)
Dim rs2 As New Recordset
rs2.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs2.EOF Then
While Not rs2.EOF
rs2!Descripcion = txtDescripcion
rs2!Marca = txtMarca
rs2.MoveNext
Wend
End If
End Sub

'################# Habilitar / Deshabilitar ################
Private Sub HabilitarNuevo()
    
    Dim saveNroSoc As Variant
    
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
    
txtExistencias.Enabled = False
    'Habilito todos los Combobox
    
'    For i = 1 To Me.Controls.Count - 1
'            If TypeOf Me.Controls(i) Is ComboBox Then
'                Me.Controls(i).Enabled = True
'            End If
'    Next i
    
' Deshabilito el campo clave
'txtID.Enabled = False
'Y pongo el Foco en el Campo Nombre
txtDescripcion.SetFocus
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
    
        For i = 1 To Me.Controls.Count - 1
        If TypeOf Me.Controls(i) Is ComboBox Then
            Me.Controls(i).Visible = False
        End If
        Next i

    
    txtBuscar.Enabled = True
    comboCriterio.Visible = True
    comboCriterio.Enabled = True

    
    ' Después habilito el campo de número de
    ' historia clínica y le coloco el foco
    txtID.Locked = False
    txtID.Enabled = True
    txtID.SetFocus
   BotoneraNeutra
End Sub




' ####################### KeyPress  ##########################


Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtDescripcion = "" Then
DeshabilitarEdicion
KeyAscii = 0
End If
End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtMarca = "" Then
txtDescripcion.SetFocus
KeyAscii = 0
End If
End Sub
'
Private Sub txtTalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtTalle = "" Then
txtMarca.SetFocus
KeyAscii = 0
End If
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtPrecio = "" Then
txtTalle.SetFocus
KeyAscii = 0
End If
End Sub

'
Private Sub txtPrecioProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtPrecioProv = "" Then
txtPrecio.SetFocus
KeyAscii = 0
End If
End Sub
Private Sub txtExistencias_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{tab}")
KeyAscii = 0
ElseIf KeyAscii = 8 And txtExistencias = "" Then
txtPrecioProv.SetFocus
KeyAscii = 0
End If
End Sub
Private Sub txtStockMinimo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Guardar
KeyAscii = 0
ElseIf KeyAscii = 8 And txtStockMinimo = "" Then
txtExistencias.SetFocus
KeyAscii = 0
End If
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

'CmdActualizar.SetFocus

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
Private Sub LimpioFlex()
titulos = " Codigo.|Descripción|Marca|Talle|Precio Unit.|Precio Prov.|Existencias|Stock Minimo"
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
        .ColWidth(1) = 2500
        .ColWidth(2) = 2500
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
'        .ColAlignment(1) = 3
'        .ColAlignment(3) = 3
'        .ColAlignment(4) = 3
        .ColAlignmentFixed = 3
    End With
End Sub
Private Sub AceptarBuscar(ByVal txtBuscar As Variant)
LimpioFlex
AbrirBase

Dim strSql As String
strSql = "Select * From Articulos"

Dim rsRef As New Recordset
rsRef.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText



Select Case comboCriterio.Text

Case "Descripcion"
    If txtBuscar <> "" Then
    rsRef.Filter = "Descripcion LIKE '*" + txtBuscar + "*'"
    End If


Case "Codigo"
    If txtBuscar <> "" Then
    rsRef.Filter = "ID =" & Val(txtBuscar)
    End If

Case "Marca"

    If txtBuscar <> "" Then
    rsRef.Filter = "Marca LIKE '*" + txtBuscar + "*'"
    End If

Case Else
rsRef.Filter = ""

End Select

    Do While Not rsRef.EOF

        i = i + 1
       
        linea = rsRef!id _
                & Chr(9) & rsRef!Descripcion _
                & Chr(9) & rsRef!Marca _
                & Chr(9) & rsRef!Talle _
                & Chr(9) & Format(rsRef!Precio, "fixed") _
                & Chr(9) & Format(rsRef!PrecioProv, "fixed") _
                & Chr(9) & Val(rsRef!Existencias) _
                & Chr(9) & Val(rsRef!StockMinimo) _

    MSHFlexGrid1.AddItem linea, i
    rsRef.MoveNext
    Loop
lblTotalRef = rsRef.RecordCount
CerrarBase
End Sub


Private Sub cmdBuscar_Click()
AceptarBuscar txtBuscar.Text
End Sub



Private Sub MSHFlexGrid1_Click()
MSHFlexGrid1.Col = 0
txtID = MSHFlexGrid1.Text
txtID.Locked = True

Aceptar txtID.Text
GeneroBarCode txtID

End Sub




'##################### IMPRESION #######################





Private Sub ImprimirListado()


AbrirBase
ImprimirSeleccion
CerrarBase
End Sub



Private Sub ImprimirSeleccion()
    
 '      cd1.ShowPrinter
 
    
    Dim rsArticulos As New Recordset
    rsArticulos.CursorLocation = adUseClient
    
    
strSql = "Select * From Articulos ORDER BY ID"

    rsArticulos.Open strSql, DB, adOpenDynamic, adLockOptimistic
    ImpTitulos
    
    ContRay = 0
    ContLin = 0
    espacioCelda = 0.4
    
    
    Do While Not rsArticulos.EOF
                   
If ContRay > 0 Then
Printer.CurrentY = Printer.CurrentY + espacioCelda
End If
                     
Dim CadenaCod As String
Dim LongitudCod As Long
CadenaCod = Format(CadenaCod, "0###")
LongitudCod = Len(CadenaCod)

Printer.CurrentX = 1.8 + (Val(LongitudCod)) - (Printer.TextWidth(Format(rsArticulos!id, "0###")))
Printer.Print Format(rsArticulos!id, "0###");


       
Printer.CurrentX = 2.1
Printer.Print Left(rsArticulos!Descripcion, 30);

Printer.CurrentX = 9.1
Printer.Print Left(rsArticulos!Marca, 30);
Printer.CurrentX = 13.1
Printer.Print Left(rsArticulos!Talle, 30);
    
       

Dim Cadena2 As String
Dim Longitud2 As Long
Cadena2 = Format(Cadena2, "#,###.#0")
Longitud2 = Len(Cadena2)
Printer.CurrentX = 19.3 + (Val(Longitud2)) - (Printer.TextWidth(Format(rsArticulos!Precio, "#,###.#0")))
'Printer.CurrentX = 17.3 + (Val(Longitud2)) - (Printer.TextWidth(Format(rsArticulos!Precio, "#,###.#0")))
Printer.Print Format(rsArticulos!Precio, "#,###.#0");

'Dim Cadena3 As String
'Dim Longitud3 As Long
'Cadena3 = Format(Cadena3, "#,###.#0")
'Longitud3 = Len(Cadena3)
'Printer.CurrentX = 19.3 + (Val(Longitud3)) - (Printer.TextWidth(Format(rsArticulos!PrecioDocena, "#,###.#0")))
'Printer.Print Format(rsArticulos!PrecioDocena, "#,###.#0");


       ContRay = ContRay + 1
       ContLin = ContLin + 1
       
       If Printer.CurrentY > 28 Then
          Printer.NewPage
          ImpTitulos
          ContLin = 0
          ContRay = 0
       End If
       
       If ContRay = 5 Then
       'Printer.DrawStyle = 2
           Printer.CurrentY = Printer.CurrentY + espacioCelda
       '    Printer.Line (1, Printer.CurrentY)-(20, Printer.CurrentY)
           ContRay = 0
       End If
       rsArticulos.MoveNext
   Loop
    Printer.Line (1, Printer.CurrentY + espacioCelda)-(20, Printer.CurrentY + espacioCelda)
    Printer.EndDoc
'MsgBox "Lista Imprimiendose", vbInformation
End Sub



Private Sub ImpTitulos()

Titulo = "Lista de Precios"

    'Imprime el Template de la hoja
Printer.ScaleMode = vbCentimeters
Printer.PaperSize = 9
Printer.Orientation = 1
Printer.FontName = "ARIAL"
Printer.DrawStyle = 0
Printer.Font.Size = 14

Printer.PaintPicture frmLogo.Picture1.Picture, 1, 1

Printer.Font.Size = 8

Printer.CurrentX = 18.8
Printer.CurrentY = 1
Printer.Print "Página " & Printer.Page


Printer.CurrentX = 18.5
Printer.CurrentY = 1.5
Printer.Print Date '& " " & Format(Time, "hh:nn am/pm")

Printer.CurrentX = 1
Printer.CurrentY = 2
'Printer.Print "TEL:4744-4416 - E-mail: dyvcom@hotmail.com"


Printer.Font.Size = 14

x = 21 / 2
Printer.CurrentX = x - (Printer.TextWidth(Titulo)) + (Printer.TextWidth(Titulo)) / 2
Printer.CurrentY = 3
Printer.Print Titulo

Printer.Font.Size = 8

'LINEAS HORIZONTALES
Printer.Line (1, 3.9)-(20, 3.9)
Printer.Line (1, 29)-(20, 29)

'LINEAS VERTICALES
Printer.Line (1, 3.9)-(1, 29)
'Aqui codigo
Printer.Line (2, 3.9)-(2, 29)
'Aqui Descripcion
Printer.Line (9, 3.9)-(9, 29)
'Aqui colores
Printer.Line (13, 3.9)-(13, 29)
'Aqui Talles
'Printer.Line (16, 3.9)-(16, 29)
'Aqui p.unit
Printer.Line (18, 3.9)-(18, 29)
'Aqui p. docena
Printer.Line (20, 3.9)-(20, 29)

' TITULO DE LOS CAMPOS

Printer.CurrentY = 4


Printer.CurrentX = 1.2
Printer.Print "ART.";



Printer.CurrentX = 4.2
Printer.Print "DESCRIPCIÓN";

Printer.CurrentX = 10.2
Printer.Print "Color";
'
Printer.CurrentX = 13.9
Printer.Print "TALLES";
'
'Printer.CurrentX = 16.5
'Printer.Print "P/UNIT";


Printer.CurrentX = 18.3
Printer.Print "Precio U."


'Printer.CurrentX = 14
'Printer.Print "TotalVenta";

'Printer.CurrentX = 17.5
'Printer.Print "Costo.M.";

'Printer.CurrentX = 19.2
'Printer.Print "Utilidad"
    
Printer.Line (1, (Printer.CurrentY + espacioCelda))-(20, (Printer.CurrentY + espacioCelda))

End Sub







Private Sub ImprimirInventario()

AbrirBase
ImprimirINV
CerrarBase

End Sub

Private Sub ImprimirINV()

  Dim rsArticulos As New Recordset
    rsArticulos.CursorLocation = adUseClient
    
    
strSql = "Select * From Articulos ORDER BY ID"

    rsArticulos.Open strSql, DB, adOpenDynamic, adLockOptimistic
    'Label1.Caption = RsArticulos.RecordCount
    
    ImpTitulosINV
    
    ContRay = 0
    ContLin = 0
    espacioCelda = 0.4
    
    
Do While Not rsArticulos.EOF
                   
If ContRay > 0 Then
Printer.CurrentY = Printer.CurrentY + espacioCelda
End If
                     
Dim CadenaCod As String
Dim LongitudCod As Long
CadenaCod = Format(CadenaCod, "0###")
LongitudCod = Len(CadenaCod)

Printer.CurrentX = 1.8 + (Val(LongitudCod)) - (Printer.TextWidth(Format(rsArticulos!id, "0###")))
Printer.Print Format(rsArticulos!id, "0###");


       
Printer.CurrentX = 2.1
Printer.Print Left(rsArticulos!Descripcion, 30);

Printer.CurrentX = 9.1
Printer.Print Left(rsArticulos!Marca, 30);


'Printer.CurrentX = 18.4
'Printer.Print Val(rsArticulos!Existencias);


'Printer.CurrentX = 9.1
'Printer.Print Left(rsArticulos!Color, 30);
Printer.CurrentX = 13.1
Printer.Print Left(rsArticulos!Talle, 30);
'
'Dim Cadena2 As String
'Dim Longitud2 As Long
'Cadena2 = Format(Cadena2, "#,###.#0")
'Longitud2 = Len(Cadena2)
'Printer.CurrentX = 17.3 + (Val(Longitud2)) - (Printer.TextWidth(Format(rsArticulos!Precio, "#,###.#0")))
'Printer.Print Format(rsArticulos!Precio, "#,###.#0");
'
Dim Cadena3 As String
Dim Longitud3 As Long
Cadena3 = Format(Cadena3, "0###")
Longitud3 = Len(Cadena3)
Printer.CurrentX = 19.3 + (Val(Longitud3)) - (Printer.TextWidth(Format(rsArticulos!Existencias, "0###")))
Printer.Print Format(rsArticulos!Existencias, "0###");


       ContRay = ContRay + 1
       ContLin = ContLin + 1
       
       If Printer.CurrentY > 28 Then
          Printer.NewPage
          ImpTitulos
          ContLin = 0
          ContRay = 0
       End If
       
       If ContRay = 5 Then
       'Printer.DrawStyle = 2
           Printer.CurrentY = Printer.CurrentY + espacioCelda
       '    Printer.Line (1, Printer.CurrentY)-(20, Printer.CurrentY)
           ContRay = 0
       End If
       rsArticulos.MoveNext
   Loop
    Printer.Line (1, Printer.CurrentY + espacioCelda)-(20, Printer.CurrentY + espacioCelda)
    Printer.EndDoc
'MsgBox "Lista Imprimiendose", vbInformation
End Sub

Private Sub ImpTitulosINV()

Titulo = "Control de Inventario"


    'Imprime el Template de la hoja
Printer.ScaleMode = vbCentimeters
Printer.PaperSize = 9
Printer.Orientation = 1
Printer.FontName = "ARIAL"
Printer.DrawStyle = 0
Printer.Font.Size = 14

'Printer.CurrentX = 1
'Printer.CurrentY = 1
'Printer.Print "DYVCOM S.R.L."

Printer.PaintPicture frmLogo.Picture1.Picture, 1, 1

Printer.Font.Size = 8

Printer.CurrentX = 18.8
Printer.CurrentY = 1
Printer.Print "Página " & Printer.Page


Printer.CurrentX = 18.5
Printer.CurrentY = 1.5
Printer.Print Date '& " " & Format(Time, "hh:nn am/pm")

Printer.CurrentX = 1
Printer.CurrentY = 2
'Printer.Print "TEL:4744-4416 - E-mail: dyvcom@hotmail.com"


Printer.Font.Size = 14

x = 21 / 2
Printer.CurrentX = x - (Printer.TextWidth(Titulo)) + (Printer.TextWidth(Titulo)) / 2
Printer.CurrentY = 3
Printer.Print Titulo

Printer.Font.Size = 8

'LINEAS HORIZONTALES
Printer.Line (1, 3.9)-(20, 3.9)
Printer.Line (1, 29)-(20, 29)

'LINEAS VERTICALES
Printer.Line (1, 3.9)-(1, 29)
'Aqui codigo
Printer.Line (2, 3.9)-(2, 29)
'Aqui Descripcion
Printer.Line (9, 3.9)-(9, 29)
'Aqui colores
Printer.Line (13, 3.9)-(13, 29)
'Aqui Talles
Printer.Line (16, 3.9)-(16, 29)
'Aqui p.unit
Printer.Line (18, 3.9)-(18, 29)
'Aqui p. docena
Printer.Line (20, 3.9)-(20, 29)

' TITULO DE LOS CAMPOS

Printer.CurrentY = 4

Printer.CurrentX = 1.2
Printer.Print "ART.";

Printer.CurrentX = 4.2
Printer.Print "DESCRIPCIÓN";

Printer.CurrentX = 10.2
Printer.Print "Color";

Printer.CurrentX = 13.9
Printer.Print "TALLES";
'
'Printer.CurrentX = 16.5
'Printer.Print "P/UNIT";
'
Printer.CurrentX = 18.3
Printer.Print "Existencias"


'Printer.CurrentX = 14
'Printer.Print "TotalVenta";

'Printer.CurrentX = 17.5
'Printer.Print "Costo.M.";

'Printer.CurrentX = 19.2
'Printer.Print "Utilidad"
    
Printer.Line (1, (Printer.CurrentY + espacioCelda))-(20, (Printer.CurrentY + espacioCelda))



End Sub



Private Sub ImprimirFicha()

AbrirBase
Titulo = "Detalle de " & txtDescripcion


    'Imprime el Template de la hoja
Printer.ScaleMode = vbCentimeters
Printer.PaperSize = 9
Printer.Orientation = 1
Printer.FontName = "ARIAL"
Printer.DrawStyle = 0
Printer.Font.Size = 14

'Printer.CurrentX = 1
'Printer.CurrentY = 1
'Printer.Print "DYVCOM S.R.L."

Printer.PaintPicture frmLogo.Picture1.Picture, 1, 1

Printer.Font.Size = 8

Printer.CurrentX = 18.8
Printer.CurrentY = 1
Printer.Print "Página " & Printer.Page


Printer.CurrentX = 18.5
Printer.CurrentY = 1.5
Printer.Print Date '& " " & Format(Time, "hh:nn am/pm")

Printer.CurrentX = 1
Printer.CurrentY = 2
'Printer.Print "TEL:4744-4416 - E-mail: dyvcom@hotmail.com"


Printer.Font.Size = 14

x = 21 / 2
Printer.CurrentX = x - (Printer.TextWidth(Titulo)) + (Printer.TextWidth(Titulo)) / 2
Printer.CurrentY = 3
Printer.Print Titulo


Printer.Line (1, 4)-(20, 4)

Printer.Font.Size = 12
Printer.CurrentX = 2
Printer.CurrentY = 4
Printer.Print "Codigo de Artículo = " & txtID
Printer.CurrentX = 2
Printer.CurrentY = Printer.CurrentY + 1
Printer.Print "Fecha de Alta = " & dtp1
Printer.CurrentX = 2
Printer.CurrentY = Printer.CurrentY + 1
Printer.Print "Descripción = " & txtDescripcion
Printer.CurrentX = 2
Printer.CurrentY = Printer.CurrentY + 1
Printer.Print "MARCA = " & txtMarca
Printer.CurrentX = 2
Printer.CurrentY = Printer.CurrentY + 1
Printer.Print "Talle = " & txtTalle
Printer.CurrentX = 2
Printer.CurrentY = Printer.CurrentY + 1
Printer.Print "Precio Unitario = " & txtPrecio
Printer.CurrentX = 2
Printer.CurrentY = Printer.CurrentY + 1
Printer.Print "Precio x 12 = " & txtPrecioProv
Printer.CurrentX = 2
Printer.CurrentY = Printer.CurrentY + 1
Printer.Print "Existencias = " & txtExistencias
Printer.CurrentX = 2
Printer.CurrentY = Printer.CurrentY + 1
Printer.Print "Stock Mínimo = " & txtStockMinimo


Printer.Line (1, (Printer.CurrentY + 1))-(20, (Printer.CurrentY + 1))

Printer.EndDoc
CerrarBase
End Sub
