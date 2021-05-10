VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmListObject 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10215
   Icon            =   "frmListObject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "CERRAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8370
      TabIndex        =   10
      Top             =   6360
      Width           =   1770
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "EXPORTAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5040
      TabIndex        =   9
      Top             =   6360
      Width           =   1515
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "IMPRIMIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6750
      TabIndex        =   8
      Top             =   6360
      Width           =   1515
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flex01 
      Height          =   4590
      Left            =   90
      TabIndex        =   7
      Top             =   1665
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   8096
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo..."
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   1305
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Buscador:"
      Height          =   1095
      Left            =   45
      TabIndex        =   2
      Top             =   135
      Width           =   10050
      Begin VB.ComboBox ComboBuscar 
         Height          =   315
         Left            =   5895
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   450
         Width           =   3435
      End
      Begin VB.TextBox txtBuscar 
         Height          =   330
         Left            =   1170
         TabIndex        =   0
         Top             =   450
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FFFF&
         Caption         =   "Buscar por"
         Height          =   195
         Left            =   315
         TabIndex        =   4
         Top             =   495
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Filtrar por:"
         Height          =   195
         Left            =   5085
         TabIndex        =   3
         Top             =   495
         Width           =   2220
      End
   End
   Begin VB.Label lblObject 
      BackColor       =   &H00000000&
      Caption         =   "lblObject"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   1440
      TabIndex        =   6
      Top             =   1350
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   4590
      Picture         =   "frmListObject.frx":57E2
      Top             =   1350
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4185
      Picture         =   "frmListObject.frx":5B61
      Top             =   1350
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmListObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdExportar_Click()
FlexGrid_To_Excel Me.flex01, flex01.Rows, flex01.Cols, "Listado de " & lblObject
End Sub

'imprimir
Private Sub cmdImprimir_Click()
ImprimirFlex flex01, "Listado de " & lblObject
End Sub
Private Sub Command1_Click()
Dim varid As Long
varid = 0
frmListObjectEdit.MapearControles lblObject, varid
frmListObjectEdit.lblObject = Me.lblObject
frmListObjectEdit.Caption = Me.lblObject
frmListObjectEdit.Tag = varid
frmListObjectEdit.Show 1
End Sub


Private Sub flex01_DblClick()
Dim columna As Integer
Dim varid As Long
columna = Me.flex01.Col
flex01.Col = 2
varid = Val(flex01.Text)
'delete
If columna = 1 Then

    If MsgBox("Atención: Eliminar este registro?", vbExclamation + vbYesNoCancel) = vbYes Then
    EliminarLinea lblObject, varid
    ListObject flex01, lblObject, Image1, Image2
    End If

Else
'edit
frmListObjectEdit.MapearControles lblObject, varid
frmListObjectEdit.lblObject = Me.lblObject
frmListObjectEdit.Caption = Me.lblObject
frmListObjectEdit.Tag = varid
frmListObjectEdit.Show 1
End If

End Sub



'################PROCEDIMIENTO BUSCAR ####################

Private Sub txtBuscar_Change()
BuscarRegistro Me.flex01, lblObject, Image1, Image2, txtBuscar, ComboBuscar
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        BuscarRegistro Me.flex01, lblObject, Image1, Image2, txtBuscar, ComboBuscar
    End If
End Sub


Private Sub cmdBuscar_Click()
BuscarRegistro Me.flex01, lblObject, Image1, Image2, txtBuscar, ComboBuscar
End Sub
