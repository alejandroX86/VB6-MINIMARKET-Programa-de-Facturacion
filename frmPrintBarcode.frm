VERSION 5.00
Begin VB.Form frmPrintBarcode 
   BackColor       =   &H00000000&
   Caption         =   "Imprimir Codigo de Barras"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   Icon            =   "frmPrintBarcode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTalle 
      Height          =   285
      Left            =   180
      TabIndex        =   28
      Top             =   6120
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txtMarca 
      Height          =   285
      Left            =   180
      TabIndex        =   27
      Top             =   5805
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cordenadas:"
      Height          =   1455
      Left            =   45
      TabIndex        =   14
      Top             =   3510
      Width           =   5685
      Begin VB.TextBox txtAltoHoja 
         Height          =   285
         Left            =   4500
         MaxLength       =   5
         TabIndex        =   25
         Top             =   315
         Width           =   735
      End
      Begin VB.TextBox txtAnchoHoja 
         Height          =   285
         Left            =   1710
         MaxLength       =   5
         TabIndex        =   23
         Top             =   270
         Width           =   735
      End
      Begin VB.TextBox txtSY 
         Height          =   285
         Left            =   4500
         MaxLength       =   5
         TabIndex        =   22
         Top             =   945
         Width           =   735
      End
      Begin VB.TextBox txtSX 
         Height          =   285
         Left            =   4500
         MaxLength       =   5
         TabIndex        =   21
         Top             =   630
         Width           =   735
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   1710
         MaxLength       =   5
         TabIndex        =   20
         Top             =   945
         Width           =   735
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   1710
         MaxLength       =   5
         TabIndex        =   19
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Alto Hoja"
         Height          =   240
         Left            =   2970
         TabIndex        =   26
         Top             =   405
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Ancho Hoja:"
         Height          =   240
         Left            =   225
         TabIndex        =   24
         Top             =   315
         Width           =   1860
      End
      Begin VB.Label Label5 
         Caption         =   "Espacio Vertical:"
         Height          =   330
         Left            =   2970
         TabIndex        =   18
         Top             =   990
         Width           =   1860
      End
      Begin VB.Label Label4 
         Caption         =   "Espacio Horizontal:"
         Height          =   240
         Left            =   2970
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "PosiciónVertical"
         Height          =   240
         Left            =   225
         TabIndex        =   16
         Top             =   945
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Posición Horizontal:"
         Height          =   240
         Left            =   225
         TabIndex        =   15
         Top             =   675
         Width           =   1860
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   90
      TabIndex        =   13
      Top             =   5220
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tamaño"
      Height          =   1140
      Left            =   3780
      TabIndex        =   9
      Top             =   2295
      Width           =   1995
      Begin VB.OptionButton optSize 
         Caption         =   "Chico"
         Height          =   192
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   225
         Width           =   972
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Mediano"
         Height          =   192
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   465
         Width           =   972
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Grande"
         Height          =   192
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   705
         Width           =   972
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   420
      Left            =   3780
      TabIndex        =   8
      Top             =   5805
      Width           =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   420
      Left            =   1710
      TabIndex        =   7
      Top             =   5805
      Width           =   1995
   End
   Begin VB.TextBox txtCantidad 
      Height          =   330
      Left            =   3780
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "1"
      Top             =   5265
      Width           =   1995
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones:"
      Height          =   1140
      Left            =   45
      TabIndex        =   2
      Top             =   2295
      Width           =   3615
      Begin VB.OptionButton Option2 
         Caption         =   "Hoja Completa"
         Height          =   330
         Left            =   1980
         TabIndex        =   4
         Top             =   495
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Una bajo la otra"
         Height          =   375
         Left            =   270
         TabIndex        =   3
         Top             =   450
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   2130
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   5730
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
         Height          =   1755
         Left            =   225
         ScaleHeight     =   1725
         ScaleWidth      =   5265
         TabIndex        =   6
         Top             =   225
         Width           =   5295
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cantidad:"
      Height          =   240
      Left            =   2880
      TabIndex        =   5
      Top             =   5310
      Width           =   780
   End
End
Attribute VB_Name = "frmPrintBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
GuardarCordenadas
Imprimir
End Sub



Private Sub Imprimir()

If Option1.Value = True Then
imprimirDeA1
Else
ImprimirDeAVarias
End If

End Sub

Private Sub imprimirDeA1()
Dim Xpos As Double
Dim Ypos As Double
Dim HSpc As Double
Dim VSpc As Double
Dim AltoHoja As Double
Dim AnchoHoja As Double


AbrirBase
strSql = "Select * From Cordenadas"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
Xpos = rs!posX
Ypos = rs!posY
HSpc = rs!EspacioH
VSpc = rs!EspacioV
AltoHoja = rs!AltoHoja
AnchoHoja = rs!AnchoHoja

End If
CerrarBase

Picture1.ScaleMode = vbCentimeters
Printer.ScaleMode = vbCentimeters
Printer.CurrentX = Xpos
Printer.CurrentY = Ypos


For i = 1 To Val(txtCantidad)


Printer.PaintPicture Picture1, Printer.CurrentX, Printer.CurrentY
Printer.CurrentY = Printer.CurrentY + VSpc

If Printer.CurrentY >= AltoHoja Then
Printer.NewPage
Printer.CurrentX = Xpos
Printer.CurrentY = Ypos
End If

Next

Printer.EndDoc

End Sub

Private Sub ImprimirDeAVarias()
Dim Xpos As Double
Dim Ypos As Double
Dim HSpc As Double
Dim VSpc As Double
Dim AltoHoja As Double
Dim AnchoHoja As Double

AbrirBase
strSql = "Select * From Cordenadas"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
Xpos = rs!posX
Ypos = rs!posY
HSpc = rs!EspacioH
VSpc = rs!EspacioV
AltoHoja = rs!AltoHoja
AnchoHoja = rs!AnchoHoja
End If
CerrarBase


Picture1.ScaleMode = vbCentimeters
Printer.Copies = 1
Printer.ScaleMode = vbCentimeters
Printer.CurrentX = Xpos
Printer.CurrentY = Ypos


For i = 1 To Val(txtCantidad)


If Printer.CurrentX <= AnchoHoja Then
Printer.PaintPicture Picture1, Printer.CurrentX, Printer.CurrentY
Printer.CurrentX = Printer.CurrentX + HSpc
Else

    Printer.CurrentX = Xpos
    Printer.CurrentY = Printer.CurrentY + VSpc
    Printer.PaintPicture Picture1, Printer.CurrentX, Printer.CurrentY
    Printer.CurrentX = Printer.CurrentX + HSpc


    If Printer.CurrentY >= AltoHoja Then
    Printer.NewPage
    Printer.CurrentX = Xpos
    Printer.CurrentY = Ypos
    End If
End If




'
If Printer.CurrentY >= AltoHoja Then
Printer.NewPage
Printer.CurrentX = Xpos
Printer.CurrentY = Ypos
End If



Next

Printer.EndDoc

End Sub



Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

txtCantidad.SelLength = Len(txtCantidad)




AbrirBase
strSql = "Select * From Cordenadas"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
txtX = rs!posX
txtY = rs!posY
txtSX = rs!EspacioH
txtSY = rs!EspacioV
txtAltoHoja = rs!AltoHoja
txtAnchoHoja = rs!AnchoHoja

End If
CerrarBase







End Sub


Private Sub GuardarCordenadas()

AbrirBase
strSql = "Select * From Cordenadas"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
rs!posX = Val(txtX)
rs!posY = Val(txtY)
rs!EspacioH = Val(txtSX)
rs!EspacioV = Val(txtSY)
rs!AltoHoja = Val(txtAltoHoja)
rs!AnchoHoja = Val(txtAnchoHoja)
rs.Update
End If
rs.Close
CerrarBase

End Sub


Private Sub Form_Activate()

    optSize(1) = 1

End Sub

Private Sub optSize_Click(Index As Integer)
    Picture1.ScaleMode = 3
    
    Select Case Index
    Case 0
        Picture1.Height = Picture1.Height * (1.4 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 8
    Case 1
        Picture1.Height = Picture1.Height * (2.4 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 10
    Case 2
        Picture1.Height = Picture1.Height * (3 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 14
    End Select


    Call Text1_Change

End Sub

Private Sub Text1_Change()
    
Call DrawBarcode(Text1, Picture1, txtMarca & " - " & txtTalle)
    
    'Picture1.Picture = frmArticulos.Picture1.Picture
    
    Clipboard.Clear
Clipboard.SetData Picture1.Image, 2

'    MinWidth = 2 * Text1.Left + Text1.Width
'    pw = 2 * Picture1.Left + Picture1.Width
'    fw = MinWidth
'    If pw > fw Then fw = pw
'    Me.Width = fw

End Sub

