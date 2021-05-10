VERSION 5.00
Begin VB.Form frmCordPrint 
   BackColor       =   &H00000000&
   Caption         =   "Cordenadas de Impresión"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   Icon            =   "frmCordPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   3765
      Left            =   3915
      TabIndex        =   16
      Top             =   315
      Width           =   3390
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Impresión:"
      Height          =   1950
      Left            =   180
      TabIndex        =   3
      Top             =   4185
      Width           =   7125
      Begin VB.CheckBox Check1 
         Caption         =   "Visible"
         Height          =   240
         Left            =   3330
         TabIndex        =   15
         Top             =   1215
         Width           =   1680
      End
      Begin VB.ComboBox ComboFont 
         Height          =   315
         Left            =   1755
         TabIndex        =   7
         Text            =   "Arial"
         Top             =   630
         Width           =   1455
      End
      Begin VB.ComboBox ComboSize 
         Height          =   315
         ItemData        =   "frmCordPrint.frx":57E2
         Left            =   3285
         List            =   "frmCordPrint.frx":5804
         TabIndex        =   6
         Text            =   "9"
         Top             =   630
         Width           =   735
      End
      Begin VB.TextBox txtCordX 
         Height          =   285
         Left            =   1755
         MaxLength       =   5
         TabIndex        =   5
         Top             =   990
         Width           =   870
      End
      Begin VB.TextBox txtCordY 
         Height          =   285
         Left            =   1755
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1350
         Width           =   870
      End
      Begin VB.Label lblCurrentList 
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
         Height          =   195
         Left            =   4545
         TabIndex        =   18
         Top             =   315
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label Label1 
         Caption         =   "Letra/Tamaño:"
         Height          =   285
         Left            =   315
         TabIndex        =   14
         Top             =   675
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Pocicion H:"
         Height          =   240
         Left            =   315
         TabIndex        =   13
         Top             =   1035
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Pocicion V:"
         Height          =   240
         Left            =   315
         TabIndex        =   12
         Top             =   1395
         Width           =   1500
      End
      Begin VB.Label lblCurrentTag 
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
         Height          =   195
         Left            =   1755
         TabIndex        =   11
         Top             =   315
         Width           =   2265
      End
      Begin VB.Label Label4 
         Caption         =   "cm."
         Height          =   195
         Left            =   2700
         TabIndex        =   10
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label Label5 
         Caption         =   "cm."
         Height          =   195
         Left            =   2700
         TabIndex        =   9
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label Label6 
         Caption         =   "Etiqueta:"
         Height          =   240
         Left            =   315
         TabIndex        =   8
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   330
      Left            =   5580
      TabIndex        =   2
      Top             =   6390
      Width           =   1500
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   330
      Left            =   4005
      TabIndex        =   1
      Top             =   6390
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   3345
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      Caption         =   "Remitos:"
      Height          =   195
      Left            =   3960
      TabIndex        =   19
      Top             =   45
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      Caption         =   "Facturas:"
      Height          =   195
      Left            =   225
      TabIndex        =   17
      Top             =   45
      Width           =   870
   End
End
Attribute VB_Name = "frmCordPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdGuardar_Click()




GuardarConfig lblCurrentList, lblCurrentTag
MsgBox ("registro actualzado!"), vbInformation
End Sub

Private Sub Form_Load()
ListCordTypes
ListarFuentes
End Sub

Private Sub ListCordTypes()
AbrirBase

strSql = "Select * From CordFact Order By ID ASC"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then
While Not rs.EOF
List1.AddItem rs!Type
rs.MoveNext
Wend
End If
rs.Close

strSql = "Select * From CordRem Order By ID ASC"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then
While Not rs.EOF
List2.AddItem rs!Type
rs.MoveNext
Wend
End If
rs.Close

CerrarBase
End Sub
Private Sub ListarFuentes()
  Dim i As Long
  For i = 0 To Screen.FontCount - 1
  ComboFont.AddItem CStr(Screen.Fonts(i))
  Next i
End Sub


Private Sub List1_Click()
TraerConfig "CordFact", List1.List(List1.ListIndex)
lblCurrentList = "CordFact"
lblCurrentTag = List1.List(List1.ListIndex)
End Sub

Private Sub List2_Click()
TraerConfig "CordRem", List2.List(List2.ListIndex)
lblCurrentList = "CordRem"
lblCurrentTag = List2.List(List2.ListIndex)
End Sub


Private Sub TraerConfig(tabla, tipo)
AbrirBase
strSql = "Select * From " & tabla & " Where Type like " & "'" & tipo & "'"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then
ComboFont = "" & rs!Font
ComboSize = "" & rs!Size
txtCordX = "" & Str(rs!CordX)
txtCordY = "" & Str(rs!CordY)
Check1.Value = rs!Visible
End If
CerrarBase


If Check1.Value = 1 Then
lblCurrentTag.ForeColor = vbBlue
Else
lblCurrentTag.ForeColor = vbRed
End If

End Sub


Private Sub GuardarConfig(tabla, tipo)
If tabla <> "" Then
AbrirBase
strSql = "Select * From " & tabla & " Where Type like " & "'" & tipo & "'"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then
rs!Font = "" & ComboFont
rs!Size = "" & ComboSize
rs!CordX = Val(txtCordX)
rs!CordY = Val(txtCordY)
rs!Visible = Check1.Value
rs.Update
End If
CerrarBase

End If
End Sub

