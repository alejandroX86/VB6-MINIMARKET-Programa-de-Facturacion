VERSION 5.00
Begin VB.Form frmCambioClave 
   BackColor       =   &H00000000&
   Caption         =   "Modulo Cambio de Clave"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   Icon            =   "frmCambioClave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese el Usuario y Tipo de Clave:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtClaveAnterior 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtClaveNueva 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtConfirmaClaveNueva 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Confirmar"
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
         Left            =   1740
         TabIndex        =   4
         Top             =   1920
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cerrar"
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
         Left            =   3300
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
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
         Left            =   180
         TabIndex        =   6
         Top             =   1920
         Width           =   1395
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Clave Anterior:"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Clave Nueva:"
         Height          =   255
         Left            =   900
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Confirma Clave Nueva:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Image Image4 
         Height          =   600
         Left            =   3780
         Picture         =   "frmCambioClave.frx":57E2
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   1260
         TabIndex        =   8
         Top             =   420
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCambioClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Validar() Then
confirmar

Else
MsgBox "ERROR, Verifique los datos ingresados", vbCritical

End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Function Validar() As Boolean
Validar = False

If txtClaveNueva = txtConfirmaClaveNueva Then
Validar = True
Exit Function
Else
Validar = False
Exit Function
End If


End Function

Private Sub confirmar()
AbrirBase
strSql = "SELECT * FROM Usuarios Where Pass Like " & "'" & txtClaveAnterior & "'"

strSql = "SELECT * FROM Usuarios Where Usuario like " & "'" & txtUsuario & "'" & " AND Clave Like " & "'" & txtClaveAnterior & "'"



rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rs.EOF Then
rs!Clave = txtClaveNueva
rs.Update
MsgBox "La Clave ha Cambiado", vbInformation
Unload Me
Else
MsgBox "ERROR, Verifique los datos ingresados", vbCritical
End If
CerrarBase
End Sub

Private Sub Command3_Click()
On Error Resume Next
    For i = 1 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is TextBox Then
                Me.Controls(i).Text = ""
            End If
    Next i
txtUsuario.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
