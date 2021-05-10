VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso al Sistema"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ingrese su Usuario y Clave de Acceso:"
      Height          =   1335
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1140
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "Admin"
         Top             =   780
         Width           =   1995
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "Admin"
         Top             =   360
         Width           =   1995
      End
      Begin VB.Image Image4 
         Height          =   600
         Left            =   3240
         Picture         =   "frmLogin.frx":57E2
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FFFF&
         Caption         =   "Clave:"
         Height          =   255
         Left            =   540
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   420
         TabIndex        =   3
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.TextBox txtAcceso 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2340
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
HacerLogin
End Sub



Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPassword.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtPassWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar_Click
        KeyAscii = 0
    End If
End Sub


Private Sub HacerLogin()
'On Error Resume Next
strSql = "SELECT * FROM Usuarios Where Usuario like " & "'" & txtUsuario & "'" & " AND Clave Like " & "'" & txtPassword & "'"
AbrirBase
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then
'ValidarPermisos
CerrarBase
frmAdvertencia.Show
Unload Me
Else
MsgBox "Usuario o Clave Incorrecta", vbCritical
CerrarBase
End If
End Sub

Private Sub ValidarPermisos()


If rs!Permiso = "Usuario del Sistema" Then
frmSwitch.Command3.Visible = False
frmSwitch.Command4.Visible = False
'frmSwitch.Command10.Visible = False
frmSwitch.Command11.Visible = False
frmSwitch.Command12.Visible = False
frmSwitch.Command13.Visible = False
frmSwitch.Command15.Visible = False
frmSwitch.lblUsuario = "Usuario del Sistema"
ElseIf rs!Permiso = "Administrador del Sistema" Then
frmSwitch.lblUsuario = "Administrador del Sistema"
Else
MsgBox "Error en permisos de Acceso, Comuniquese con el Programador del sistema", vbCritical
End If



End Sub

