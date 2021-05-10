VERSION 5.00
Begin VB.Form frmLogo 
   BackColor       =   &H00000000&
   Caption         =   "frmLogo"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmLogo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   1395
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1150
      Left            =   1680
      Picture         =   "frmLogo.frx":57E2
      ScaleHeight     =   1125
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Aca va el logo de la empresa para que salga impreso en todos los formularios y reportes:"
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

