VERSION 5.00
Begin VB.Form frmAdvertencia 
   BackColor       =   &H00000000&
   Caption         =   "Bienvenido a MINIMARKET"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   Icon            =   "frmAdvertencia.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmAdvertencia.frx":57E2
   ScaleHeight     =   5955
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ir a Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   3135
   End
End
Attribute VB_Name = "frmAdvertencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMenu.Show
Unload Me
End Sub

