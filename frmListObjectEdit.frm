VERSION 5.00
Begin VB.Form frmListObjectEdit 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   Icon            =   "frmListObjectEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   330
      Left            =   4230
      TabIndex        =   1
      Top             =   2790
      Width           =   1635
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   330
      Left            =   2295
      TabIndex        =   0
      Top             =   2790
      Width           =   1635
   End
   Begin VB.Label lblObject 
      Caption         =   "lblObject"
      Height          =   240
      Left            =   45
      TabIndex        =   2
      Top             =   3105
      Visible         =   0   'False
      Width           =   2580
   End
End
Attribute VB_Name = "frmListObjectEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
GuardarDatos lblObject, Me.Tag
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub



Public Sub MapearControles(NombreTabla As String, id As Variant)

AbrirBase
strSql = "SELECT * FROM [" & NombreTabla & "]" & " WHERE ID=" & Val(id)
Dim rs1 As New Recordset
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic


For G = 0 To rs1.Fields.Count - 1

avance = avance + 300

nombrecampo = rs1.Fields.Item(G).Name
valorcampo = rs1.Fields.Item(G).Value
tipocampo = rs1.Fields.Item(G).Type
tamanocampo = rs1.Fields.Item(G).DefinedSize

Set lbl = frmListObjectEdit.Controls.Add("vb.label", "lbl" & nombrecampo)
With lbl
    .Caption = nombrecampo
    .Left = 250
    .Top = avance
    .Width = 1500
    .Height = 200
    .Visible = True
End With
Set lbl = Nothing

Set txt = frmListObjectEdit.Controls.Add("vb.textbox", "txt" & nombrecampo)
With txt
If G = 0 And id = 0 Then
.Text = 0
Else

    .Text = "" & valorcampo
        
        
    If rs1.Fields.Item(G).Type = adCurrency Then
    .Text = "" & Replace(Format(valorcampo, "standard"), ",", ".")
    End If

End If
    
    
    .Left = 2000
    .Top = avance
    .Width = 3000
    .Height = 100
    .Visible = True
    .MaxLength = tamanocampo
    
    If G = 0 Then
    .Locked = True
    End If
    
    
End With
Set txt = Nothing


Next G

Me.cmdAceptar.Top = avance + 500
Me.cmdCancelar.Top = avance + 500


CerrarBase
End Sub


Private Sub GuardarDatos(NombreTabla As String, id As Variant)
Dim aryVar As Variant
Dim aryTemp As Variant

aryVar = ""
aryTemp = ""

For i = 1 To Me.Controls.Count - 1
If TypeOf Me.Controls(i) Is TextBox Then
aryVar = aryVar & Replace(Me.Controls(i).Text, ",", ".") & ","
End If
Next i

If Len(aryVar) > 0 Then
'remover ultima comma
aryVar = Left(aryVar, (Len(aryVar) - 1))
End If



AbrirBase
strSql = "SELECT * FROM [" & NombreTabla & "]" & " WHERE ID=" & Val(id)
Dim rs1 As New Recordset
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic

If rs1.EOF Then
rs1.AddNew
End If
aryTemp = Split(aryVar, ",")


If Me.Tag = 0 Then
aryTemp(0) = ObjenerProximoID(Me.lblObject)
End If

For G = 0 To rs1.Fields.Count - 1

    If rs1.Fields.Item(G).Type = adCurrency Or rs1.Fields.Item(G).Type = adInteger Then
    rs1.Fields.Item(G).Value = Val(aryTemp(G))
    Else
    rs1.Fields.Item(G).Value = "" & aryTemp(G)
    End If

Next G
rs1.Update
rs1.Close



CerrarBase
ListObject frmListObject.flex01, lblObject, frmListObject.Image1, frmListObject.Image2

Me.Tag = ""


End Sub

Private Function ObjenerProximoID(NombreTabla As Variant) As Long

strSql = "SELECT * FROM [" & NombreTabla & "]"
Dim rs2 As New Recordset
rs2.Open strSql, DB, adOpenKeyset, adLockOptimistic

If Not rs2.EOF Then
rs2.MoveLast
x = rs2!id + 1
Else
x = 1
End If
rs2.Close

ObjenerProximoID = x

End Function

'cambiar la propiedad KeyPreview del
'formulario a True y escribir el siguiente código:
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
KeyAscii = 0
End If
End Sub

