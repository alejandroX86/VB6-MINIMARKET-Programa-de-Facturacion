VERSION 5.00
Begin VB.Form frmListConectObject 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   Icon            =   "frmListConectObject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   330
      Left            =   4095
      TabIndex        =   5
      Top             =   5535
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   330
      Left            =   6390
      TabIndex        =   4
      Top             =   5535
      Width           =   2085
   End
   Begin VB.ListBox List2 
      Height          =   4545
      Left            =   4680
      TabIndex        =   2
      Top             =   675
      Width           =   3840
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   45
      TabIndex        =   1
      Top             =   675
      Width           =   3705
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   180
      Width           =   3705
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Productos sin alicuota asignada"
      Height          =   285
      Left            =   4725
      TabIndex        =   3
      Top             =   180
      Width           =   3750
   End
End
Attribute VB_Name = "frmListConectObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'botones
Private Sub cmdGuardar_Click()
AbrirBase
GuardarConexiones
CerrarBase
MsgBox "Conexiones Guardadas", vbInformation
End Sub
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub GuardarConexiones()

strSql = "Select * FROM Impuestos WHERE Descripcion LIKE " & "'" & Combo1 & "'"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
'elimino todos los conectados al actual
Dim rs1 As New Recordset
strSql = "DELETE FROM cnTaxProduct Where IDImpuesto=" & Val(rs!id)
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic

'vuelvo a conectar con todos los de la lista
Dim strDescr As String
For i = 0 To List1.ListCount - 1

strDescr = List1.List(i)


    Dim rs2 As New Recordset
    strSql = "Select * FROM Articulos Where Descripcion Like " & "'" & strDescr & "'"
    rs2.Open strSql, DB, adOpenKeyset, adLockOptimistic
    If Not rs2.EOF Then
    strSql = "Select * FROM cnTaxProduct"
        Dim rs3 As New Recordset
        rs3.Open strSql, DB, adOpenKeyset, adLockOptimistic
        rs3.AddNew
        rs3!IDIMPUESTO = rs!id
        rs3!idArticulo = rs2!id
        rs3.Update
        rs3.Close
    
    End If
    rs2.Close
Next

End If
rs.Close
End Sub



Private Sub Combo1_Click()
List1.Clear
List2.Clear
AbrirBase
LlenarListas
CerrarBase
End Sub

Private Sub Form_Load()
Me.Caption = "Productos e Impuestos"
VerConexiones
Combo1.ListIndex = 0
End Sub

Private Sub VerConexiones()
AbrirBase
FillComboObject "Impuestos"
LlenarListas
CerrarBase
End Sub

Private Sub FillComboObject(varObject As Variant)
strSql = "Select * From " & varObject
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
While Not rs.EOF
Combo1.AddItem rs!Descripcion
rs.MoveNext
Wend
End If
rs.Close
End Sub


Private Sub LlenarListas()
'AbrirBase
FillObjetosNoConectados "Articulos", "cnTaxProduct"
FillObjetosConectadosAlActual "Articulos", "cnTaxProduct", Combo1
'CerrarBase
End Sub


Private Sub FillObjetosConectadosAlActual(varObject1 As Variant, varObject2 As Variant, Combo1 As ComboBox)


strSql = "Select * From IMPUESTOS WHERE Descripcion LIKE " & "'" & CStr(Combo1.Text) & "'"
Dim rs0 As New Recordset
rs0.Open strSql, DB, adOpenKeyset, adLockOptimistic

If Not rs0.EOF Then


strSql = "Select * From " & varObject2 & " WHERE IDIMPUESTO=" & rs0!id
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic



If Not rs.EOF Then
While Not rs.EOF


    strSql = "Select * From " & varObject1 & " Where ID=" & rs!idArticulo
    Dim rs1 As New Recordset
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then
    List1.AddItem rs1!Descripcion
    End If
    rs1.Close

rs.MoveNext
Wend
End If
rs.Close


End If
rs0.Close


End Sub




Private Sub FillObjetosNoConectados(varObject1 As Variant, varObject2 As Variant)

strSql = "Select * From " & varObject1
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
While Not rs.EOF


    strSql = "Select * From " & varObject2 & " Where IDArticulo=" & rs!id
    Dim rs1 As New Recordset
    rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic
    If rs1.EOF Then
    List2.AddItem rs!Descripcion
    End If
    rs1.Close

rs.MoveNext
Wend
End If
rs.Close


End Sub



Private Sub List1_DblClick()
PasarElemento List1, List2
End Sub


Private Sub List2_DblClick()
PasarElemento List2, List1
End Sub

Private Sub PasarElemento(List01 As ListBox, List02 As ListBox)
If Not (List01.ListCount) = -1 Then
    If List01.List(List01.ListIndex) <> "" Then
    List02.AddItem List01.List(List01.ListIndex)
    List01.RemoveItem List01.ListIndex
    End If
End If
End Sub
