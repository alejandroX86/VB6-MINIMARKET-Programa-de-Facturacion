VERSION 5.00
Begin VB.Form frmConfigFacturacion 
   BackColor       =   &H00000000&
   Caption         =   "Configuración de Facturación:"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   Icon            =   "frmConfigFacturacion.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Configuración de Remitos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   44
      Top             =   3645
      Width           =   7845
      Begin VB.ComboBox ComboRem 
         Height          =   315
         Left            =   4320
         TabIndex        =   45
         Top             =   315
         Width           =   3345
      End
      Begin VB.TextBox txtRem 
         Height          =   285
         Left            =   1755
         MaxLength       =   13
         TabIndex        =   8
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label18 
         Caption         =   "Impresora:"
         Height          =   240
         Left            =   3420
         TabIndex        =   47
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label17 
         Caption         =   "Ultimo Nº Remito:"
         Height          =   285
         Left            =   225
         TabIndex        =   46
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Configuración de Recibos X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   40
      Top             =   7875
      Width           =   7845
      Begin VB.TextBox txtRecX 
         Height          =   285
         Left            =   1755
         MaxLength       =   13
         TabIndex        =   7
         Top             =   315
         Width           =   1185
      End
      Begin VB.ComboBox ComboRecX 
         Height          =   315
         Left            =   4320
         TabIndex        =   41
         Top             =   315
         Width           =   3345
      End
      Begin VB.Label Label16 
         Caption         =   "Ultimo Nº Recibo X:"
         Height          =   285
         Left            =   225
         TabIndex        =   43
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label15 
         Caption         =   "Impresora:"
         Height          =   240
         Left            =   3420
         TabIndex        =   42
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Configuración de Recibos C:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   36
      Top             =   6975
      Width           =   7845
      Begin VB.ComboBox ComboRecC 
         Height          =   315
         Left            =   4320
         TabIndex        =   37
         Top             =   315
         Width           =   3345
      End
      Begin VB.TextBox txtRecC 
         Height          =   285
         Left            =   1755
         MaxLength       =   13
         TabIndex        =   6
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label14 
         Caption         =   "Impresora:"
         Height          =   240
         Left            =   3420
         TabIndex        =   39
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label13 
         Caption         =   "Ultimo Nº Recibo C:"
         Height          =   285
         Left            =   225
         TabIndex        =   38
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Configuración de Facturas X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   32
      Top             =   2745
      Width           =   7845
      Begin VB.TextBox txtFactX 
         Height          =   285
         Left            =   1755
         MaxLength       =   13
         TabIndex        =   3
         Top             =   315
         Width           =   1185
      End
      Begin VB.ComboBox ComboFactX 
         Height          =   315
         Left            =   4320
         TabIndex        =   33
         Top             =   315
         Width           =   3345
      End
      Begin VB.Label Label12 
         Caption         =   "Ultimo Nº Factura X:"
         Height          =   285
         Left            =   225
         TabIndex        =   35
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label11 
         Caption         =   "Impresora:"
         Height          =   240
         Left            =   3420
         TabIndex        =   34
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Configuración de Facturas C:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   28
      Top             =   1845
      Width           =   7845
      Begin VB.ComboBox ComboFactC 
         Height          =   315
         Left            =   4320
         TabIndex        =   29
         Top             =   315
         Width           =   3345
      End
      Begin VB.TextBox txtFactC 
         Height          =   285
         Left            =   1755
         MaxLength       =   13
         TabIndex        =   2
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Impresora:"
         Height          =   240
         Left            =   3420
         TabIndex        =   31
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label9 
         Caption         =   "Ultimo Nº Factura C:"
         Height          =   285
         Left            =   225
         TabIndex        =   30
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Configuración de Recibos A:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   24
      Top             =   5085
      Width           =   7845
      Begin VB.TextBox txtRecA 
         Height          =   285
         Left            =   1755
         MaxLength       =   13
         TabIndex        =   4
         Top             =   315
         Width           =   1185
      End
      Begin VB.ComboBox ComboRecA 
         Height          =   315
         Left            =   4320
         TabIndex        =   25
         Top             =   315
         Width           =   3345
      End
      Begin VB.Label Label5 
         Caption         =   "Ultimo Nº Recibo A:"
         Height          =   285
         Left            =   225
         TabIndex        =   27
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label6 
         Caption         =   "Impresora:"
         Height          =   240
         Left            =   3420
         TabIndex        =   26
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Configuración de Recibos B:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   9
      Top             =   6030
      Width           =   7845
      Begin VB.TextBox txtRecB 
         Height          =   285
         Left            =   1755
         MaxLength       =   13
         TabIndex        =   5
         Top             =   315
         Width           =   1185
      End
      Begin VB.ComboBox ComboRecB 
         Height          =   315
         Left            =   4320
         TabIndex        =   21
         Top             =   315
         Width           =   3345
      End
      Begin VB.Label Label8 
         Caption         =   "Ultimo Nº Recibo B:"
         Height          =   285
         Left            =   225
         TabIndex        =   23
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "Impresora:"
         Height          =   240
         Left            =   3420
         TabIndex        =   22
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Configuración de Facturas B:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   17
      Top             =   945
      Width           =   7845
      Begin VB.TextBox txtFactB 
         Height          =   285
         Left            =   1755
         MaxLength       =   13
         TabIndex        =   1
         Top             =   315
         Width           =   1185
      End
      Begin VB.ComboBox ComboFactB 
         Height          =   315
         Left            =   4320
         TabIndex        =   18
         Top             =   315
         Width           =   3345
      End
      Begin VB.Label Label4 
         Caption         =   "Ultimo Nº Factura B:"
         Height          =   285
         Left            =   225
         TabIndex        =   20
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Impresora:"
         Height          =   240
         Left            =   3420
         TabIndex        =   19
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración de Facturas A:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   13
      Top             =   45
      Width           =   7845
      Begin VB.ComboBox ComboFactA 
         Height          =   315
         Left            =   4320
         TabIndex        =   14
         Top             =   315
         Width           =   3345
      End
      Begin VB.TextBox txtFactA 
         Height          =   285
         Left            =   1755
         MaxLength       =   13
         TabIndex        =   0
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Impresora:"
         Height          =   240
         Left            =   3420
         TabIndex        =   16
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Ultimo Nº Factura A:"
         Height          =   285
         Left            =   225
         TabIndex        =   15
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.PictureBox cmdAceptar 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   8685
      TabIndex        =   10
      Top             =   660
      Width           =   8745
   End
   Begin VB.PictureBox cmdCancelar 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   8685
      TabIndex        =   11
      Top             =   330
      Width           =   8745
   End
   Begin VB.PictureBox cmdAplicar 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   8685
      TabIndex        =   12
      Top             =   0
      Width           =   8745
   End
End
Attribute VB_Name = "frmConfigFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
GuardarTodo
Unload Me
End Sub

Private Sub cmdAplicar_Click()
GuardarTodo
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub



Private Sub Form_Load()
Me.Height = 5460
Me.Width = 8085
Me.Left = 0
Me.Top = 0
'SetDeviceIndependentWindow Me
AbrirBase
traerNumeros
TraerImpresoras
CerrarBase

End Sub

Private Sub traerNumeros()

txtFactA = TraerUltimaFactura("FacturasA", txtFactA)
txtFactB = TraerUltimaFactura("FacturasB", txtFactB)
txtFactC = TraerUltimaFactura("FacturasC", txtFactC)
txtFactX = TraerUltimaFactura("FacturasX", txtFactX)

txtRecA = TraerUltimoRecibo("RecibosA", txtRecA)
txtRecB = TraerUltimoRecibo("RecibosB", txtRecB)
txtRecC = TraerUltimoRecibo("RecibosC", txtRecC)
txtRecX = TraerUltimoRecibo("RecibosX", txtRecX)

txtRem = TraerUltimoRemito("Remitos", txtRem)

ComboFactA = TraerImpresora("FacturasA")
ComboFactB = TraerImpresora("FacturasB")
ComboFactC = TraerImpresora("FacturasC")
ComboFactX = TraerImpresora("FacturasX")

ComboRecA = TraerImpresora("RecibosA")
ComboRecB = TraerImpresora("RecibosB")
ComboRecC = TraerImpresora("RecibosC")
ComboRecX = TraerImpresora("RecibosX")

ComboRem = TraerImpresora("Remitos")


End Sub

Private Function TraerImpresora(strTipoFact As Variant) As String
Dim strSql As String
strSql = "Select * FROM IMPRESORA WHERE DESCRIPCION LIKE " & "'" & strTipoFact & "'"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
TraerImpresora = IIf(IsNull(rs!ruta), "", rs!ruta)
End If
rs.Close
Set rs = Nothing
End Function

Private Sub TraerImpresoras()
    For i = 1 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is ComboBox Then
                RellenoCombo Me.Controls(i)
            End If
    Next i
End Sub
Private Function RellenoCombo(combo As ComboBox)

strTmp = combo.Text
combo.Clear

For i = 0 To Printers.Count - 1
combo.AddItem Printers(i).DeviceName
Next

combo.Text = strTmp

End Function

Private Function TraerUltimaFactura(tabla, id) As Long
Dim strSql As String
strSql = "Select * From " & tabla
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then
rs.MoveLast
id = rs!numfact
Else
id = 1
End If
rs.Close
TraerUltimaFactura = Val(id)
End Function

Private Function TraerUltimoRecibo(tabla As String, txtNumrec As TextBox) As Long
Dim strSql As String
strSql = "Select * From " & tabla & " ORDER BY NUMRECIBO ASC"
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then

While Not rs.EOF
If rs!NumRecibo > 0 Then
txtNumrec = rs!NumRecibo
End If
rs.MoveNext
Wend
End If
rs.Close

TraerUltimoRecibo = Val(txtNumrec)
End Function

Private Function TraerUltimoRemito(tabla As String, txtNumRem As TextBox) As Long
Dim strSql As String
strSql = "Select * From " & tabla
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then
rs.MoveLast
id = rs!numfact
Else
id = 1
End If
rs.Close
TraerUltimoRemito = Val(id)
End Function

Private Sub GuardarTodo()
AbrirBase
GuardarUltimaFactura "FacturasA", Val(txtFactA)
GuardarUltimaFactura "FacturasB", Val(txtFactB)
GuardarUltimaFactura "FacturasC", Val(txtFactC)
GuardarUltimaFactura "FacturasX", Val(txtFactX)
GuardarUltimoRecibo "RecibosA", Val(txtRecA)
GuardarUltimoRecibo "RecibosB", Val(txtRecB)
GuardarUltimoRecibo "RecibosC", Val(txtRecC)
GuardarUltimoRecibo "RecibosX", Val(txtRecX)

GuardarUltimoRemito "Remitos", Val(txtRem)

GuardarImpresoras "FacturasA", ComboFactA
GuardarImpresoras "FacturasB", ComboFactB
GuardarImpresoras "FacturasC", ComboFactC
GuardarImpresoras "FacturasX", ComboFactX
GuardarImpresoras "RecibosA", ComboRecA
GuardarImpresoras "RecibosB", ComboRecB
GuardarImpresoras "RecibosC", ComboRecC
GuardarImpresoras "RecibosX", ComboRecX

GuardarImpresoras "Remitos", ComboRem
CerrarBase
'MsgBox "Parametros Actualizados", vbInformation
End Sub

Private Sub GuardarImpresoras(strTipoFact As Variant, valor As Variant)
Dim strSql As String
Dim rsPri As New Recordset
strSql = "SELECT * FROM IMPRESORA WHERE DESCRIPCION LIKE " & "'" & strTipoFact & "'"
rsPri.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsPri.EOF Then
rsPri!ruta = valor
rsPri.Update
End If
rsPri.Close
End Sub

Private Sub GuardarUltimaFactura(tabla, id)
If IsNumeric(id) And id <> "" Then

Dim strSql2 As String
strSql2 = "Select * From " & tabla & " Where NumFact>" & Val(id)
Dim rs2 As New Recordset
rs2.Open strSql2, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rs2.EOF Then
While Not rs2.EOF
rs2.Delete
rs2.MoveNext
Wend
End If
rs2.Close

Dim strSql As String
strSql = "Select * From " & tabla & " Where NumFact=" & Val(id)
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then
'MsgBox "La tabla " & Tabla & " No permite dupicados", vbExclamation
Else
rs.AddNew
rs!numfact = Val(id)
'rs!Impresa = True
'rs!Pagada = True
rs.Update
End If
rs.Close
End If

End Sub



Private Sub GuardarUltimoRecibo(tabla, id)
If IsNumeric(id) And id <> "" Then

Dim strSql As String
'strSql = "Select * From " & Tabla & " Where NumRecibo=" & Val(id)

strSql = "Select * From " & tabla ' & " Where NumRecibo=" & Val(id)

Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
'If Not rs.EOF Then
'MsgBox "La tabla " & Tabla & " No permite dupicados", vbExclamation
'Else
x = rs.RecordCount + 1
rs.AddNew
rs!id = x
rs!NumRecibo = Val(id)
rs.Update
'End If
rs.Close
End If

End Sub




Private Sub GuardarUltimoRemito(tabla, id)
If IsNumeric(id) And id <> "" Then

Dim strSql2 As String
strSql2 = "Select * From " & tabla & " Where NumFact>" & Val(id)
Dim rs2 As New Recordset
rs2.Open strSql2, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not rs2.EOF Then
While Not rs2.EOF
rs2.Delete
rs2.MoveNext
Wend
End If
rs2.Close

Dim strSql As String
strSql = "Select * From " & tabla & " Where NumFact=" & Val(id)
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then
'MsgBox "La tabla " & Tabla & " No permite dupicados", vbExclamation
Else
rs.AddNew
rs!numfact = Val(id)
'rs!Impresa = True
'rs!Pagada = True
rs.Update
End If
rs.Close
End If

End Sub

