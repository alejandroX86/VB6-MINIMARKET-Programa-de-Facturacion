VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBarcode 
   BackColor       =   &H00000000&
   Caption         =   "Codigo de Barras"
   ClientHeight    =   7920
   ClientLeft      =   960
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "frmBarcode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
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
      Height          =   315
      Left            =   240
      ScaleHeight     =   285
      ScaleWidth      =   1665
      TabIndex        =   15
      Top             =   7440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   7440
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   8325
      TabIndex        =   13
      Top             =   7440
      Width           =   1545
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Frame2"
      Height          =   840
      Left            =   45
      TabIndex        =   6
      Top             =   3480
      Width           =   9825
      Begin VB.TextBox txtCantidad 
         Height          =   330
         Left            =   5535
         TabIndex        =   11
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Height          =   330
         Left            =   7200
         TabIndex        =   9
         Top             =   360
         Width           =   1320
      End
      Begin VB.TextBox txtCodarticulo 
         Height          =   330
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   2310
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FFFF&
         Caption         =   "Cantidad:"
         Height          =   240
         Left            =   4770
         TabIndex        =   10
         Top             =   405
         Width           =   1365
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FFFF&
         Caption         =   "Codigo:"
         Height          =   240
         Left            =   270
         TabIndex        =   7
         Top             =   405
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Buscar Articulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1035
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   9810
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox comboCriterio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmBarcode.frx":57E2
         Left            =   3540
         List            =   "frmBarcode.frx":57EF
         TabIndex        =   3
         Text            =   "Descripcion"
         Top             =   360
         Width           =   2115
      End
      Begin VB.TextBox txtBuscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         MaxLength       =   30
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Buscar por:"
         Height          =   195
         Index           =   12
         Left            =   2700
         TabIndex        =   5
         Top             =   420
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmBarcode.frx":580F
      Height          =   2385
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   4207
      _Version        =   393216
      BackColor       =   12640511
      ForeColor       =   8388608
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Bindings        =   "frmBarcode.frx":5825
      Height          =   2895
      Left            =   45
      TabIndex        =   12
      Top             =   4440
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   12648384
      ForeColor       =   8388608
      Rows            =   10
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
ImprimirBarcodes
End Sub


Private Sub ImprimirBarcodes()


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
rs.Close
Set rs = Nothing

Picture1.ScaleMode = vbCentimeters
Printer.Copies = 1
Printer.ScaleMode = vbCentimeters
Printer.CurrentX = Xpos
Printer.CurrentY = Ypos



strSql = "SELECT * FROM BARCODES"
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic
While Not rs.EOF




Picture1.Cls
Picture1.ScaleMode = 3
Picture1.Height = Picture1.Height * (2.4 * 40 / Picture1.ScaleHeight)
Picture1.FontSize = 8
Call DrawBarcode(rs!id, Picture1, rs!Marca & " - " & rs!Talle)



For i = 1 To rs!Cantidad


If Printer.CurrentX <= AnchoHoja Then
Printer.PaintPicture Picture1, Printer.CurrentX, Printer.CurrentY
Printer.CurrentX = Printer.CurrentX + HSpc
Else

    Printer.CurrentX = Xpos
    Printer.CurrentY = Printer.CurrentY + VSpc
    Printer.PaintPicture Picture1, Printer.CurrentX, Printer.CurrentY
    Printer.CurrentX = Printer.CurrentX + HSpc


    If Printer.CurrentY >= AltoHoja And Printer.CurrentX >= AnchoHoja Then
    Printer.NewPage
    Printer.CurrentX = Xpos
    Printer.CurrentY = Ypos
    End If
    
    
End If


If Printer.CurrentY >= AltoHoja And Printer.CurrentX >= AnchoHoja Then
Printer.NewPage
Printer.CurrentX = Xpos
Printer.CurrentY = Ypos
End If



Next



rs.MoveNext
Wend

CerrarBase
Printer.EndDoc











End Sub




Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub


Private Sub Form_Load()
MostrarTodos
End Sub

Private Sub LimpioBarcode()


strSql = "DELETE FROM BARCODES"

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic


End Sub

Private Sub LimpioFlex()

titulos = " Codigo.|Descripción|Color|talle|Precio"

 MSHFlexGrid1.Clear
    With MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With

titulos2 = " Codigo.|Descripción|Color|talle|Cantidad"
    
    With MSHFlexGrid2
        .Clear
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos2
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With

End Sub
Private Sub MostrarTodos()
LimpioFlex
AbrirBase
Dim strSql As String
strSql = "SELECT * FROM Articulos"
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    Do While Not rs.EOF
        i = i + 1
        linea = rs!id _
                & Chr(9) & rs!Descripcion _
                & Chr(9) & rs!Marca _
                & Chr(9) & rs!Talle _
                & Chr(9) & Format(rs!Precio, "Fixed")

    MSHFlexGrid1.AddItem linea, i
    rs.MoveNext
    Loop
    lblTotalRef = rs.RecordCount

LimpioBarcode


CerrarBase


AutoFlex Me.MSHFlexGrid1

End Sub


'######################## MENUES #############################33

Private Sub mnuSalir_Click()
Unload Me
End Sub

Private Sub MSHFlexGrid1_Click()


MSHFlexGrid1.Col = 0
txtCodarticulo = MSHFlexGrid1.Text


txtCantidad = ""
txtCantidad.SetFocus



End Sub
Private Sub Command1_Click()
If IsNumeric(txtCantidad) Then
AceptarAgregar txtCantidad.Text
End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(txtCantidad) Then
        KeyAscii = 0
        AceptarAgregar txtCantidad.Text
    End If
End Sub

Private Sub AceptarAgregar(ByVal txtCantidad As Variant)
AbrirBase
strSql = "Select * From Articulos Where ID=" & txtCodarticulo
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then

Dim rs1 As New Recordset

strSql = "Select * From Barcodes"
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic
rs1.AddNew

rs1!id = rs!id
rs1!Descripcion = rs!Descripcion
rs1!Marca = rs!Marca
rs1!Talle = rs!Talle
rs1!Precio = rs!Precio
rs1!Cantidad = txtCantidad
rs1.Update
rs1.Close
Set rs1 = Nothing




End If
rs.Close
MostrarBarcodes
CerrarBase

End Sub


Private Sub MostrarBarcodes()



titulos2 = " Codigo.|Descripción|Color|talle|Precio|Cantidad"
    
    With MSHFlexGrid2
        .Clear
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos2
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With
    
    
    
Dim strSql As String
strSql = "SELECT * FROM Barcodes"
Dim rs As New Recordset
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

    Do While Not rs.EOF
        i = i + 1
                
                linea = rs!id _
                & Chr(9) & rs!Descripcion _
                & Chr(9) & rs!Marca _
                & Chr(9) & rs!Talle _
                & Chr(9) & Format(rs!Precio, "Fixed") _
                & Chr(9) & rs!Cantidad

    MSHFlexGrid2.AddItem linea, i
    rs.MoveNext
    Loop
    lblTotalRef = rs.RecordCount

AutoFlex Me.MSHFlexGrid2




End Sub


'################PROCEDIMIENTO BUSCAR ####################

Private Sub txtBuscar_Change()
AceptarBuscar txtBuscar.Text
End Sub



Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        AceptarBuscar txtBuscar.Text
    End If
End Sub
Private Sub AceptarBuscar(ByVal txtBuscar As Variant)
LimpioFlex
AbrirBase
Dim rs As New Recordset
Dim strSql As String
strSql = "Select * From Articulos"

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

Select Case comboCriterio.Text
Case "Codigo"
    If txtBuscar <> "" Then
    rs.Filter = "ID =" & Val(txtBuscar)
    End If
Case "Descripcion"
    If txtBuscar <> "" Then
    rs.Filter = "Descripcion LIKE '*" + txtBuscar + "*'"
    End If
Case "Color"
    If txtBuscar <> "" Then
    rs.Filter = "Marca LIKE '*" + txtBuscar + "*'"
    End If

Case Else
rs.Filter = ""
End Select

    Do While Not rs.EOF
        i = i + 1
        linea = rs!id _
                & Chr(9) & rs!Descripcion _
                & Chr(9) & rs!Marca _
                & Chr(9) & rs!Talle _
                & Chr(9) & Format(rs!Precio, "Fixed")

    MSHFlexGrid1.AddItem linea, i
    rs.MoveNext
    Loop
    lblTotalRef = rs.RecordCount
AutoFlex Me.MSHFlexGrid1

CerrarBase
End Sub

Private Sub mnuCerrar_Click()
Unload Me
End Sub


Private Sub cmdBuscar_Click()
AceptarBuscar txtBuscar.Text
End Sub
Private Sub mnuImprimir_Click()
    If MsgBox("¿Imprimir Listado?", vbYesNo + vbInformation, "Impresión") = vbYes Then
    'IniciarImpresion
    End If
End Sub

Private Sub MSHFlexGrid2_DblClick()

MSHFlexGrid2.Col = 0
txtCodarticulo = MSHFlexGrid2.Text

If MsgBox("Atención: Eliminar Este Item???", vbYesNoCancel + vbExclamation) = vbYes Then

AbrirBase

strSql = "DELETE * FROM BARCODES WHERE ID =" & txtCodarticulo
Dim rs1 As New Recordset
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic



MostrarBarcodes



CerrarBase



End If



End Sub




