VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmClonarListaPrecios 
   BackColor       =   &H00000000&
   Caption         =   "Clonar Lista de Precios"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   Icon            =   "frmClonarListaPrecios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Clonar A:"
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Top             =   6360
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "CLONAR"
         Height          =   255
         Left            =   5460
         TabIndex        =   5
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   180
      Width           =   3915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmClonarListaPrecios.frx":57E2
      Height          =   5715
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   10081
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   16777152
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblID 
      Caption         =   "lblID"
      Height          =   195
      Left            =   4800
      TabIndex        =   6
      Top             =   780
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Seleccionar Lista:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmClonarListaPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CLONAR
End Sub


Private Sub Form_Load()
'dtp1 = Date
'dtp2 = Date
RELLENARCOMBO
'VerUtilidad
'lblCosto = ""
'lblGanancia = ""
VerListaDePrecios
End Sub


Private Sub RELLENARCOMBO()

AbrirBase
c1 = "Select * From Clientes"
Dim rs1 As New Recordset
rs1.Open c1, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not (rs1.BOF And rs1.EOF) Then
While Not rs1.EOF
Combo1.AddItem rs1!nombre
rs1.MoveNext
Wend
rs1.MoveFirst
Combo1.Text = rs1!nombre
lblID = rs1!id
End If
CerrarBase
End Sub
Private Sub Combo1_Click()
AbrirBase
VerCodigo
CerrarBase
VerListaDePrecios
End Sub
Private Sub VerCodigo()
c2 = "Select * From Clientes Where Nombre Like " & "'" & Combo1.Text & "'"

Dim rs2 As New Recordset
rs2.Open c2, DB, adOpenKeyset, adLockOptimistic, adCmdText

If Not (rs2.BOF And rs2.EOF) Then
lblID = rs2!id
Else
lblID = "0"
End If
End Sub

Private Sub VerListaDePrecios()

titulos = " COD.|DESCRIPCION|MARCA|PRECIO"

 MSHFlexGrid1.Clear
    With MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .ColWidth(0) = 800
        .ColWidth(1) = 2500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With

strSql = "SELECT * FROM PreciosClientes WHERE CodProveedor=" & Val(lblID) & " ORDER BY CodArticulo ASC"



AbrirBase
rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

TotVent = "0"

While Not rs.EOF
i = i + 1
                linea = Format(rs!codarticulo, "0000") _
                & Chr(9) & rs!Descripcion _
                & Chr(9) & rs!Marca _
                & Chr(9) & Format(rs!Precio, "fixed") _

MSHFlexGrid1.AddItem linea, i
rs.MoveNext
Wend
CerrarBase
End Sub


Private Sub CLONAR()

Dim rsClonOrig As New Recordset
AbrirBase

strSql = "SELECT * FROM PRECIOSCLIENTES WHERE CodProveedor=" & Val(lblID) & " ORDER BY CODARTICULO ASC"


rsClonOrig.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
If Not rsClonOrig.EOF Then
While Not rsClonOrig.EOF

        Dim rsClonDest As New Recordset
        vx = "SELECT * FROM PRECIOSCLIENTES WHERE CodProveedor=" & Val(Label2) & " AND CODARTICULO=" & Val(rsClonOrig!codarticulo)
        rsClonDest.Open vx, DB, adOpenKeyset, adLockOptimistic, adCmdText
        ' si existe que lo actualize
        If Not rsClonDest.EOF Then
        rsClonDest!Precio = rsClonOrig!Precio
        'rsClonDest!PrecioDocena = rsClonOrig!PrecioDocena
        rsClonDest.Update
        Else
        rsClonDest.AddNew
        rsClonDest!CodProveedor = Val(Label2)
        rsClonDest!codarticulo = rsClonOrig!codarticulo
        rsClonDest!Descripcion = rsClonOrig!Descripcion
        rsClonDest!Marca = rsClonOrig!Marca
        
        rsClonDest!Precio = rsClonOrig!Precio
        'rsClonDest!PrecioDocena = rsClonOrig!PrecioDocena
        rsClonDest.Update
        End If
rsClonDest.Close
rsClonOrig.MoveNext
Wend
MsgBox "Lista Clonada!", vbInformation
End If

CerrarBase

Unload frmPreciosClientes
frmPreciosClientes.Show
Unload Me


End Sub




