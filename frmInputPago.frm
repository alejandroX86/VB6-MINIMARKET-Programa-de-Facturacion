VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInputPago 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imputar Pago"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   Icon            =   "frmInputPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   315
      Left            =   2880
      TabIndex        =   12
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   315
      Left            =   600
      TabIndex        =   11
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle de Pago:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   4875
      Begin VB.Label Label7 
         Caption         =   "Total:"
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Fact:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "NumFact:"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   420
         TabIndex        =   14
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cod:"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSaldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   8
         Top             =   2100
         Width           =   3615
      End
      Begin VB.Label lblTotal 
         Caption         =   "lblTotal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label lblFecha 
         Caption         =   "lblFecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label lblNumFact 
         Caption         =   "lblNumFact"
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
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label lblNombre 
         Caption         =   "lblNombre"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   660
         Width           =   3615
      End
      Begin VB.Label lblID 
         Caption         =   "lblID"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   795
      End
   End
   Begin MSComCtl2.DTPicker dtp1 
      Height          =   375
      Left            =   1860
      TabIndex        =   1
      Top             =   2700
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   141164545
      CurrentDate     =   38421
   End
   Begin VB.TextBox txtImporte 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1860
      MaxLength       =   12
      TabIndex        =   0
      Text            =   "txtImporte"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Importe Pago:"
      Height          =   255
      Left            =   780
      TabIndex        =   10
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Fecha Pago:"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   2760
      Width           =   1035
   End
End
Attribute VB_Name = "frmInputPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
AbrirBase
GuardarPago
VerCuentaCorriente
CerrarBase
Unload Me
End Sub

Private Sub Command2_Click()

Unload Me



End Sub


Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(txtImporte) Then
        KeyAscii = 0
        Command1_Click
    End If
End Sub
Private Sub GuardarPago()


If IsNumeric(txtImporte) Then

Dim rsCta As New Recordset
strSql = "Select * From CuentasCorrientes"
rsCta.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
Dim nx As Long

If rsCta.EOF Then
rsCta.AddNew
rsCta!id = "1"
rsCta!CodCliente = Val(lblID)
rsCta!numfact = Val(lblNumFact)
rsCta!Fecha = dtp1
rsCta!Descripcion = "Pago Factura Nº " & Val(lblNumFact)
rsCta!Haber = Val(txtImporte)
rsCta.Update
Else
rsCta.MoveLast
nx = Val(rsCta!id)
rsCta.AddNew
rsCta!id = nx + 1
rsCta!CodCliente = Val(lblID)
rsCta!numfact = Val(lblNumFact)
rsCta!Fecha = dtp1
rsCta!Descripcion = "Pago Factura Nº " & Val(lblNumFact)
rsCta!Haber = Val(txtImporte)
rsCta.Update
End If
'rsCta.MoveFirst
'Dim totD, totH As Double
'While Not rsCta.EOF
'totD = Val(totD) + rsCta!Debe
'totH = Val(totH) + rsCta!Haber
'rsCta.MoveNext
'Wend
'rsCta.MoveLast
'rsCta!Saldo = Val(totD) - Val(totH)
'rsCta.Update
MsgBox "Cuenta Corriente Actualizada", vbInformation


End If


End Sub

Private Sub VerCuentaCorriente()


titulos = "ID|Nº FACT.|FECHA|DESCRIPCIÓN|DEBE|HABER"

 frmClientes.MSHFlexGrid1.Clear
    With frmClientes.MSHFlexGrid1
        .Rows = 2
        .FixedRows = 1
        .Col = 0
        .ColSel = 0
        .Row = 1
        .RowSel = 1
        .FormatString = titulos
        .ColWidth(0) = 0
        .ColWidth(1) = 1000
        .ColWidth(2) = 1200
        .ColWidth(3) = 2500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .Visible = True
        .AllowBigSelection = True
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
    End With


Dim rsCons As New Recordset
strSql = "Select * From CuentasCorrientes Where Codcliente=" & Val(frmClientes.txtCodDist)
rsCons.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText

TotDebe = "0"
TotHaber = "0"
While Not rsCons.EOF
                i = i + 1
                
                linea = rsCons!id _
                & Chr(9) & rsCons!numfact _
                & Chr(9) & rsCons!Fecha _
                & Chr(9) & rsCons!Descripcion _
                & Chr(9) & Format(rsCons!Debe, "fixed") _
                & Chr(9) & Format(rsCons!Haber, "fixed")

frmClientes.MSHFlexGrid1.AddItem linea, i
TotDebe = TotDebe + rsCons!Debe
TotHaber = TotHaber + rsCons!Haber
rsCons.MoveNext
Wend
frmClientes.lblTotalRef = rsCons.RecordCount
frmClientes.lblTotal = "Saldo Cta/Cte: $ " & Format(TotDebe - TotHaber, "#,###.#0")

End Sub




