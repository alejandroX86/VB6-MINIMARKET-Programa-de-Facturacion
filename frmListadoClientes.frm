VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmListadoClientes 
   BackColor       =   &H00000000&
   Caption         =   "Imprimir Listado de Clientes"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "frmListadoClientes.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3210
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Seleccionar Tipo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   2220
      Width           =   4215
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox ckDeudores 
         Caption         =   "Deudores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CheckBox ckHabilitados 
         Caption         =   "Activos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Orden del Listado:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1515
      Left            =   5940
      TabIndex        =   13
      Top             =   660
      Width           =   2295
      Begin VB.OptionButton OptAlf 
         Caption         =   "Alfabético"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   900
         Width           =   1635
      End
      Begin VB.OptionButton OptNro 
         Caption         =   "Numérico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Value           =   -1  'True
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   8115
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Min             =   1e-4
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1740
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Procesando Datos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5940
      TabIndex        =   9
      Top             =   2340
      Width           =   2295
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5940
      TabIndex        =   8
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   2
      Top             =   2340
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1515
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   5715
      Begin VB.CheckBox ckTodos 
         Caption         =   "Imprimir todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox TbNomHasta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2700
         TabIndex        =   7
         Top             =   900
         Width           =   2835
      End
      Begin VB.TextBox TbNomDesde 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2700
         TabIndex        =   6
         Top             =   480
         Width           =   2835
      End
      Begin VB.TextBox TbDesde 
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
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TbHasta 
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
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   1
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label lbldesde 
         Caption         =   "Desde Cliente :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta Cliente  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListadoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstSocios As ADODB.Recordset
Dim RstCuotas As ADODB.Recordset
Dim i As Integer
Dim Titulo As String
Dim Hoy As Date
Dim Widthchr As Currency
Dim strSql As String
Dim QueCob As Integer
Dim OKCuotas As Boolean
Dim OKAnt As Boolean
Dim StrRstCuotas As String
Dim CodCob As Integer
Dim ContLin As Integer
Dim ContHoj As Integer
Dim ContRay As Integer
Dim letra As String
Dim Raya As String
Dim Año As Integer
Dim Mes As Integer
'Estas son para imprimir desde un numero de socio a otro
Dim DesdeSoc As Long
Dim HastaSoc As Long
Dim DeudaSocial As Currency
Dim donde(11) As Integer
Dim TotalMensual(11) As Currency


Private Sub ckTodos_Click()

If ckTodos.Value = 1 Then

TbDesde = ""
TbHasta = ""
TbNomDesde = ""
TbNomHasta = ""
TbDesde.Enabled = False
TbHasta.Enabled = False
cmdImprimir.Visible = True
cmdImprimir.SetFocus
Else
DeshabilitarEdicion
End If



    
    

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = 8500
Me.Height = 3650
    'AbroBaseDatos
    OKCuotas = False
    Hoy = Date
    Raya = String(157, "-")
    
    'Estas son las coordenadas para las x en los meses del la planilla
    donde(0) = 141.7
    donde(1) = 143.7
    donde(2) = 145.7
    donde(3) = 147.7
    donde(4) = 149.7
    donde(5) = 151.7
    donde(6) = 153.7
    donde(7) = 155.7
    donde(8) = 157.7
    donde(9) = 159.7
    donde(10) = 161.7
    donde(11) = 163.7
    
    For i = 0 To 11
        TotalMensual(i) = 0
    Next i
    
    'DeshabilitarEdicion
    
    
    
    
End Sub



Private Sub cmdCancelar_Click()
DeshabilitarEdicion
End Sub


Private Sub DeshabilitarEdicion()
    
  
    cmdImprimir.Visible = False
    
    TbDesde = ""
    TbHasta = ""
    TbNomDesde = ""
    TbNomHasta = ""
    
    TbDesde.Enabled = True
    TbHasta.Enabled = True
    TbDesde.SetFocus
   
   
   
End Sub





Private Sub cmdSalir_Click()
Unload Me
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub




' ################ ACEPTAR 1#############################

Private Sub TbDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(TbDesde) Then
        KeyAscii = 0
        Aceptar1 TbDesde.Text
    End If
End Sub


Private Sub Aceptar1(ByVal TbDesde As Variant)
'On Error GoTo Error_Guardar

AbrirBase
'DB.BeginTrans
AceptarRegistro1
'Error_Guardar:
'    If Err.Number <> 0 Then
'    DB.RollbackTrans
'    CerrarBase
'    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
'    Else
'    DB.CommitTrans
    CerrarBase
    'HabilitarEdicion
'    End If

End Sub


Private Sub AceptarRegistro1()
Dim strSql As String
strSql = "SELECT * FROM Clientes WHERE ID=" & Val(TbDesde)

rs.Open strSql, DB, adOpenKeyset, adLockOptimistic, adCmdText
           
  ' Si existe
If Not (rs.BOF And rs.EOF) Then
          
TbNomDesde = rs!nombre
TbHasta.SetFocus
TbDesde.Enabled = False
Else
MsgBox "No Existe ese Nº", vbCritical
TbDesde.SetFocus
End If

End Sub




' ################ ACEPTAR 2#############################

Private Sub TbHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(TbHasta) Then
        KeyAscii = 0
        Aceptar2 TbHasta.Text
    End If
End Sub


Private Sub Aceptar2(ByVal TbHasta As Variant)
'On Error GoTo Error_Guardar

AbrirBase
'DB.BeginTrans
AceptarRegistro2
'Error_Guardar:
'    If Err.Number <> 0 Then
'    DB.RollbackTrans
'    CerrarBase
'    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
'    Else
'    DB.CommitTrans
    CerrarBase
    'HabilitarEdicion
'    End If
End Sub


Private Sub AceptarRegistro2()
Dim strSql As String
strSql = "SELECT * FROM Clientes WHERE ID=" & Val(TbHasta)

rs.Open (strSql), DB, adOpenKeyset, adLockOptimistic, adCmdText
           
  ' Si existe
If Not (rs.BOF And rs.EOF) Then
          
TbNomHasta = rs!nombre
cmdImprimir.Visible = True
cmdImprimir.SetFocus
TbHasta.Enabled = False
Else
MsgBox "No Existe ese Nº", vbCritical
TbHasta.SetFocus
End If

End Sub






Private Sub cmdImprimir_Click()
    'On Error GoTo eror
    
    AbrirBase
    ImprimirSeleccion
    CerrarBase
    
Frame3.Visible = False

'eror:
'Resume Next

End Sub

Private Sub ImprimirSeleccion()

    
On Error Resume Next
    cd1.ShowPrinter
    
    DesdeSoc = Val(TbDesde.Text)
    HastaSoc = Val(TbHasta.Text)
    
    Frame3.Visible = True
    
    Set RstSocios = New ADODB.Recordset
    RstSocios.CursorLocation = adUseClient
    
    
    If ckTodos.Value = 1 Then
    strSql = "SELECT * FROM Clientes"
    Titulo = "Listado de Clientes - "
            If ckHabilitados = 1 Then
            strSql = strSql & " Where Activo=" & 1
            Titulo = "Listado de Clientes Activos - "

                If ckDeudores = 1 Then
                strSql = strSql & " AND Saldo>" & 0
                Titulo = "Listado de Clientes Activos Deudores - "
                End If
            
            Else
            
                If ckDeudores = 1 Then
                strSql = strSql & " WHERE Saldo>" & 0
                Titulo = "Listado de Clientes Deudores - "
                End If
            
            End If
            
            

    Else
    strSql = "select * from Clientes where ID between " & Str(DesdeSoc) & " And " & Str(HastaSoc)
    Titulo = "Listado de Clientes Desde " & TbDesde & " Hasta " & TbHasta & " - "
            If ckHabilitados = 1 Then
            strSql = strSql & " AND Activo=" & 1
    Titulo = "Listado de Clientes Desde " & TbDesde & " Hasta " & TbHasta & " - "
            End If
            
            If ckDeudores = 1 Then
            strSql = strSql & " AND saldo>" & 0
    Titulo = "Listado de Clientes Desde " & TbDesde & " Hasta " & TbHasta & " - "
            End If
    
    End If
    
    
    
    If OptAlf.Value = True Then
        strSql = strSql & " Order by Nombre asc"
    Else
        strSql = strSql & " Order by ID"
    End If



    RstSocios.Open strSql, DB, adOpenDynamic, adLockOptimistic
    Label1.Caption = RstSocios.RecordCount
    ProgressBar1.Max = RstSocios.RecordCount
    ProgressBar1.Min = 0
    
    ImpTitulos
    
    ContRay = 0
    ContLin = 0
    espacioCelda = 0.4
    
    
    Do While Not RstSocios.EOF
            
       ProgressBar1.Value = RstSocios.AbsolutePosition
       Label1.Caption = RstSocios.AbsolutePosition
       
If ContRay > 0 Then
Printer.CurrentY = Printer.CurrentY + espacioCelda
End If
                     
Dim CadenaCod As String
Dim LongitudCod As Long
CadenaCod = Format(CadenaCod, "0###")
LongitudCod = Len(CadenaCod)

Printer.CurrentX = 2 + (Val(LongitudCod)) - (Printer.TextWidth(Format(RstSocios!id, "0###")))
Printer.Print Format(RstSocios!id, "0###");


       
       Printer.CurrentX = 2.3
       Printer.Print Left(RstSocios!nombre, 22);
       
       
       
       
    Printer.CurrentX = 6.5
       Printer.Print Left(RstSocios!domicilio, 22);

       
       Printer.CurrentX = 11.5
       Printer.Print Left(RstSocios!telefono, 22);

       Printer.CurrentX = 13.7
       Printer.Print Left(RstSocios!Localidad, 22);


'DEUDA
If Check1.Value = 1 Then
Dim rsDeuda As New Recordset
strSql = "Select * From CuentasCorrientes Where CodCliente=" & Val(RstSocios!id)
rsDeuda.Open strSql, DB, adOpenDynamic, adLockOptimistic, adCmdText

TotDebe = "0"
TotHaber = "0"
While Not rsDeuda.EOF
TotDebe = TotDebe + rsDeuda!Debe
TotHaber = TotHaber + rsDeuda!Haber
rsDeuda.MoveNext
Wend
Printer.CurrentX = 19 - Printer.TextWidth(Format(TotDebe - TotHaber, "#,###.#0"))
Printer.Print Format(TotDebe - TotHaber, "#,###.#0");
rsDeuda.Close
End If
If ckDeudores.Value = 1 Then

Dim Cadena2 As String
Dim Longitud2 As Long
Cadena2 = Format(Cadena2, "#,###.#0")
Longitud2 = Len(Cadena2)

Printer.CurrentX = 20 + (Val(Longitud2)) - (Printer.TextWidth(Format(RstSocios!Saldo, "#,###.#0")))
Printer.Print Format(RstSocios!Saldo, "#,###.#0");

End If

       ContRay = ContRay + 1
       ContLin = ContLin + 1
       
       If Printer.CurrentY > 28 Then
          Printer.NewPage
          ImpTitulos
          ContLin = 0
          ContRay = 0
       End If
       
       If ContRay = 5 Then
       Printer.DrawStyle = 2
           Printer.CurrentY = Printer.CurrentY + espacioCelda
           Printer.Line (1, Printer.CurrentY)-(20, Printer.CurrentY)
           ContRay = 0
       End If
       RstSocios.MoveNext
   Loop
    
    Printer.EndDoc
    'cmdsalir_click

End Sub



Private Sub ImpTitulos()


    'Imprime el Template de la hoja
Printer.ScaleMode = vbCentimeters
Printer.PaperSize = 9
Printer.Orientation = 1
Printer.FontName = "ARIAL"
Printer.DrawStyle = 0
Printer.Font.Size = 14

Printer.PaintPicture frmLogo.Picture1.Picture, 1, 1
Printer.Font.Size = 8

Printer.CurrentX = 18.8
Printer.CurrentY = 1
Printer.Print "Página " & Printer.Page


Printer.CurrentX = 18.5
Printer.CurrentY = 1.5
Printer.Print Date '& " " & Format(Time, "hh:nn am/pm")

Printer.Font.Size = 14

x = 21 / 2
Printer.CurrentX = x - (Printer.TextWidth(Titulo)) + (Printer.TextWidth(Titulo)) / 2
Printer.CurrentY = 1.5
Printer.Print Titulo

Printer.Font.Size = 8


' TITULO DE LOS CAMPOS

Printer.CurrentY = 4


Printer.CurrentX = 1
Printer.Print "Codigo";

Printer.CurrentX = 2.3
Printer.Print "Nombre";

Printer.CurrentX = 6.5
Printer.Print "Dirección";

Printer.CurrentX = 11.5
Printer.Print "Teléfono";

Printer.CurrentX = 13.7
Printer.Print "Localidad";

Printer.CurrentX = 18
Printer.Print "SALDO"
'
'Printer.CurrentX = 19.2
'Printer.Print "Saldo"
    
Printer.Line (1, (Printer.CurrentY + espacioCelda))-(20, (Printer.CurrentY + espacioCelda))

End Sub







