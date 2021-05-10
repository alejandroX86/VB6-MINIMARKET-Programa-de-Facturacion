Attribute VB_Name = "ModFlexFunctions"
Public Sub RefreshGrid(XFlex As MSHFlexGrid, NombreTabla As String)
AbrirBase
strSql = "SELECT * FROM [" & NombreTabla & "]"
Dim rs1 As New Recordset
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic

Dim titulos As String
titulos = ""
Dim z As Integer
For z = 0 To rs1.Fields.Count - 1
titulos = titulos & Trim(rs1.Fields.Item(z).Name) & "|"
Next z

titulos = Left(titulos, Len(titulos) - 1)

With XFlex
.Clear
.Cols = 2
.Rows = 2
.FixedRows = 1
.Col = 0
.ColSel = 0
.Row = 0
.RowSel = 0
.FormatString = titulos
.Visible = True
.AllowBigSelection = True
.AllowUserResizing = flexResizeColumns
.SelectionMode = flexSelectionByRow
.ColAlignmentFixed = 3
End With

If Not rs1.EOF Then
rs1.MoveFirst
i = 0
While Not rs1.EOF
i = i + 1
Fila = ""



'            If InStr(1, valor, ",", 1) > 0 Then
'            valor = Format(CDbl(valor), "standard")
'            End If



    For G = 0 To rs1.Fields.Count - 1
        
               
        If rs1.Fields.Item(G).Type = adCurrency Then
        Fila = Fila & Format(rs1.Fields.Item(G).Value, "standard") & Chr(9)
        ElseIf rs1.Fields.Item(G).Type = adInteger Then
        Fila = Fila & Format(rs1.Fields.Item(G).Value, "0000") & Chr(9)
        Else
        Fila = Fila & rs1.Fields.Item(G).Value & Chr(9)
        End If
    
    
    Next G

XFlex.AddItem Fila, i
rs1.MoveNext
Wend
End If

rs1.Close

AutoFlex XFlex
FlexRayado XFlex, &HFFFFFF, &H8000000F
CerrarBase
End Sub

'##########################################################
'###################### FORMATO DEL FLEXGRID #######################
'##########################################################
Public Sub AutoFlex(flex As MSHFlexGrid)
'On Error Resume Next
flex.Col = 0
While Not flex.Col = flex.Cols - 1

        flex.Row = 1
        While Not flex.Row = flex.Rows - 1
        If Val(flex.ColWidth(flex.Col)) < Val((Len(flex.Text) * 130)) Then
        flex.ColWidth(flex.Col) = Val((Len(flex.Text) * 130))
        End If

        flex.Row = flex.Row + 1
        Wend

flex.Col = flex.Col + 1
Wend

'Flex.Width = totalancho

End Sub



Public Sub AutoajustarFlex(flex As MSHFlexGrid, Check As Variant)

flex.Col = 0
While Not flex.Col = flex.Cols - 1
    
    If Check(flex.Col).Value = 1 Then
    
        flex.Row = 1
        While Not flex.Row = flex.Rows - 1
        
            If flex.ColWidth(flex.Col) < (Len(flex.Text) * 118) Then
            flex.ColWidth(flex.Col) = (Len(flex.Text) * 118)
            End If
        
        flex.Row = flex.Row + 1
        Wend
    
    End If

flex.Col = flex.Col + 1
Wend

End Sub

Public Function EImpar(ByVal iNum As Long) As Boolean
  EImpar = (iNum Mod 2)
End Function
Public Sub FlexRayado(flex As MSHFlexGrid, lCorPar As Long, lCorImpar As Long)
  Dim iLinha As Integer
  flex.FillStyle = flexFillRepeat
  For iLinha = 1 To flex.Rows - 1
     With flex
       .Row = iLinha
       If EImpar(iLinha) Then 'Se a linha for impar:
         'Seleciona a partir da primeira coluna
         .Col = 0
         'Seleciona até a última coluna
         .ColSel = .Cols - 1
         'Aplica a cor
         .CellBackColor = lCorImpar
       Else 'Se a linha for par:
         'Seleciona a partir da primeira coluna
         .Col = 0
         'Seleciona até a última coluna
         .ColSel = .Cols - 1
         'Aplica a cor
         .CellBackColor = lCorPar
       End If
     End With
  Next
  flex.FillStyle = flexFillSingle
End Sub




'FlexGrid_To_Excel flex01, flex01.Rows, flex01.Cols, "listado de Rubros"
'FlexGrid_To_Excel Me.MSHFlexGrid1, MSHFlexGrid1.Rows, MSHFlexGrid1.Cols, "listado de Rubros"

Public Sub FlexGrid_To_Excel(TheFlexgrid As MSHFlexGrid, _
  TheRows As Integer, TheCols As Integer, NombreArchivo As String, _
  Optional GridStyle As Integer = 1, Optional WorkSheetName _
  As String)

Dim objXL As New Excel.Application
Dim wbXL As New Excel.Workbook
Dim wsXL As New Excel.Worksheet
Dim intRow As Integer ' counter
Dim intCol As Integer ' counter

Dim valor As Variant

If Not IsObject(objXL) Then
    MsgBox "You need Microsoft Excel to use this function", _
       vbExclamation, "Print to Excel"
    Exit Sub
End If

'On Error Resume Next is necessary because
'someone may pass more rows
'or columns than the flexgrid has
'you can instead check for this,
'or rewrite the function so that
'it exports all non-fixed cells
'to Excel

On Error Resume Next

' open Excel

Set wbXL = objXL.Workbooks.Add
Set wsXL = objXL.ActiveSheet


wbXL.SaveAs App.Path + "\" & NombreArchivo


' name the worksheet
With wsXL
    If Not WorkSheetName = "" Then
        .Name = WorkSheetName
    End If
End With
    
' fill worksheet
For intRow = 1 To TheRows
    For intCol = 1 To TheCols
        With TheFlexgrid
        
            valor = .TextMatrix(intRow - 1, intCol - 1)
            
            
            If valor = "" Then valor = "0000"
            'valor = CStr(valor)
            
            
            'esto esta obviado excepticonalmente
            'valor = IIf(IsNumeric(valor), Format(Val(valor), "standard"), valor)
            'valor = IIf(IsDate(valor), CDate(valor), valor)
            
            
'            If InStr(1, valor, ".", 1) > 0 Then
'            valor = CDate(valor)
'            End If
'
'
            If InStr(1, valor, ",", 1) > 0 Then
            valor = Format(CDbl(valor), "standard")
            End If
            
            
            If IsNumeric(valor) Then
            valor = CDbl(valor)
            End If
            
            
            wsXL.Cells(intRow, intCol).Value = valor
        End With
    Next
Next

' format the look
For intCol = 1 To TheCols
    wsXL.Columns(intCol).AutoFit
    'wsXL.Columns(intCol).AutoFormat (1)
    wsXL.Range("a1", Right(wsXL.Columns(TheCols).AddressLocal, 1) & TheRows).AutoFormat GridStyle
Next


If MsgBox("Informe generado, ¿Abrir Informe?", vbYesNoCancel + vbInformation) = vbYes Then
objXL.Visible = True
'wbXL.Close False
Else
'objXL.Visible = False
wbXL.Close True
End If

Set wbXL = Nothing
Set wsXL = Nothing
Set objXL = Nothing

End Sub







Public Sub ExcelToFlexgrid(MSFlexGrid1 As MSHFlexGrid, ruta As String)
  Dim xlObject    As Excel.Application
  Dim xlWB        As Excel.Workbook
  
      Set xlObject = New Excel.Application
      Set xlWB = xlObject.Workbooks.Open("" & ruta & "") 'Open your book here
      Clipboard.Clear
      With xlObject.ActiveWorkbook.ActiveSheet
         .Range("A1:AF100").Copy 'Set selection to Copy
      End With
      
     With MSFlexGrid1
         .Redraw = False     'Dont draw until the end, so we avoid that flash
         .Row = 0            'Paste from first cell
         .Col = 0
         .RowSel = .Rows - 1 'Select maximum allowed (your selection shouldnt be greater than this)
         .ColSel = .Cols - 1
         .Clip = Replace(Clipboard.GetText, vbNewLine, vbCr) 'Replace carriage return with the correct one
         .Col = 1            'Just to remove that blue selection from Flexgrid
         .Redraw = True      'Now draw
     End With

     xlObject.DisplayAlerts = False 'To avoid "Save woorkbook" messagebox

     'Close Excel
     xlWB.Close
     xlObject.Application.Quit
     
     Set xlWB = Nothing
     Set xlObject = Nothing
End Sub

