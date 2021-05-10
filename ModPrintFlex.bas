Attribute VB_Name = "ModPrintFlex"
Public Sub ImprimirFlex(flexgrid As MSHFlexGrid, titulo As String)
Printer.PaperSize = 9
Printer.ScaleMode = vbCentimeters
ContRay = 0
ContLin = 0
espacioCelda = 0.4

'OBTENEMOS LOS ANCHOS
Dim colw(0 To 16) As Double
Dim medida(0 To 16) As Double
Dim medidaActual As Double
Dim MedidaTotal As Double
MedidaTotal = 1
For i = 0 To flexgrid.Cols - 1
flexgrid.Col = i
medida(i) = 0
colw(i) = 0

    'If Check(i).Value = 1 Then

        For CC = 0 To flexgrid.Rows - 1
        flexgrid.Row = CC
        
          medidaActual = Printer.TextWidth(flexgrid.Text)
            
            If CDbl(medida(i)) < CDbl(medidaActual) Then
            medida(i) = CDbl(medidaActual)
            End If
        
    
        Next CC
        
    colw(i) = CDbl(MedidaTotal)
    MedidaTotal = CDbl(MedidaTotal) + CDbl(medida(i)) + 0.3
    
    'End If

Next i


'If MedidaTotal > 28 Then
'
'MsgBox "Demasiados campos, No se puede imprimir el documento"
'Exit Sub
'
'Else

If MedidaTotal > 20 Then
Printer.Orientation = 2
largomaximo = 19
'MedidaTotal = 28
Else
Printer.Orientation = 1
largomaximo = 28
'MedidaTotal = 19
End If

'End If

MedidaTotal = 20

'Printer.PaintPicture LoadPicture(LeerIni("0", "LOGO", "MYLOGO")), 1, 1, 2.5, 2.5

Printer.Font.Size = 8

Printer.CurrentX = MedidaTotal - Printer.TextWidth("Página " & Printer.Page)
Printer.CurrentY = 1
Printer.Print "Página " & Printer.Page

Printer.CurrentX = MedidaTotal - Printer.TextWidth(Format(Date, "dd-mm-yyyy"))
Printer.CurrentY = 1.5
Printer.Print Format(Date, "dd-mm-yyyy")

Printer.Font.Bold = True
Printer.Font.Size = 12
x = 21 / 2
Printer.CurrentX = x - (Printer.TextWidth(titulo)) + (Printer.TextWidth(titulo)) / 2
Printer.CurrentY = 2
Printer.Print titulo
Printer.Font.Bold = False
Printer.Font.Size = 8

For vw = 0 To flexgrid.Rows - 1
flexgrid.Row = vw
Printer.CurrentY = Printer.CurrentY + espacioCelda
    
If flexgrid.Row = 0 Then
Printer.Line (1, Printer.CurrentY + espacioCelda)-(MedidaTotal, Printer.CurrentY + espacioCelda)
Printer.Line (1, Printer.CurrentY + espacioCelda)-(MedidaTotal, Printer.CurrentY + espacioCelda)
Printer.CurrentY = Printer.CurrentY - espacioCelda
End If
    
    For i = 0 To flexgrid.Cols - 1
    flexgrid.Col = i
    
        'If Check(i).Value = 1 Then
        
        If IsNumeric(flexgrid.Text) Then
        Printer.CurrentX = CDbl(colw(i)) + medida(i) / 1.28 - Printer.TextWidth(flexgrid.Text)
        Else
        Printer.CurrentX = CDbl(colw(i))
        End If
        
        Printer.Print flexgrid.Text;
        'End If
    
    Next i
    

If vw <> 0 Then
ContRay = ContRay + 1
ContLin = ContLin + 1
End If

If Printer.CurrentY > largomaximo Then
Printer.NewPage
'ImprimirTitulos
'Printer.PaintPicture Picture1.Picture, 1, 1
Printer.CurrentX = MedidaTotal - Printer.TextWidth("Página " & Printer.Page)
Printer.CurrentY = 1
Printer.Print "Página " & Printer.Page
ContLin = 0
ContRay = 0
End If

If ContRay = 5 Then
'Printer.DrawStyle = 2
Printer.CurrentY = Printer.CurrentY + espacioCelda
Printer.Line (1, Printer.CurrentY)-(MedidaTotal, Printer.CurrentY)
ContRay = 0
'Printer.DrawStyle = 0
End If
    
Next vw
    
    
Printer.CurrentY = Printer.CurrentY + espacioCelda
Printer.CurrentX = 1
Printer.Print "Total de Registros: " & vw - 2
Printer.EndDoc
'MsgBox "Listado Imprimiendose", vbInformation

End Sub







