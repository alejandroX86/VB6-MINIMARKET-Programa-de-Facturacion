Attribute VB_Name = "ModListObject"
Public Sub ListObject(XFlex As MSHFlexGrid, NombreTabla As String, Image1 As Image, Image2 As Image)
AbrirBase
strSql = "SELECT * FROM [" & NombreTabla & "]"
Dim rs1 As New Recordset
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic

Dim titulos As String
titulos = ""

titulos = "Edit|Del|"

Dim z As Integer
For z = 0 To rs1.Fields.Count - 1
titulos = titulos & Trim(rs1.Fields.Item(z).Name) & "|"
frmListObject.ComboBuscar.AddItem rs1.Fields.Item(z).Name
Next z

titulos = Left(titulos, Len(titulos) - 1)

With XFlex
.Clear
.GridLines = flexGridNone
.Cols = 2
.Rows = 2
.FixedCols = 0
.FixedRows = 1
.Col = 0
.ColSel = 0
.Row = 0
.RowSel = 0
.FormatString = titulos
.Visible = True
.AllowBigSelection = True
.AllowUserResizing = flexResizeColumns
'.SelectionMode = flexSelectionByRow
.ColAlignmentFixed = 3
End With

If Not rs1.EOF Then
rs1.MoveFirst
i = 0
While Not rs1.EOF
i = i + 1
Fila = ""


Fila = "" & Chr(9) & "" & Chr(9)

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

        XFlex.Row = i
        XFlex.Col = 0
        Set XFlex.CellPicture = Image1.Picture

        XFlex.Row = i
        XFlex.Col = 1
        Set XFlex.CellPicture = Image2.Picture


rs1.MoveNext
Wend
End If

rs1.Close

frmListObject.ComboBuscar.ListIndex = 1


AutoFlex XFlex
FlexRayado XFlex, &HFFFFFF, &H8000000F

CerrarBase
End Sub

Public Sub EliminarLinea(NombreTabla As String, id As Variant)
AbrirBase
strSql = "DELETE FROM [" & NombreTabla & "]" & " WHERE ID=" & Val(id)
Dim rs1 As New Recordset
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic
CerrarBase
End Sub






Public Sub BuscarRegistro(XFlex As MSHFlexGrid, NombreTabla As String, Image1 As Image, Image2 As Image, txtBuscar As Variant, Criterio As Variant)
AbrirBase
strSql = "SELECT * FROM [" & NombreTabla & "]"
Dim rs1 As New Recordset
rs1.Open strSql, DB, adOpenKeyset, adLockOptimistic

Dim titulos As String
titulos = ""

titulos = "Edit|Del|"

Dim z As Integer
For z = 0 To rs1.Fields.Count - 1
titulos = titulos & Trim(rs1.Fields.Item(z).Name) & "|"
Next z


If txtBuscar = "" Then
txtBuscar = " "
End If



For z = 0 To rs1.Fields.Count - 1
If rs1.Fields.Item(z).Name = Criterio Then

If rs1.Fields.Item(z).Type = adCurrency Or rs1.Fields.Item(z).Type = adInteger Then
    rs1.Filter = Criterio & "=" & Val(txtBuscar)
    Else
    rs1.Filter = Criterio & " LIKE '*" + txtBuscar + "*'"
End If

End If
Next z


titulos = Left(titulos, Len(titulos) - 1)

With XFlex
.Clear
.GridLines = flexGridNone
.Cols = 2
.Rows = 2
.FixedCols = 0
.FixedRows = 1
.Col = 0
.ColSel = 0
.Row = 0
.RowSel = 0
.FormatString = titulos
.Visible = True
.AllowBigSelection = True
.AllowUserResizing = flexResizeColumns
'.SelectionMode = flexSelectionByRow
.ColAlignmentFixed = 3
End With

If Not rs1.EOF Then
rs1.MoveFirst
i = 0
While Not rs1.EOF
i = i + 1
Fila = ""


Fila = "" & Chr(9) & "" & Chr(9)

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

        XFlex.Row = i
        XFlex.Col = 0
        Set XFlex.CellPicture = Image1.Picture

        XFlex.Row = i
        XFlex.Col = 1
        Set XFlex.CellPicture = Image2.Picture


rs1.MoveNext
Wend
End If

rs1.Close

AutoFlex XFlex
FlexRayado XFlex, &HFFFFFF, &H8000000F

CerrarBase
End Sub






