
Function CellUsesLiteralValue(cell As Range) As Boolean
    If Not cell.HasFormula Then
        CellUsesLiteralValue = False
    Else
        CellUsesLiteralValue = cell.formula Like "*[=^/*+-/()><, ]#*"
    End If
End Function
Sub FindLiteralsInWorkbook()
Dim C As Range, A As Range, Addresses As String, i As Single
Dim cell As Range
Sheets.Add
Set X = ActiveSheet
X.Range("A1") = "Link"
X.Range("B1") = "Sheet"
X.Range("C1") = "Formula"
i = 1
For Each sh In ActiveWorkbook.Worksheets
    On Error Resume Next
    Set C = sh.Cells.SpecialCells(xlConstants)
    If C Is Nothing Then
      Set C = sh.Cells.SpecialCells(xlFormulas)
    Else
      Set C = Union(C, sh.Cells.SpecialCells(xlFormulas))
    End If
    For Each A In C.Areas
      For Each cell In A
         If CellUsesLiteralValue(cell) = True Then
              i = i + 1
              X.Hyperlinks.Add Anchor:=X.Range("A" & i), _
              Address:=ActiveWorkbook.Path & "\" & ActiveWorkbook.name, _
              SubAddress:=sh.name & "!" & cell.Address, _
              TextToDisplay:=sh.name & "!" & cell.Address
              X.Range("B" & i) = sh.name
              X.Range("C" & i) = "'" & cell.formula
         End If
      Next cell
    Next A
    On Error GoTo 0
Next sh
X.Columns("A:E").AutoFit
End Sub

