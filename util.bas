Attribute VB_Name = "util"
Function ListRow_get(row As ListRow, name As String) As String
    'row.Range.Columns("B").value
    'MsgBox TypeName(row) & " " & TypeName(row.Parent) & " " & row.Parent.DataBodyRange.Address()
    'MsgBox row.Parent.HeaderRowRange.Address()
    
    For Each col In row.Parent.ListColumns
        If name = col.name Then
            ListRow_get = Application.Intersect(row.Range, col.Range).value
            Exit Function
        End If
    Next col
End Function


' Just some testing stuff

Sub test_playground()
    'straa = "form=1 (ex: 5)"
    straa = "english"
    
    firstpart = Split(straa, " ", 2)
    
    MsgBox firstpart(0)
    MsgBox firstpart(1)
End Sub

Sub aa()
    Worksheets("Controls").Cells(20, 2).value = "test"
End Sub
