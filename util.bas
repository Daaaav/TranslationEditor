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

Function read_file(file As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile file
    read_file = st.ReadText(-1)
    st.Close
End Function

Sub write_file(file As String, contents As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    
    st.Charset = "utf-8"
    st.Open
    st.WriteText contents, 0
    
    'Strip out UTF-8 BOM
    'Creates a temporary byte array to store BOM-less data
    st.Position = 0
    st.Type = 1
    st.Position = 3
    
    Dim byteData() As Byte
    byteData = st.Read
    st.Close
    
    ' Write final text
    st.Open
    st.Write byteData
    st.SaveToFile file, 2
    st.Close
End Sub

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
