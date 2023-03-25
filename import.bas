Attribute VB_Name = "import"
Sub indicate_import_progress(file As String)
    Application.ScreenUpdating = True
    set_status "Loading... " & file
    Application.ScreenUpdating = False
End Sub

Sub clear_sheet(file As String)
    On Error Resume Next
    Worksheets(file).ListObjects("nice_table").Delete
    On Error GoTo 0
End Sub

Function get_file_xml(file As String) As String
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Const ForReading = 1
    Const ForWriting = 2
    
    Set doc = FSO.OpenTextFile(get_cell_path() & "\" & file, ForReading)
    contents = doc.ReadAll
    doc.Close
    
    ' Workaround for a REALLY painful MSXML bug(?)
    ' If an attribute looks like attribute="&apos;text"
    ' then it MAY! be loaded in as "text" instead of "'text".
    ' So just replace apostrophes directly following " by a curly quote (U+2018)
    ' and sanitize it out later.
    contents = Replace(contents, """&apos;", """" & Chr(&HE2) & Chr(&H80) & Chr(&H98))
    
    ' For the icing on the cake: MSXML2.DOMDocument.LoadXML only works with UTF-16!
    ' So just use Load with a temporary file instead of converting the string
    temp_name = FSO.GetTempName()
    Set doc_output = FSO.CreateTextFile(temp_name)
    doc_output.Write contents
    doc_output.Close
    
    get_file_xml = temp_name
End Function

Sub delete_temp_file(filename As String)
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FSO.DeleteFile filename
End Sub

Sub import_simple(file As String)
    indicate_import_progress file

    Dim XDoc As Object, root As Object

    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    Dim temp_doc_filename As String
    temp_doc_filename = get_file_xml(file)
    success = XDoc.Load(temp_doc_filename)
    delete_temp_file temp_doc_filename
    
    If Not success Then
        MsgBox "Can't import " & file & ", file not found!", vbExclamation
        Exit Sub
    End If
    
    Set root = XDoc.DocumentElement
    
    Dim row As Integer
    row = 1
    Dim schema() As Variant
    Dim schema_max As Integer
    
    If file = "strings.xml" Then
        If Not IsNull(root.getAttribute("max_local_for")) Then
            Worksheets("Controls").Range("B18").value = root.getAttribute("max_local_for")
            schema_max = 5
        Else
            Worksheets("Controls").Range("B18").value = ""
            schema_max = 4
        End If
    
        ReDim schema(schema_max)
        schema(0) = "english"
        schema(1) = "translation"
        schema(2) = "case"
        schema(3) = "explanation"
        schema(4) = "max"
        If schema_max = 5 Then
            schema(5) = "max_local"
        End If
    ElseIf file = "numbers.xml" Then
        schema_max = 3
        ReDim schema(schema_max)
        schema(0) = "value"
        schema(1) = "form"
        schema(2) = "english"
        schema(3) = "translation"
    ElseIf file = "roomnames.xml" Then
        schema_max = 4
        ReDim schema(schema_max)
        schema(0) = "x"
        schema(1) = "y"
        schema(2) = "english"
        schema(3) = "translation"
        schema(4) = "explanation"
    ElseIf file = "roomnames_special.xml" Then
        schema_max = 2
        ReDim schema(schema_max)
        schema(0) = "english"
        schema(1) = "translation"
        schema(2) = "explanation"
    Else
        MsgBox "Can't import " & file & ", schema not handled!", vbExclamation
        Exit Sub
    End If
    
    col_A = 1
    col_Z = schema_max + 1
    
    clear_sheet file
    
    With Worksheets(file)
        For i = 0 To schema_max
            .Cells(1, i + 1).value = schema(i)
        Next i
        
        row = row + 1
        
        For Each subNode In root.ChildNodes
            ' Format as text, don't guess/convert numbers
            .Range(.Cells(row, col_A), .Cells(row, col_Z)).NumberFormat = "@"
            
            ' Comments in roomnames_special.xml are kinda special, we want to keep them.
            If TypeName(subNode) <> "IXMLDOMComment" Then
                For i = 0 To schema_max
                    attr_name = schema(i)
                    .Cells(row, i + 1).value = subNode.getAttribute(attr_name)
                    .Cells(row, i + 1).Errors(xlNumberAsText).Ignore = True
                Next i
            End If
            
            row = row + 1
        Next subNode
        
        Dim table_range
        Set table_range = .Range(.Cells(1, col_A), .Cells(row - 1, col_Z))
        
        .ListObjects.Add(xlSrcRange, table_range, , xlYes).name = "nice_table"
    End With
End Sub

Sub import_strings_plural(ByRef forms() As Boolean, ByRef forms_example() As Integer)
    Dim file As String
    file = "strings_plural.xml"
    
    indicate_import_progress file

    Dim XDoc As Object, root As Object

    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    Dim temp_doc_filename As String
    temp_doc_filename = get_file_xml(file)
    success = XDoc.Load(temp_doc_filename)
    delete_temp_file temp_doc_filename
    
    If Not success Then
        MsgBox "Can't import " & file & ", file not found!", vbExclamation
        Exit Sub
    End If
    
    Set root = XDoc.DocumentElement
    
    num_forms = 0
    For f = 0 To 254
        If forms(f) Then
            num_forms = num_forms + 1
        End If
    Next f
    
    Dim row As Integer
    row = 1
    Dim schema_max As Integer
    If Not IsNull(root.getAttribute("max_local_for")) Then
        schema_max = 6 + num_forms
    Else
        schema_max = 5 + num_forms
    End If
    ReDim schema(schema_max) As String
    
    schema(0) = "english_plural"
    schema(1) = "english_singular"
    schema(2) = "explanation"
    schema(3) = "max"
    schema(4) = "var"
    schema(5) = "expect"
    If schema_max = 6 + num_forms Then
        schema(6) = "max_local"
        sch_ix = 7
    Else
        sch_ix = 6
    End If
    
    For f = 0 To 254
        If forms(f) Then
            schema(sch_ix) = "form " & f & " (ex: " & forms_example(f) & ")"
            sch_ix = sch_ix + 1
        End If
    Next f
    
    col_A = 1
    col_Z = schema_max + 1
    
    clear_sheet file
    
    With Worksheets(file)
        For i = 0 To schema_max
            .Cells(1, i + 1).value = schema(i)
        Next i
        
        row = row + 1
        
        For Each subNode In root.ChildNodes
            ' <string english_plural= english_singular= explanation= max= var= expect=>
            ' Format as text, don't guess/convert numbers
            .Range(.Cells(row, col_A), .Cells(row, col_Z)).NumberFormat = "@"
            
            If TypeName(subNode) <> "IXMLDOMComment" Then
                For i = 0 To schema_max
                    attr_name = schema(i)
                    
                    Dim new_value As String
                    new_value = ""
                    If attr_name Like "form *" Then
                        parts = Split(attr_name, " ", 3)
                        
                        For Each subsubNode In subNode.ChildNodes
                            ' <translation form= translation=>
                            If TypeName(subsubNode) = "IXMLDOMComment" Then
                            ElseIf subsubNode.getAttribute("form") = parts(1) Then
                                new_value = subsubNode.getAttribute("translation")
                            End If
                        Next subsubNode
                    Else
                        new_value = subNode.getAttribute(attr_name)
                    End If
                    
                    .Cells(row, i + 1).value = new_value
                    .Cells(row, i + 1).Errors(xlNumberAsText).Ignore = True
                Next i
            End If
            
            row = row + 1
        Next subNode
        
        Dim table_range
        Set table_range = .Range(.Cells(1, col_A), .Cells(row - 1, col_Z))
        
        .ListObjects.Add(xlSrcRange, table_range, , xlYes).name = "nice_table"
    End With
End Sub

Sub import_cutscenes()
    Dim file As String
    file = "cutscenes.xml"
    
    indicate_import_progress file

    Dim XDoc As Object, root As Object

    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    Dim temp_doc_filename As String
    temp_doc_filename = get_file_xml(file)
    success = XDoc.Load(temp_doc_filename)
    delete_temp_file temp_doc_filename
    
    If Not success Then
        MsgBox "Can't import " & file & ", file not found!", vbExclamation
        Exit Sub
    End If
    
    Set root = XDoc.DocumentElement
    
    Dim row As Integer
    row = 1
    Dim schema_max As Integer
    schema_max = 13
    ReDim schema(schema_max) As String
    
    schema(0) = "id"
    schema(1) = "explanation"
    schema(2) = "speaker"
    schema(3) = "english"
    schema(4) = "translation"
    schema(5) = "case"
    schema(6) = "tt"
    schema(7) = "wraplimit"
    schema(8) = "centertext"
    schema(9) = "pad"
    schema(10) = "pad_left"
    schema(11) = "pad_right"
    schema(12) = "padtowidth"
    schema(13) = "buttons"
    
    col_A = 1
    col_Z = schema_max + 1
    
    clear_sheet file
    
    With Worksheets(file)
        For i = 0 To schema_max
            .Cells(1, i + 1).value = schema(i)
        Next i
        
        row = row + 1
        
        For Each subNode In root.ChildNodes
            ' <cutscene id= explanation=>
            
            If TypeName(subNode) <> "IXMLDOMComment" Then
                Dim script_id As String, script_explanation As String
                script_id = subNode.getAttribute("id")
                script_explanation = subNode.getAttribute("explanation")
                
                For Each subsubNode In subNode.ChildNodes
                    ' <dialogue speaker= english= translation= ...>
                    
                    ' Format as text, don't guess/convert numbers
                    .Range(.Cells(row, col_A), .Cells(row, col_Z)).NumberFormat = "@"
                
                    For i = 0 To schema_max
                        attr_name = schema(i)
                        If attr_name = "id" Then
                            .Cells(row, i + 1).value = script_id
                        ElseIf attr_name = "explanation" Then
                            .Cells(row, i + 1).value = script_explanation
                        ElseIf TypeName(subsubNode) <> "IXMLDOMComment" Then
                            .Cells(row, i + 1).value = subsubNode.getAttribute(attr_name)
                        End If
                        .Cells(row, i + 1).Errors(xlNumberAsText).Ignore = True
                    Next i
                    
                    If TypeName(subsubNode) <> "IXMLDOMComment" Then
                        row = row + 1
                    End If
                Next subsubNode
            End If
        Next subNode
        
        Dim table_range
        Set table_range = .Range(.Cells(1, col_A), .Cells(row - 1, col_Z))
        
        .ListObjects.Add(xlSrcRange, table_range, , xlYes).name = "nice_table"
    End With
End Sub

Sub get_used_forms(ByRef forms() As Boolean, ByRef forms_example() As Integer)
    For f = 0 To 254
        forms_example(f) = 0
    Next f

    Dim row As ListRow
    For Each row In Worksheets("numbers.xml").ListObjects("nice_table").ListRows
        form_str = ListRow_get(row, "form")
        If form_str <> "" Then
            Dim form As Integer
            form = CInt(form_str)
            If form >= 0 And form <= 254 Then
                forms(form) = True
                
                If forms_example(form) = 0 Then
                    ' Yes, we don't want 0 as an example unless it's the only possible example
                    forms_example(form) = ListRow_get(row, "value")
                End If
            End If
        End If
    Next row
End Sub
