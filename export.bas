Attribute VB_Name = "export"
Sub indicate_export_progress(file As String)
    set_status "Saving... " & file
End Sub

Sub export_simple(file As String)
    indicate_export_progress file

    Dim root_name As String
    Dim elem_name As String
    
    If file = "strings.xml" Then
        root_name = "strings"
        elem_name = "string"
    ElseIf file = "numbers.xml" Then
        root_name = "numbers"
        elem_name = "number"
    ElseIf file = "roomnames.xml" Then
        root_name = "roomnames"
        elem_name = "roomname"
    ElseIf file = "roomnames_special.xml" Then
        root_name = "roomnames_special"
        elem_name = "roomname"
    Else
        MsgBox "Can't export " & file & ", name not recognized!", vbExclamation
        Exit Sub
    End If

    Dim XDoc As Object, root As Object, elem As Object
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    Set root = XDoc.createElement(root_name)
    XDoc.appendChild root
    Dim attr As Object
    
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Worksheets(file).ListObjects("nice_table")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "No table for " & file, vbExclamation
        Exit Sub
    End If

    Dim row As ListRow
    For Each row In tbl.ListRows
        If file = "roomnames_special.xml" And ListRow_get(row, "english") = "" Then
            Set elem = XDoc.createComment(" - ")
            root.appendChild elem
        Else
            Set elem = XDoc.createElement(elem_name)
            root.appendChild elem
    
            For Each col In row.Parent.ListColumns
                key = col.name
                value = ListRow_get(row, col.name)
    
                If (file = "strings.xml" And key = "max" And value = "") _
                Or (file = "numbers.xml" And key = "form" And ListRow_get(row, "value") = "lots") _
                Or (file = "numbers.xml" And (key = "english" Or key = "translation") And ListRow_get(row, "english") = "") Then
                    ' Don't include this attribute
                Else
                    Set attr = XDoc.createAttribute(key)
                    attr.NodeValue = value
                    elem.setAttributeNode attr
                End If
            Next
        End If
    Next row
        
    ' Save the XML file
    Dim filename As String
    filename = get_cell_path() & "\" & file
    XDoc.Save (filename)
    
    sanitize_xml file, filename

End Sub

Sub export_strings_plural()
    Dim file As String
    file = "strings_plural.xml"
    
    indicate_export_progress file

    Dim XDoc As Object, root As Object, elem As Object, subElem As Object
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    Set root = XDoc.createElement("strings_plural")
    XDoc.appendChild root
    Dim attr As Object
    
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Worksheets(file).ListObjects("nice_table")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "No table for " & file, vbExclamation
        Exit Sub
    End If

    Dim row As ListRow
    For Each row In tbl.ListRows
        Set elem = XDoc.createElement("string")
        root.appendChild elem
        
        For Each col In Array("english_plural", "english_singular", "explanation", "max", "expect")
            Dim key As String
            key = col
            
            value = ListRow_get(row, key)
            If Not ((key = "max" Or key = "expect") And value = "") Then
                Set attr = XDoc.createAttribute(key)
                attr.NodeValue = value
                elem.setAttributeNode attr
            End If
        Next

        ' Now find each plural form column
        For Each col In row.Parent.ListColumns
            If col.name Like "form *" Then
                parts = Split(col.name, " ", 3)
                
                Set subElem = XDoc.createElement("translation")
                elem.appendChild subElem
                
                Set attr = XDoc.createAttribute("form")
                attr.NodeValue = parts(1)
                subElem.setAttributeNode attr

                Set attr = XDoc.createAttribute("translation")
                attr.NodeValue = ListRow_get(row, col.name)
                subElem.setAttributeNode attr
            End If
        Next
    Next row
        
    ' Save the XML file
    Dim filename As String
    filename = get_cell_path() & "\" & file
    XDoc.Save (filename)
    
    sanitize_xml file, filename
End Sub

Sub export_cutscenes()
    Dim file As String
    file = "cutscenes.xml"
    
    indicate_export_progress file

    Dim XDoc As Object, root As Object, elem As Object, subElem As Object
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    Set root = XDoc.createElement("cutscenes")
    XDoc.appendChild root
    Dim attr As Object
    
    last_sid = "none yet"

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Worksheets(file).ListObjects("nice_table")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "No table for " & file, vbExclamation
        Exit Sub
    End If

    Dim row As ListRow
    For Each row In tbl.ListRows
        sid = ListRow_get(row, "id")
        If sid <> last_sid Then
            ' New cutscene
            Set elem = XDoc.createElement("cutscene")
            root.appendChild elem
            
            Set attr = XDoc.createAttribute("id")
            attr.NodeValue = sid
            elem.setAttributeNode attr
            
            Set attr = XDoc.createAttribute("explanation")
            attr.NodeValue = ListRow_get(row, "explanation")
            elem.setAttributeNode attr
            
            last_sid = sid
        End If
        
        Set subElem = XDoc.createElement("dialogue")
        elem.appendChild subElem
        
        For Each col In Array( _
            "speaker", "english", "translation", _
            "case", "tt", "wraplimit", "centertext", _
            "pad", "pad_left", "pad_right", "padtowidth" _
        )
            Dim key As String
            key = col
            
            value = ListRow_get(row, key)
            If key = "translation" Or value <> "" Then
                Set attr = XDoc.createAttribute(key)
                attr.NodeValue = value
                subElem.setAttributeNode attr
            End If
        Next
        
    Next row
        
    ' Save the XML file
    Dim filename As String
    filename = get_cell_path() & "\" & file
    XDoc.Save (filename)
    
    sanitize_xml file, filename
End Sub
