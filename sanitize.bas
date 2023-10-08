Attribute VB_Name = "sanitize"
Sub sanitize_xml(file As String, filename As String)
    ' This function is called at the end of every export, because MSXML's output is suboptimal,
    ' and also litters the text with ellipsis characters.
    ' But we can get it to be consistent with TinyXML.
    
    Dim contents As String
    contents = read_file(filename)
    
    ' Add XML header and linebreak before root tag
    contents = "<?xml version=""1.0"" encoding=""UTF-8""?>" & Chr(10) & file_comment(file) & Chr(10) & contents

    ' Linebreak before root end tag
    contents = Replace(contents, "</strings>", Chr(10) & "</strings>")
    contents = Replace(contents, "</numbers>", Chr(10) & "</numbers>")
    contents = Replace(contents, "</strings_plural>", Chr(10) & "</strings_plural>")
    contents = Replace(contents, "</cutscenes>", Chr(10) & "</cutscenes>")
    contents = Replace(contents, "</roomnames>", Chr(10) & "</roomnames>")
    contents = Replace(contents, "</roomnames_special>", Chr(10) & "</roomnames_special>")
    
    ' Level 1 tags
    contents = Replace(contents, "<string ", Chr(10) & "    <string ")
    contents = Replace(contents, "</string>", Chr(10) & "    </string>") ' only in strings_plural.xml
    contents = Replace(contents, "<number ", Chr(10) & "    <number ")
    contents = Replace(contents, "<cutscene ", Chr(10) & "    <cutscene ")
    contents = Replace(contents, "</cutscene>", Chr(10) & "    </cutscene>")
    contents = Replace(contents, "<roomname ", Chr(10) & "    <roomname ")
    contents = Replace(contents, "<!-- - -->", Chr(10) & "    <!-- - -->")
    
    ' Level 2 tags
    contents = Replace(contents, "<translation ", Chr(10) & "        <translation ")
    contents = Replace(contents, "<dialogue ", Chr(10) & "        <dialogue ")
    
    ' Decontaminate special characters
    If Worksheets("Controls").Range("B18").value = "" Then
        contents = Replace(contents, ChrW(&H2026), "...") ' ellipsis character, but only outside of CJK
    End If
    contents = Replace(contents, ChrW(&H2018), "&apos;") ' U+2018 curly quote, see import->get_file_xml
    contents = Replace(contents, "'", "&apos;")
    contents = Replace(contents, Chr(13), "") ' Windows carriage return
    
    write_file filename, contents
End Sub

Function file_comment(file As String) As String
    If file = "roomnames.xml" Then
        file_comment = "<!-- You can translate these in-game to get better context! See README.txt -->"
    Else
        file_comment = "<!-- Please read README.txt for information about the language files -->"
    End If
End Function
