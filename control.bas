Attribute VB_Name = "control"
Sub button_setpath_click()
    folder_dialog
End Sub

Sub button_resetpath_click()
    Worksheets("Controls").Range("A7:B7").value = ""
End Sub

Sub button_fullreset_click()
    answer = MsgBox("Clear all sheets to be blank?", vbYesNo)
    
    If answer <> vbYes Then
        Exit Sub
    End If
    
    clear_sheet "strings.xml"
    clear_sheet "numbers.xml"
    clear_sheet "strings_plural.xml"
    clear_sheet "cutscenes.xml"
    clear_sheet "roomnames.xml"
    clear_sheet "roomnames_special.xml"
    
    Worksheets("Controls").Range("B18").value = ""
    
    set_status "No data loaded"
End Sub

Sub button_import_click()
    If cell_path_empty() Then
        MsgBox "Please set the path first.", vbExclamation
        Exit Sub
    End If
    
    answer = MsgBox("Load in all XML files now?" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "This will overwrite all sheets with data from the XML files.", vbYesNo)
    
    If answer <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False

    import_simple "strings.xml"
    import_simple "numbers.xml"
    
    Dim forms(254) As Boolean
    Dim forms_example(254) As Integer
    get_used_forms forms, forms_example
    
    import_strings_plural forms, forms_example
    import_cutscenes
    import_simple "roomnames.xml"
    import_simple "roomnames_special.xml"

    Application.ScreenUpdating = True
    
    set_status "Loading complete!"
End Sub

Sub button_reform_strings_plural_click()
    If cell_path_empty() Then
        MsgBox "Please set the path first.", vbExclamation
        Exit Sub
    End If
    
    answer = MsgBox("This will load only strings_plural.xml, based on the forms you have set in the numbers.xml sheet." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Reload now?", vbYesNo)
    
    If answer <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim forms(254) As Boolean
    Dim forms_example(254) As Integer
    get_used_forms forms, forms_example
    
    import_strings_plural forms, forms_example

    Application.ScreenUpdating = True
    
    set_status "Loading complete!"
End Sub

Sub button_export_click()
    If cell_path_empty() Then
        MsgBox "Please set the path first.", vbExclamation
        Exit Sub
    End If

    export_simple "strings.xml"
    export_simple "numbers.xml"
    
    export_strings_plural
    export_cutscenes
    export_simple "roomnames.xml"
    export_simple "roomnames_special.xml"

    set_status "Saving complete!"
End Sub

Sub folder_dialog()
    Dim objShell
    Set objShell = CreateObject("Shell.Application")

    Dim objFolder
    Set objFolder = objShell.BrowseForFolder(0, "Select the language folder:", 0)

    If objFolder Is Nothing Then
        Exit Sub
    End If
    
    Worksheets("Controls").Range("A7").value = "Path:"
    Worksheets("Controls").Range("B7").value = objFolder.Self.path
End Sub

Function cell_path_empty() As Boolean
    cell_path_empty = (get_cell_path() = "")
End Function

Function get_cell_path() As String
    get_cell_path = Worksheets("Controls").Range("B7").value
End Function

Sub set_status(message As String)
    Worksheets("Controls").Range("B16").value = message
End Sub
