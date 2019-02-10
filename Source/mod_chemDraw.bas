Attribute VB_Name = "mod_chemDraw"
Option Explicit

Sub modifyChemDraw()
''\\
''\\Vincent Poral - vincent.poral@gmail.com - 2019
''\\Sub use to modify CDXML file to automatise the numbering of the molecules in a word document
''\\
Dim obj_XML As Object
Dim oSeqNode As Object
Dim oSeqNodes As Object
Dim dic_molecule As Object
Dim col_scheme As Collection
Dim itm_scheme As Variant
Dim str_time As String
Dim ole_object As InlineShape
'On Error GoTo errHandling
Set obj_XML = CreateObject("MSXML2.DOMDocument") 'Create object to read XML type file | Might raise error
Set dic_molecule = CreateObject("scripting.dictionary") 'Create dictionnary to sort molecule and reference
'
Set dic_molecule = getDB() 'Fill-up dictionnary
obj_XML.async = False: obj_XML.validateOnParse = False 'Option to read CDXML as not true XML
obj_XML.SetProperty "ProhibitDTD", False
str_time = Format(Now, "yymmddhhmmss") 'time stamp for backup folder
Set col_scheme = getScheme(str_time) 'Get collection containing the path to all "cdxml" scheme
If col_scheme.Count = 0 Then Exit Sub

For Each itm_scheme In col_scheme 'Loop through CDXML files
    If obj_XML.Load(itm_scheme) Then
        obj_XML.Save (Replace(itm_scheme, "\scheme\", "\scheme\Backup_" & str_time & "\")) 'Save backup
       ' The document loaded successfully.
       ' Now do something intersting.
       DisplayNode obj_XML.ChildNodes, dic_molecule
    Else
       ' The document failed to load.
       ' See the previous listing for error information.
    End If
    obj_XML.Save (itm_scheme)
Next

Set obj_XML = Nothing
updateCDLink 'Update link
Exit Sub
errHandling:
MsgBox (Err.Description)
On Error Resume Next
Set obj_XML = Nothing
End Sub


Function getScheme(str_time As String) As Collection
''\\
''\\Vincent Poral - vincent.poral@gmail.com - 2019
''\\Function used to create a collection of CDXML files in the child folder "\scheme"
''\\
Dim fso As Object
Dim fol_scheme As Object
Dim fil_scheme As Object
'
On Error Resume Next
Set getScheme = New Collection
'
If Dir(ActiveDocument.Path & "\scheme", vbDirectory) <> "" Then
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fol_scheme = fso.GetFolder(ActiveDocument.Path & "\scheme")
    For Each fil_scheme In fol_scheme.Files
        If LCase(Right(fil_scheme.Name, 6)) = ".cdxml" Then getScheme.Add fil_scheme.Path
    Next
End If
If getScheme.Count > 0 Then fso.CreateFolder (ActiveDocument.Path & "\scheme\Backup_" & str_time)
End Function

Function getDB() As Object
''\\
''\\Vincent Poral - vincent.poral@gmail.com - 2019
''\\Function used to create a database of the reference and numbers. Taken from the CSV file.
''\\
Dim dic As Object
Dim str_path As String
Dim i As Integer
Dim var_refs As Variant
Dim linefromFile As String
Set getDB = CreateObject("scripting.dictionary")
Dim fDialog As FileDialog, result As Integer
Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
'
On Error Resume Next
fDialog.AllowMultiSelect = False
fDialog.Title = "Select a file"
fDialog.InitialFileName = ActiveDocument.Path
fDialog.Filters.Clear
fDialog.Filters.Add "CSV files", "*.csv"
fDialog.Filters.Add "All files", "*.*"
'
If fDialog.Show = -1 Then
    str_path = fDialog.SelectedItems(1)
    Open str_path For Input As #1
    i = 0
    Do Until EOF(1)
        Line Input #1, linefromFile
        If i <> 0 Then
            var_refs = Split(linefromFile, ";")
            getDB.Add var_refs(0), var_refs(1)
        End If
    i = i + 1
    Loop
    Close #1
End If
'
End Function

Sub insertOleScheme()
''\\
''\\Vincent Poral - vincent.poral@gmail.com - 2019
''\\sub used to insert a scheme in the word document.
''\\

Dim dia_file As Object
Dim str_item As String
Dim str_path As String
str_path = ActiveDocument.Path
If Len(Dir(ActiveDocument.Path & "\scheme\", vbDirectory)) <> 0 Then str_path = ActiveDocument.Path & "\scheme\"
Set dia_file = Application.FileDialog(3)
With dia_file
    .Title = "Select a file"
    .AllowMultiSelect = False
    .InitialFileName = str_path
    .Filters.Clear
    .Filters.Add "File type", "*.cdxml"
    If .Show <> -1 Then Exit Sub
    str_item = .SelectedItems(1)
End With
Set dia_file = Nothing
Selection.InlineShapes.AddOLEObject ClassType:="ChemDraw.Document.6.0", _
    FileName:="""" & str_item & """", LinkToFile:=True, _
    DisplayAsIcon:=False
End Sub


Function findInString(str_string As String) As Collection
''\\
''\\Vincent Poral - vincent.poral@gmail.com - 2019
''\\Function used to find refrence (\ref{}) in a string.
''\\
 Dim i As Integer, j As Integer
 Dim str_char As String
 Dim str_found As String
 Dim boo_exitChar As Boolean
'
On Error Resume Next
 Set findInString = New Collection
 boo_exitChar = False
 If InStr(str_string, "\{") <> 0 Then
    For i = 1 To Len(str_string)
        str_char = Mid(str_string, i, 1)
        If str_char = "\" And i + 1 < Len(str_string) Then
            If Mid(str_string, i + 1, 1) = "{" Then
                str_found = ""
                For j = i + 2 To Len(str_string)
                    If Mid(str_string, j, 1) = "}" Then boo_exitChar = True
                    If boo_exitChar = False Then
                        str_found = str_found & Mid(str_string, j, 1)
                    ElseIf boo_exitChar = True Then
                        boo_exitChar = False
                        'Debug.Print (str_found)
                        findInString.Add str_found
                        i = j
                        GoTo nextloop
                    End If
                Next
                i = j
            End If
        End If
nextloop:
    Next
 End If
End Function

Function updateCDLink()
''\\
''\\Vincent Poral - vincent.poral@gmail.com - 2019
''\\Function used to refresh link.
''\\
Dim ole_object As InlineShape
On Error Resume Next
For Each ole_object In ActiveDocument.InlineShapes
    If InStr(ole_object.OLEFormat.ClassType, "ChemDraw.Document") And ole_object.Field.Type = wdFieldLink Then
        ole_object.Field.Update
    End If
Next
End Function


Sub DisplayNode(ByRef Nodes As Object, dic_molecule As Object)
Dim xNode As Object
Dim col_collection As Collection
Dim var_collection As Variant
'
For Each xNode In Nodes
    If xNode.NodeType = 3 Then ' If Nodetype = Node_Text
        If InStr(xNode.NodeValue, "\{") <> 0 Then
            Set col_collection = findInString(xNode.NodeValue)
            If col_collection.Count > 0 Then
                For Each var_collection In col_collection
                    If dic_molecule.Exists(var_collection) Then
                        xNode.NodeValue = Replace(xNode.NodeValue, "\{" & var_collection & "}", dic_molecule.Item(var_collection)) 'replace reference by number
                    'Debug.Print (var_collection)
                    End If
                Next
            End If
        End If
   End If
   If xNode.HasChildNodes Then
      DisplayNode xNode.ChildNodes, dic_molecule
   End If
Next xNode
End Sub

