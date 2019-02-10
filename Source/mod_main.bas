Attribute VB_Name = "mod_main"
Option Explicit
''
''\\ Vincent Poral - vincent.poral@gmail.com - 2018
''\\ Main macros
''
Sub newNumbering()
''\\ Find reference, replace with number and create a csv file containing the reference and the number
''\\ Maximum of 32700 molecule to change in a document.
 Dim str_cmpdReference As String, str_refToReplace As String
 Dim i As Integer, j As Integer
 Dim arr_multiref As Variant, iarr_multiref As Variant
 Dim boo_multiref As Boolean
 Dim key_dic As Variant
 Dim path_refDB As String
 Dim str_refToDisplay As String, str_bkId As String
 Dim dic_molecule As Object
 Dim rng_doc As Range
 Dim var_variable As Variable
'
 Application.ScreenUpdating = False
'
On Error GoTo errHandling
 'ActiveDocument.Save 'Backup
 Set dic_molecule = CreateObject("scripting.dictionary") 'Set a dictionary for the molecule reference
'
 Call numberToReference ' Reverse molecule to reference
 Application.ScreenUpdating = False

 ActiveDocument.Bookmarks.Add Range:=Selection.Range, Name:="_OpenAt"
'
 i = 1 ' i = number for the molecule
 j = 0 ' j = number for the bookmark id
 path_refDB = ActiveDocument.Path & "\" & ActiveDocument.Name & "_refDB.csv"
'
 For Each var_variable In ActiveDocument.Variables
    If InStr(var_variable.Name, "ID") Then var_variable.delete 'Delete variables
 Next
'
 Set rng_doc = ActiveDocument.Content
'
 While rng_doc.Find.Execute(FindText:="\\cmpd\{*\}", Forward:=True, MatchWildcards:=True)
    If rng_doc.Find.Found = True Then
        str_cmpdReference = Right(rng_doc.Text, Len(rng_doc.Text) - 5)
        If InStr(str_cmpdReference, ",") > 0 And InStr(str_cmpdReference, "cmpd{") = 0 _
            And Len(str_cmpdReference) < 200 Then ''\\ Test if multi-reference
             boo_multiref = True
        ElseIf InStr(str_cmpdReference, "cmpd{") <> 0 Or Len(str_cmpdReference) < 2 Then 'test for non close bracket
            MsgBox ("Bracket not close")
            Selection.Collapse (wdCollapseStart)
            Exit Sub
        Else
            boo_multiref = False
        End If
        '
        str_cmpdReference = Left(Replace(str_cmpdReference, "{", ""), Len(str_cmpdReference) - 2)
        str_refToReplace = "\cmpd{" & str_cmpdReference & "}"
        Select Case boo_multiref
            Case False 'If not multireference
                If Not (dic_molecule.Exists(str_cmpdReference)) Then 'if reference not in dictionary -> add it
                    dic_molecule.Add str_cmpdReference, i
                    i = i + 1
                End If
                rng_doc.Text = dic_molecule.Item(str_cmpdReference)
'
            Case True 'If multireference
                arr_multiref = Split(Replace(str_cmpdReference, " ", ""), ",")
                For Each iarr_multiref In arr_multiref 'Add reference to dictionary if there not existing
                    If Not (dic_molecule.Exists(iarr_multiref)) Then
                        dic_molecule.Add iarr_multiref, i
                        i = i + 1
                    End If
                Next
                str_refToDisplay = getMultiText(dic_molecule, arr_multiref) 'get text to display for multi reference
                rng_doc.Text = str_refToDisplay
        End Select
'
        str_bkId = "ID" & j
        rng_doc.Bookmarks.Add ("_ld_" & str_bkId & "_ld_") 'Add bookmark with ID
        rng_doc.Font.Bold = True
        ActiveDocument.Variables.Add Name:=str_bkId, Value:=str_cmpdReference 'Create variable for the document
        j = j + 1
        rng_doc.Collapse wdCollapseEnd
    End If
 Wend
'
 If ActiveDocument.Variables("setCSV") = False Then Application.ScreenUpdating = True: Exit Sub
'
 If Dir(path_refDB) <> "" Then Kill (path_refDB) 'Create csv file with the molecule reference and its number
 Open path_refDB For Append As #1
 Print #1, "Reference; Molecule Number"
 For Each key_dic In dic_molecule
    Print #1, key_dic & ";" & dic_molecule.Item(key_dic)
 Next
 Close #1
 Application.ScreenUpdating = True
Exit Sub

errHandling:
 MsgBox ("newNumbering error: " & Err.Description & " (" & Err.Number & ")")

'ActiveDocument.SaveAs FileName:=str_PathDocument & "numbered_" & str_NameDocument 'Backup"
End Sub
''
''
''
''
Sub numberToReference()
 Dim bk As Bookmark
'
 Application.ScreenUpdating = False
 ActiveDocument.Bookmarks.ShowHidden = True ' to avoid problem with hidden bookmarks
 On Error GoTo errHandling
'
 For Each bk In ActiveDocument.Bookmarks ' loop through all the bookmarks
    If InStr(bk.Name, "_ld_") <> 0 Then ' Find bookmarks with reference _ld_
        bk.Range.Text = "\cmpd{" & ActiveDocument.Variables(Split(bk.Name, "_ld_")(1)).Value & "}" 'Replace number by reference saved in variable
    End If
 Next
 Application.ScreenUpdating = True
 ActiveDocument.Bookmarks.ShowHidden = False
'
 Exit Sub
errHandling:
 MsgBox ("numberToReference error: " & Err.Description & " (" & Err.Number & ")")
 Application.ScreenUpdating = True
 ActiveDocument.Bookmarks.ShowHidden = False
'
End Sub
''
''
''
''
Sub insertCompound()
 Application.ScreenUpdating = False
 Selection.Text = "\cmpd{}"
 Selection.Collapse (wdCollapseEnd)
 Selection.Move wdCharacter, -1
 Application.ScreenUpdating = True
End Sub




