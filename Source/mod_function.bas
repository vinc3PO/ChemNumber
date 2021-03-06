Attribute VB_Name = "mod_function"
Option Explicit
''
''\\ Vincent Poral - vincent.poral@gmail.com - 2018
''\\ Function used for mutli reference
''
Function getMultiText(dic_REF As Object, var_list As Variant) As String
'' order multi reference and return reorder list of the references
Dim i As Integer, j As Integer, int_init As Integer, int_end As Integer
Dim str_ref As String, str_multi As String
'
For i = 0 To UBound(var_list)
    var_list(i) = dic_REF.Item(var_list(i))
Next i
'
var_list = sortList(var_list) ' Sorting the number form lower to higher
For i = 0 To UBound(var_list) ' Loop to find consecutive numbers to concatenate them as x-z
    j = i
    int_init = var_list(i)
    int_end = 0
    On Error GoTo endWhile
    While (CInt(var_list(j)) + 1) = CInt(var_list(j + 1))
        int_end = var_list(j + 1)
        j = j + 1
    Wend
endWhile:
    If int_end < int_init Then ' If no consecutive number add x or x,
        If i = 0 Then str_ref = int_init Else str_ref = str_ref & ", " & int_init
    Else ' if consecutive numbers add x,y or x,z (need 3+ consecutive for dash)
        If int_end = int_init + 1 Then
            str_multi = int_init & ", " & int_end
        Else
            str_multi = int_init & "-" & int_end
        End If
        If i = 0 Then str_ref = str_multi Else str_ref = str_ref & ", " & str_multi
    End If
    i = j
Next i
getMultiText = str_ref
End Function

Function sortList(var_list As Variant) As Variant()
'
 Dim var_sortedList() As Variant
 Dim i As Integer, j As Integer
 Dim int_low As Variant, int_high As Variant, int_value As Variant, int_lastLow As Variant, int_lastHigh As Variant
 Dim boo_start As Boolean, boo_loop As Boolean
'
 'On Error GoTo errHandling
 int_low = var_list(0)
 int_high = var_list(0)
 boo_start = True
 j = 0
 While j < UBound(var_list) + 1
    If boo_start = True Then
        boo_start = False
        For i = 0 To UBound(var_list)
            int_value = var_list(i)
            If (int_value - int_low) < 0 Then
                int_low = int_value
            End If
        Next
        ReDim Preserve var_sortedList(j)
        var_sortedList(j) = int_low
        j = j + 1
    Else
        int_lastLow = int_low
        boo_loop = True
        For i = 0 To UBound(var_list)
            int_value = var_list(i)
            If int_lastLow - int_value < 0 Then
                If boo_loop = True Then int_low = int_value: boo_loop = False
                If (int_lastLow - int_value) - (int_lastLow - int_low) > 0 Then
                    int_low = int_value
                End If
            End If
        Next
        If int_lastLow = int_low Then GoTo WExit
        ReDim Preserve var_sortedList(j)
        var_sortedList(j) = int_low
        j = j + 1
    End If
 Wend
'
WExit:
 sortList = var_sortedList
 Exit Function
'
errHandling:
 Debug.Print ("sort list error")
 sortList = var_list
End Function

