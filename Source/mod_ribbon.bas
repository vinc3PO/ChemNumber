Attribute VB_Name = "mod_ribbon"
Option Explicit
    Dim rib_chemNumbering As IRibbonUI
    
Sub init_chemNumbering(Ribbon As IRibbonUI)
 Set rib_chemNumbering = Ribbon
End Sub

Sub getEnabled(control As IRibbonControl, ByRef returnedVal)
 returnedVal = True
End Sub

Sub refToNumber(control As IRibbonControl)
 newNumbering
End Sub

Sub numberToRef(control As IRibbonControl)
 numberToReference
End Sub

Sub insertRef(control As IRibbonControl)
 insertCompound
End Sub

Sub CSVToggle(control As IRibbonControl, pressed As Boolean)
 ActiveDocument.Variables("setCSV").Value = pressed
End Sub

Sub schemeNumbering(control As IRibbonControl)
 modifyChemDraw
End Sub

Sub insertScheme(control As IRibbonControl)
 insertOleScheme
End Sub

Sub getPressed(control As IRibbonControl, ByRef returnedVal)
 On Error GoTo errHandling
 returnedVal = ActiveDocument.Variables("setCSV").Value
 Exit Sub
'
errHandling:
 ActiveDocument.Variables.Add "setCSV", "True"
 returnedVal = True
End Sub

Sub refreshScheme(control As IRibbonControl)
 updateCDLink
End Sub





