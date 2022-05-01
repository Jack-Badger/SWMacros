Attribute VB_Name = "LinearPatternSuppressionE"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim swEqMgr As SldWorks.EquationMgr
    Set swEqMgr = swModel.GetEquationMgr
    
    Dim num As Long
    num = VBA.InputBox("Number of Instances", "Linear Pattern Suppression")
    
    Dim GlobalVariableName As String
    GlobalVariableName = VBA.InputBox("Global Variable Name", "Linear Pattern Suppression", "NUM")
    
    Dim ComponentName As String
    ComponentName = VBA.InputBox("Component Name", "Linear Pattern Suppression")
    
    Dim i As Long
    
    For i = 1 To num
        
        swEqMgr.Add2 -1, """" & ComponentName & "<" & i & ">.Part"" = IF(""" & GlobalVariableName & """ > " & i - 1 & ",""unsuppressed"", ""suppressed"")", True
    
    Next i
    
    swModel.EditRebuild3

End Sub
