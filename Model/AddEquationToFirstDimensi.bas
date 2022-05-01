Attribute VB_Name = "AddEquationToFirstDimensi"
' Add Equation to First Dimension of Last Feature Added
' https://forum.solidworks.com/thread/245256

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Sub
    
    Dim swFeature As SldWorks.Feature
    Set swFeature = swModel.Extension.GetLastFeatureAdded
    If swFeature Is Nothing Then Exit Sub
    
    Dim swDispDim As SldWorks.DisplayDimension
    
    Set swDispDim = swFeature.GetFirstDisplayDimension
    If swDispDim Is Nothing Then Exit Sub
    
    Dim equation As String
    equation = """" & swDispDim.GetNameForSelection & """ = sqr(""XC""*""XC""+""YC""*""YC"")*1in"
    
    Dim swEqnMgr As SldWorks.EquationMgr
    Set swEqnMgr = swModel.GetEquationMgr
        
    Debug.Print equation
    
    Dim result As Long
        
    result = swEqnMgr.Add2(-1, equation, True)

    Debug.Print result
    
End Sub
