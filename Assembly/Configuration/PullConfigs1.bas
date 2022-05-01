Attribute VB_Name = "PullConfigs1"
' PullConfigs.swp
' Will PULL configurations from the selected component and recreate them in the Assembly.
' The selected Component will then be configured 1:1 with the Assembly.
' by 369 12.03.21
' Created in SW2018

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swAssem As SldWorks.AssemblyDoc
Dim swComp As SldWorks.Component2
Dim swAssModel As SldWorks.ModelDoc2
Dim swCompModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr
Dim vConfigNames As Variant
Dim swConfig As SldWorks.Configuration
Dim CompID As Long

Sub main()

    Set swApp = Application.SldWorks
    
    Set swAssModel = swApp.ActiveDoc
    
    Set swAssem = swAssModel
    
    ' TODO: Error Checking / Validation
    
    Set swSelMgr = swAssModel.SelectionManager
    
    Set swComp = swSelMgr.GetSelectedObject6(1, -1)
    
    CompID = swComp.GetID
    
    Set swCompModel = swComp.GetModelDoc2
    
    vConfigNames = swCompModel.GetConfigurationNames
    
    Set swCompModel = Nothing
    
    Dim i As Long
    
    For i = 0 To UBound(vConfigNames)
    
        Set swConfig = swAssModel.AddConfiguration3(vConfigNames(i), "", "", 0)
        
        swAssem.GetComponentByID(CompID).ReferencedConfiguration = vConfigNames(i)
        
        swAssem.EditRebuild
    
    Next i
    
    Set swConfig = Nothing
    Set swSelMgr = Nothing
    Set swCompModel = Nothing
    Set swApp = Nothing
    Set swComp = Nothing
    Set swAssModel = Nothing
    
End Sub
