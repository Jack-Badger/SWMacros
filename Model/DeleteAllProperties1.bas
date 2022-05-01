Attribute VB_Name = "DeleteAllProperties1"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Updates swModel, False
    
        On Error GoTo RestoreUpdates
        
        swApp.SendMsgToUser2 ProcessAllConfigurations(swModel), swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
    
    End If
    
RestoreUpdates:
    Updates swModel, True
    
    Set swApp = Nothing

End Sub


Private Function ProcessAllConfigurations(swModel As SldWorks.ModelDoc2) As String

    Dim DeletionsCount As Long
    
    Dim ConfigCount As Long
    
    Dim swConfig As SldWorks.Configuration

    Set swConfig = swModel.GetActiveConfiguration
    
    Dim ActiveConfigName As String
    ActiveConfigName = swConfig.Name
    
    Dim vConfigNames As Variant
    
    vConfigNames = swModel.GetConfigurationNames
    
    Dim ConfigName As String
    
    Dim i As Long
    
    For i = 0 To UBound(vConfigNames)
    
        ConfigCount = ConfigCount + 1
    
        ConfigName = vConfigNames(i)
        
        swModel.ShowConfiguration2 ConfigName
        
        DeletionsCount = DeletionsCount + DeleteAllConfigSpecificProperties(swModel.GetActiveConfiguration)
            
    Next i
    
    Call swModel.ShowConfiguration2(ActiveConfigName)
    
    ProcessAllConfigurations = "Deleted " & DeletionsCount & " total properties across " & ConfigCount & " configurations."
    
End Function


Private Function DeleteAllConfigSpecificProperties(swConfig As SldWorks.Configuration) As Long

    Dim DeletionsCount As Long

    Dim swPropMgr As SldWorks.CustomPropertyManager
    Set swPropMgr = swConfig.CustomPropertyManager

    If swPropMgr.Count > 0 Then

        Dim vPropNames As Variant
        
        vPropNames = swPropMgr.GetNames
        
        Dim i As Long
        
        For i = 0 To UBound(vPropNames)
        
            swPropMgr.Delete2 vPropNames(i)
            
            DeletionsCount = DeletionsCount + 1
        
        Next i

    End If
    
    DeleteAllConfigSpecificProperties = DeletionsCount
    
End Function

Private Sub Updates(swModel As SldWorks.ModelDoc2, update As Boolean)

    Dim swView As SldWorks.ModelView
    Set swView = swModel.ActiveView
    swView.EnableGraphicsUpdate = update
    swModel.FeatureManager.EnableFeatureTree = update
    swModel.FeatureManager.EnableFeatureTreeWindow = update
    
End Sub
