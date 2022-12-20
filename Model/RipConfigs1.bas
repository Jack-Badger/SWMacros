Attribute VB_Name = "RipConfigs1"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel.GetSaveFlag() Then
    
        swApp.SendMsgToUser2 "Config Ripping cannot requires a saved file!", swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
        Exit Sub
        
    End If
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    Dim vConfigs As Variant
    
    If swSelMgr.GetSelectedObjectCount2(-1) > 0 Then
    
        vConfigs = GetSelectedConfigNames(swSelMgr)

    Else
    
        vConfigs = swModel.GetConfigurationNames
    
    End If
    
    Dim swConfig As SldWorks.Configuration
    
    Dim SaveFolder As String
    'SaveFolder = swApp.GetCurrentMacroPathFolder
    SaveFolder = "C:\temp"
    
    Dim SaveName As String
    SaveName = InputBox("Save Base Name", "Config Ripper")
    
    If SaveName = "" Then
    
        Exit Sub
    
    End If
      
    Dim i As Long
    For i = 0 To UBound(vConfigs)
    
        Set swModel = swApp.ActiveDoc
        Dim bs As Boolean
        
        bs = swModel.ShowConfiguration2(vConfigs(i))
        
        Set swConfig = swModel.GetActiveConfiguration
        
        Dim j As Long
        
        For j = 0 To UBound(vConfigs)
        
            If j <> i Then
            
            bs = swModel.DeleteConfiguration2(vConfigs(j))
            
            End If
            
        Next j
        
        Dim swFSWarning As swFileSaveWarning_e
        Dim swFSError As swFileSaveError_e
        
        swModel.ViewZoomtofit2
        
        bs = swModel.Extension.SaveAs(SaveFolder & "\" & SaveName & Replace(vConfigs(i), "\", "_") & ".SLDPRT", swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Copy, Nothing, swFSError, swFSWarning)
        
        bs = swModel.ReloadOrReplace(False, swModel.GetPathName, True)
    
    Next i
    
End Sub

Private Function GetSelectedConfigNames(swSelMgr As SldWorks.SelectionMgr) As String()

    Dim totalCount As Long
    totalCount = swSelMgr.GetSelectedObjectCount2(-1)
    
    If totalCount = 0 Then Exit Function
    
    Dim configNames() As String
    
    ReDim configNames(totalCount - 1)
    
    Dim foundCount As Long
    
    Dim i As Long
    For i = 0 To UBound(configNames)
    
        If swSelMgr.GetSelectedObjectType3(i + 1, -1) = swSelectType_e.swSelCONFIGURATIONS Then
        
            Dim swConfig As SldWorks.Configuration
            Set swConfig = swSelMgr.GetSelectedObject6(i + 1, -1)
            
            configNames(foundCount) = swConfig.Name
            foundCount = foundCount + 1
            
        End If
    Next i
    
    If foundCount = 0 Then Exit Function
    
    ReDim Preserve configNames(foundCount - 1)
    
    GetSelectedConfigNames = configNames
    
End Function
