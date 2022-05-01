Attribute VB_Name = "GetFeatureID1"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        Exit Sub
    End If
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager

    If swSelMgr.GetSelectedObjectCount2(-1) = 0 Then
    
        Exit Sub
    
    End If

    Dim i As Long
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        
        Dim swFeature As SldWorks.Feature
        Set swFeature = swSelMgr.GetSelectedObject6(i, -1)
        
        swApp.SendMsgToUser2 swFeature.Name & ", Id:" & swFeature.GetID, swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
    
    Next i
    
End Sub

