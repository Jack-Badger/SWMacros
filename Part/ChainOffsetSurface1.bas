Attribute VB_Name = "ChainOffsetSurface1"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
    
        Exit Sub
        
    End If
    
    If Not swModel.GetType = swDocumentTypes_e.swDocPART Then
    
        swApp.SendMsgToUser2 "This macro only valid for Part Documents.", swMbInformation, swMbOk
        Exit Sub
    
    End If
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    If (swSelMgr.GetSelectedObjectCount2(-1) = 0) Or (Not swSelMgr.GetSelectedObjectType3(1, -1) = swSelectType_e.swSelFACES) Then
    
        swApp.SendMsgToUser2 "Please Select a Face and try again.", swMbInformation, swMbOk
        Exit Sub
    
    End If
    
    Dim Flip As Boolean
    Flip = swApp.SendMsgToUser2("Flip Direction?", swMbQuestion, swMbYesNo) = swMbHitYes
    
    Dim Count As Long
    Count = VBA.InputBox("How Many?", "Offset Face Chain")
    
    Dim swEntity As SldWorks.Entity
    Set swEntity = swSelMgr.GetSelectedObject6(1, -1)
    
    Dim i As Long
    For i = 1 To Count
    
        swEntity.Select4 False, Nothing
    
        swModel.InsertOffsetSurface 0.1, Flip
        
        Dim swFeature As SldWorks.Feature
        Set swFeature = swModel.Extension.GetLastFeatureAdded
    
        Set swEntity = swFeature.GetFaces(0)
    
    Next i
    
End Sub

