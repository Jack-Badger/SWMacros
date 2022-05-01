Attribute VB_Name = "InsertCutThroAllSelectedS"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim swPart As SldWorks.PartDoc
    Set swPart = swModel
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    If swSelMgr.GetSelectedObjectCount2(-1) = 0 Then
    
        Exit Sub
    
    End If
    
    Dim SketchCollection As New VBA.Collection
    
    Dim i As Long
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
    
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelSKETCHES Then
        
            Dim swFeature As SldWorks.Feature
            Set swFeature = swSelMgr.GetSelectedObject6(i, -1)
            
            SketchCollection.Add swFeature
            
        End If

    Next i
    
    If SketchCollection.Count = 0 Then
    
        Exit Sub
        
    End If
    
    Dim swFeatureMgr As SldWorks.FeatureManager
    
    Set swFeatureMgr = swModel.FeatureManager
    
    For i = 1 To SketchCollection.Count
    
        With swFeatureMgr
    
            Set swFeature = SketchCollection(i)
            
            .EditRollback swMoveRollbackBarToAfterFeature, swFeature.Name
             
            swFeature.Select2 False, 0
            
            Dim swCutFeature As SldWorks.Feature
            Set swCutFeature = .FeatureCut4(True, False, False, swEndCondThroughAll, 0, 0.001, 0.001, False, False, False, False, 0, 0, False, False, False, False, False, False, False, False, False, False, swStartSketchPlane, 0, False, False)
        
        End With 'swFeatureMgr
    Next i
    
    swFeatureMgr.EditRollback swMoveRollbackBarToEnd, ""

End Sub



