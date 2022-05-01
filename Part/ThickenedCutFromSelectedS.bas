Attribute VB_Name = "ThickenedCutFromSelectedS"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
        
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    If swSelMgr.GetSelectedObjectCount2(-1) < 1 Then
    
        Exit Sub
        
    End If
    
    Dim swBodies As New VBA.Collection
    
    Dim i As Long
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
    

        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelSURFACEBODIES Then
        
            Dim swBody As SldWorks.Body2
            Set swBody = swSelMgr.GetSelectedObject6(i, -1)
            
            swBodies.Add swBody

        End If
        
    Next i
    
    Dim count As Long
    
    For i = swBodies.count To 1 Step -1
    
        If swBodies(i).Select(False, 1) Then
    
            Dim swFeature As SldWorks.Feature
            Set swFeature = swModel.FeatureManager.FeatureCutThicken(0.01, 2, 0, False, True, True)
            
            If Not swFeature Is Nothing Then
            
                swModel.FeatureManager.EditRollback swMoveRollbackBarTo_e.swMoveRollbackBarToBeforeFeature, swFeature.Name
            
                count = count + 1
            
            End If
    
        End If
        
    Next i
    
    swModel.FeatureManager.EditRollback swMoveRollbackBarTo_e.swMoveRollbackBarToEnd, ""
    
    swApp.SendMsgToUser2 "Successfully Added " & count & " Thicken Cut Features", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
    
End Sub

