Attribute VB_Name = "DeleteSketchRelations1"
' Delete Sketch Relations in selected sketches

Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    With swApp: Select Case True
       Case .ActiveDoc Is Nothing
       Case .ActiveDoc.GetType = swDocumentTypes_e.swDocASSEMBLY _
          , .ActiveDoc.GetType = swDocumentTypes_e.swDocPART
                Call ProcessModel(.ActiveDoc)
    End Select: End With
    
End Sub

Private Function ProcessModel(swModel As SldWorks.ModelDoc2)

        Dim swSelMgr As SldWorks.SelectionMgr
        Set swSelMgr = swModel.SelectionManager
        
        Dim swFeature As SldWorks.Feature
        
        Dim swSketch As SldWorks.Sketch
        
        Dim swSketchRelMgr As SldWorks.SketchRelationManager

        Dim i As Integer
        
        For i = 1 To swSelMgr.GetSelectedObjectCount2(Mark:=-1)
        
            If swSelMgr.GetSelectedObjectType3(index:=i, Mark:=-1) = swSelectType_e.swSelSKETCHES Then
            
                Set swFeature = swSelMgr.GetSelectedObject6(index:=i, Mark:=-1)
                
                Set swSketch = swFeature.GetSpecificFeature2
                
                Set swSketchRelMgr = swSketch.RelationManager
                
                swSketchRelMgr.DeleteAllRelations
                
            End If
            
        Next
        
        Call swModel.Extension.Rebuild(swRebuildOptions_e.swRebuildAll)
        
End Function



