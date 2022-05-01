Attribute VB_Name = "DeconstructSketchedLines1"
Option Explicit

Dim swApp As SldWorks.SldWorks
Sub main()

Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    Dim swMasterFeature As SldWorks.Feature
    Set swMasterFeature = swSelMgr.GetSelectedObject6(1, -1)
    
    Dim swMasterSketch As SldWorks.Sketch
    Set swMasterSketch = swMasterFeature.GetSpecificFeature2
    
    Dim swSelectType As swSelectType_e
    Dim swEntity As SldWorks.Entity
    
    Set swEntity = swMasterSketch.GetReferenceEntity(swSelectType)
    
    'Dim swRefPlane As SldWorks.RefPlane
    
    Dim swSketchMgr As SldWorks.SketchManager
    Set swSketchMgr = swModel.SketchManager
    
    Dim SketchLines As VBA.Collection
    Set SketchLines = CollectSketchLines(swMasterSketch)
    
    Dim SketchBaseName As String
    SketchBaseName = VBA.InputBox("Base name for deconstructed sketches", "Deconstruct Sketched Lines", "Line")
    
    Dim swFeatMgr As SldWorks.FeatureManager
    Set swFeatMgr = swModel.FeatureManager
    
    swFeatMgr.EditRollback swMoveRollbackBarTo_e.swMoveRollbackBarToBeforeFeature, swMasterFeature.Name
    
    Dim i As Long
    For i = 1 To SketchLines.Count
    
        swEntity.Select4 False, Nothing
        
        With swSketchMgr
        
            .InsertSketch True
            .AddToDB = True
            
            Dim swFeature As SldWorks.Feature
            Set swFeature = .ActiveSketch
            
            swFeature.Name = SketchBaseName & "_" & i
            
            Dim swSketchLine As SldWorks.SketchLine
            Set swSketchLine = SketchLines(i)
            
            Dim StartPt As SldWorks.SketchPoint
            Set StartPt = swSketchLine.GetStartPoint2
            
            Dim EndPt As SldWorks.SketchPoint
            Set EndPt = swSketchLine.GetEndPoint2
        
            .CreateLine StartPt.X, StartPt.Y, StartPt.Z, EndPt.X, EndPt.Y, EndPt.Z
        
            .AddToDB = False
            .InsertSketch True
    
        End With
    
    Next i
    
    'TODO change to state when macro was called
    swFeatMgr.EditRollback swMoveRollbackBarTo_e.swMoveRollbackBarToEnd, ""
    
    swMasterFeature.Select2 False, -1
    
    swModel.EditSketch
    
    With swSketchMgr
    
    .AddToDB = True
    
        For i = 1 To SketchLines.Count
        
            Dim swPart As SldWorks.PartDoc
            ' TODO what about an assembly?
            Set swPart = swModel
            
            Set swFeature = swPart.FeatureByName(SketchBaseName & "_" & i)
        
            SelectInSketch swFeature.GetSpecificFeature2, swSketchSegments_e.swSketchLine, 1, False
            
            Dim swsketchsegment As SldWorks.SketchSegment
            
            Set swsketchsegment = SketchLines(i)
            
            swsketchsegment.Select4 True, Nothing
            
            swModel.SketchAddConstraints "sgCOLINEAR"
        
        Next i
        
    .AddToDB = False
        
    End With
    
    swModel.InsertSketch2 True
    

    
End Sub

Private Function SelectInSketch(swSketch As SldWorks.Sketch, swSketchSegmentType As swSketchSegments_e, num As Long, Append As Boolean) As Boolean

    Dim vSketchSegments As Variant
    vSketchSegments = swSketch.GetSketchSegments
    
    Dim bFound As Boolean
    
    Dim lFound As Long
    
    Dim swsketchsegment As SldWorks.SketchSegment
    
    Dim i As Long
    i = -1 'not sure bout dis
    
    Do While Not SelectInSketch And i < UBound(vSketchSegments)
    
        i = i + 1
    
        Set swsketchsegment = vSketchSegments(i)
        
        If swsketchsegment.GetType = swSketchSegmentType Then
        
            lFound = lFound + 1
            
            If lFound = num Then
            
                bFound = True
                    
            End If
            
        End If
        
    Loop
    
    If bFound Then
    
        SelectInSketch = swsketchsegment.Select4(Append, Nothing)
    
    End If
    
End Function


Private Function CollectSketchLines(swSketch As SldWorks.Sketch) As VBA.Collection

    Set CollectSketchLines = New VBA.Collection
    
    Dim vSketchSegs As Variant
    vSketchSegs = swSketch.GetSketchSegments
      
    Dim i As Long
    For i = 0 To UBound(vSketchSegs)

        Dim swsketchsegment As SldWorks.SketchSegment
        Set swsketchsegment = vSketchSegs(i)
        
        If swsketchsegment.GetType = swSketchSegments_e.swSketchLine Then
        
            Dim swSketchLine As SldWorks.SketchLine
            Set swSketchLine = swsketchsegment
            CollectSketchLines.Add swSketchLine
        
        End If
        
    Next i

End Function
