Attribute VB_Name = "AddMidpointAndPerpendicul"
' yes its a mess but it works! :)

Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    Dim swSketchMgr As SldWorks.SketchManager
    Set swSketchMgr = swModel.SketchManager
    
    Dim swSketch As SldWorks.Sketch
    Set swSketch = swSketchMgr.ActiveSketch
    
    Dim swSketchSegment As SldWorks.SketchSegment
    
    Dim swSketchLine As SldWorks.SketchLine
    
    Dim swEdge As SldWorks.Edge
    
    
    Dim swSketchPoint As SldWorks.SketchPoint
    
    If swSelMgr.GetSelectedObjectCount2(-1) <> 2 Then
    
        Exit Sub
        
    End If
    
    
    
    
    Dim swSelectType As swSelectType_e
    
    swSelectType = swSelMgr.GetSelectedObjectType3(1, -1)
    
    Select Case swSelectType
    
            Case swSelectType_e.swSelSKETCHSEGS:
            Set swSketchSegment = swSelMgr.GetSelectedObject6(1, -1)
            Set swSketchLine = swSketchSegment
            Set swSketchPoint = swSketchLine.GetStartPoint2
    
        Case swSelectType_e.swSelSKETCHPOINTS:
           Set swSketchPoint = swSelMgr.GetSelectedObject6(1, -1)
        
        Case swSelectType_e.swSelREFEDGES:
            Set swEdge = swSelMgr.GetSelectedObject6(1, -1)
        
        Case Else:
            Exit Sub
        
    End Select
    
    
    
    
    swSelectType = swSelMgr.GetSelectedObjectType3(2, -1)
    
        Select Case swSelectType
    
        Case swSelectType_e.swSelSKETCHSEGS:
            Set swSketchSegment = swSelMgr.getselectobject6(2, -1)
            Set swSketchLine = swSketchSegment
            Set swSketchPoint = swSketchLine.GetStartPoint2
    
        Case swSelectType_e.swSelSKETCHPOINTS:
           Set swSketchPoint = swSelMgr.GetSelectedObject6(2, -1)
        
        Case swSelectType_e.swSelREFEDGES:
            Set swEdge = swSelMgr.GetSelectedObject6(2, -1)
        
        Case Else:
            Exit Sub
        
    End Select

    
    
    
    If swSketchPoint Is Nothing Then
    
        Exit Sub
        
    End If
    
    If swEdge Is Nothing Then
    
        Exit Sub
        
    End If
    
    
    
    Dim swEntity As SldWorks.Entity
    Set swEntity = swEdge

    
    swSketchPoint.Select4 False, Nothing
    swEntity.Select4 True, Nothing
    
    swModel.SketchAddConstraints "sgATMIDDLE"
    
    swSketchSegment.Select4 False, Nothing
    swEntity.Select4 True, Nothing
    
    swModel.SketchAddConstraints "sgPERPENDICULAR"
    
End Sub
