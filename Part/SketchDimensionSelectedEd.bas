Attribute VB_Name = "SketchDimensionSelectedEd"
Option Explicit

Dim swApp As SldWorks.SldWorks

Dim swModel As SldWorks.ModelDoc2

Sub main()

    Set swApp = Application.SldWorks
        
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    'Dim swSelectType As swSelectType_e
    'swSelectType = swSelMgr.GetSelectedObjectType3(1, -1)
    'Debug.Print "select type:" & swSelectType
     
    Dim Edges() As SldWorks.Edge
    ReDim Edges(1 To swSelMgr.GetSelectedObjectCount2(-1))
    
    Dim i As Long
    For i = 1 To UBound(Edges)
    
        Set Edges(i) = swSelMgr.GetSelectedObject6(i, -1)
    
    Next i
    
    Dim swFeature As SldWorks.Feature
    
    For i = 1 To UBound(Edges)
    
        Set swFeature = ProcessEdge(Edges(i))
        swFeature.Name = "Edge_Measure" & i
    Next i
    
End Sub

Private Function ProcessEdge(swEdge As SldWorks.Edge) As SldWorks.Feature
    
    Dim vFaces As Variant
    vFaces = swEdge.GetTwoAdjacentFaces2

    Dim swFace As SldWorks.Face2
    Set swFace = vFaces(0)
    
    Dim swSurface As SldWorks.Surface
    Set swSurface = swFace.GetSurface
    
    If swSurface.IsPlane() Then
    
    Else
            Set swFace = vFaces(1)
            Set swSurface = swFace.GetSurface
            Debug.Assert swSurface.IsPlane() = True
    End If
    
    Dim swEntity As SldWorks.Entity
    Set swEntity = swFace
    
    swEntity.Select False
    
    Dim swSketchManager As SldWorks.SketchManager
    Set swSketchManager = swModel.SketchManager
        
    swSketchManager.InsertSketch True
    
    Dim swFeature As SldWorks.Feature
    Set swFeature = swSketchManager.ActiveSketch
    
    swSketchManager.AddToDB = True
    
        Dim swSketchLine As SldWorks.SketchLine
        Set swSketchLine = swSketchManager.CreateLine(0, 0, 0, 1, 1, 0)
        
        Dim swSketchSegment As SldWorks.SketchSegment
        Set swSketchSegment = swSketchLine
        swSketchSegment.Select4 False, Nothing
        Dim userPref As Boolean
        userPref = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swInputDimValOnCreate)
        
        swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swInputDimValOnCreate, False
        
        Dim swDispDim As SldWorks.DisplayDimension
        Set swDispDim = swModel.AddDimension2(0, 0.5, 0)
        
        Dim swDim As SldWorks.Dimension
        Set swDim = swDispDim.GetDimension
        
        swDim.DrivenState = swDimensionDrivenState_e.swDimensionDriven
        
        swDim.Name = "Length"
        
        swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swInputDimValOnCreate, userPref

        Dim swSketchPoint As SldWorks.SketchPoint
        Set swSketchPoint = swSketchLine.GetStartPoint2
        
        Dim swVertex As SldWorks.Vertex
        Set swVertex = swEdge.GetStartVertex
        Set swEntity = swVertex
        
        swSketchPoint.Select4 False, Nothing
        swEntity.Select4 True, Nothing
        
        swModel.SketchAddConstraints "sgCOINCIDENT"

        Set swSketchPoint = swSketchLine.GetEndPoint2
    
        Set swVertex = swEdge.GetEndVertex
        Set swEntity = swVertex
        
        swSketchPoint.Select4 False, Nothing
        swEntity.Select4 True, Nothing
        
        swModel.SketchAddConstraints "sgCOINCIDENT"
        
        swSketchManager.AddToDB = False
        swSketchManager.InsertSketch True
       
       Set ProcessEdge = swFeature

End Function
