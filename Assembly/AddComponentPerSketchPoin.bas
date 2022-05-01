Attribute VB_Name = "AddComponentPerSketchPoin"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swAssModel As SldWorks.ModelDoc2
    
    Set swAssModel = swApp.ActiveDoc
    
    If swAssModel Is Nothing Then
    
        EndWithError "Yo Stupid - No Document. ;)"
    
    End If
    
    If swAssModel.GetType <> swDocumentTypes_e.swDocASSEMBLY Then
    
        EndWithError "You need to open an assembly document."

    End If
  
    If swAssModel.GetSaveFlag Then
    
        Select Case AskQuestion("The file is unsaved." & vbCr & "Do you want to save before we proceed?")
        
            Case swMessageBoxResult_e.swMbHitYes:
                
                Dim SaveOptions As swSaveAsOptions_e
                Dim SaveErrors As swFileSaveError_e
                Dim SaveWarnings As swFileSaveWarning_e
                            
                If Not swAssModel.Save3(SaveOptions, SaveErrors, SaveWarnings) Then
                
                    EndWithError "Something went wrong with Save." & vbCr & "Error:" & SaveErrors & "Warning:" & SaveWarnings
                    
                End If
            
            Case swMessageBoxResult_e.swMbHitCancel:
                Exit Sub
            
        End Select
        
    End If
    
    
    Dim swAssem As SldWorks.AssemblyDoc
    Set swAssem = swAssModel
     
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swAssModel.SelectionManager
    
    Dim PlaneCollection As VBA.Collection
    Set PlaneCollection = New Collection
    
    Dim SketchCollection As VBA.Collection
    Set SketchCollection = New Collection
    
    If swSelMgr.GetSelectedObjectCount2(-1) < 1 Then
        EndWithError "Nothing Selected"
    End If
    
    Dim swFeature As SldWorks.Feature
    Dim NameForSelection As String
    
    Dim i As Long
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
    
        Dim selObjType As swSelectType_e
        selObjType = swSelMgr.GetSelectedObjectType3(i, -1)
    
        Select Case selObjType
        
            Case swSelectType_e.swSelDATUMPLANES:
            
                Set swFeature = swSelMgr.GetSelectedObject6(i, -1)
                
                NameForSelection = swFeature.GetNameForSelection(swSelDATUMPLANES)
                
                If AddToCollection(swFeature, PlaneCollection, NameForSelection) Then
                
                    Debug.Print "Added Plane:" & swFeature.name
                    
                End If
                
                Set swFeature = Nothing

            Case swSelectType_e.swSelSKETCHES:
            
                Set swFeature = swSelMgr.GetSelectedObject6(i, -1)
                
                NameForSelection = swFeature.GetNameForSelection(swSelSKETCHES)
                
                If AddToCollection(swFeature, SketchCollection, NameForSelection) Then
                
                    Debug.Print "Added Sketch:" & swFeature.name
                    
                End If
                
                Set swFeature = Nothing
                
            Case Else: 'Ignore
        
        End Select
        
    Next i
    
    If PlaneCollection.count <> 2 Then
    
        EndWithError "You need to select 2 planes."
    
    End If
    
    If SketchCollection.count < 1 Then
    
        EndWithError "No Sketches selected"
    
    End If
    
    Set SketchCollection = ProcessSketchCollection(SketchCollection)
    
    If SketchCollection.count < 1 Then
    
        EndWithError "Sketches have no points"
    
    End If
    
    
    If Not TryProcessPlaneCollection(PlaneCollection) Then
    
        EndWithError "You need to select an assembly plane and a plane from the component to be added."
 
    End If
    
    Dim swAssPlane As SldWorks.Feature
    Set swAssPlane = PlaneCollection("AP")
    
    Dim swCompPlane As SldWorks.Feature
    Set swCompPlane = PlaneCollection("CP")
    
    Dim swComponent As SldWorks.Component2
    Set swComponent = PlaneCollection("C")
      
    Debug.Print swAssPlane.name
    
    Debug.Print swCompPlane.name
    
    Debug.Print swComponent.Name2
        
    Dim refConfig As String
    refConfig = swComponent.ReferencedConfiguration
    
    Dim path As String
    path = swComponent.GetPathName
    
    Dim swSelData As SldWorks.SelectData
    Set swSelData = swSelMgr.CreateSelectData
    swSelData.Mark = 1
    
    Updates swAssModel, False
    

    
    Dim total As Long
    total = SketchCollection("Count")
    SketchCollection.Remove "Count" ' VBA sucks... or I do lol
    'Dim bs As Boolean
    
    'Dim progBar As SldWorks.UserProgressBar
    'bs = swApp.GetUserProgressBar(progBar)
    'bs = progBar.Start(1, total, "Adding Components")
    
    On Error GoTo EH
    
    Dim j As Long
    Dim added As Long
    
    For j = 1 To SketchCollection.count
    
        Set swFeature = SketchCollection(j)
        
        Dim sketchNameForSelection As String
        sketchNameForSelection = swFeature.GetNameForSelection(swSelSKETCHES)
        
        Dim swSketch As SldWorks.Sketch
        Set swSketch = swFeature.GetSpecificFeature2
        
        Dim vSketchPoints As Variant
        vSketchPoints = swSketch.GetSketchPoints2
        
        Dim k As Long
        For k = 0 To UBound(vSketchPoints)
            
            Dim swSketchPoint As SldWorks.SketchPoint
            
            Set swSketchPoint = vSketchPoints(k)
            
            If swSketchPoint.Type = swSketchPointType_e.swSketchPointType_User Then
            
                Dim newComponent As SldWorks.Component2
                
                Set newComponent = AddComponent(swAssem, path, refConfig)
                
                swAssPlane.Select2 False, 1
                
                Dim swModel As SldWorks.ModelDoc2
                Set swModel = newComponent.GetModelDoc2
            
                Set swFeature = GetFeatureByName(swModel, swCompPlane.name)
                Set swFeature = newComponent.GetCorresponding(swFeature)
                swFeature.Select2 True, 1
                
                Dim swMate As SldWorks.Mate2
                Set swMate = AddMate(swAssem, swMatePARALLEL)
                
                Set swFeature = GetFeatureByName(swModel, "Origin")
                Set swFeature = newComponent.GetCorresponding(swFeature)
                Dim swOriginSketch  As SldWorks.Sketch
                Set swOriginSketch = swFeature.GetSpecificFeature2
                
                Dim swOriginSkPoint As SldWorks.SketchPoint
                Set swOriginSkPoint = swOriginSketch.GetSketchPoints2(0)
                           
                swOriginSkPoint.Select4 False, swSelData
                swSketchPoint.Select4 True, swSelData
                
                Set swMate = AddMate(swAssem)
                added = added + 1
                'progBar.UpdateProgress added
            End If
            
        Next k
    
    Next j
    
    'progBar.End
    
    Updates swAssModel, True
    
    swApp.SendMsgToUser2 "Added " & added & " components.", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
    End
    
EH:
    Updates swAssModel, True
    
    EndWithError "Summit Bad went down!"
    
End Sub

Private Function GetFeatureByName(swModel As SldWorks.ModelDoc2, name As String) As SldWorks.Feature
    
    If swModel.GetType = swDocumentTypes_e.swDocPART Then
    
        Dim swPart As SldWorks.PartDoc
        Set swPart = swModel
        
        Set GetFeatureByName = swPart.FeatureByName(name)
    
    Else
    
        Dim swAssem As SldWorks.AssemblyDoc
        Set swAssem = swModel
        
        Set GetFeatureByName = swAssem.FeatureByName(name)
    
    End If
    
End Function


Private Function AddMate(swAssem As SldWorks.AssemblyDoc, Optional MateType As swMateType_e = swMateCOINCIDENT, Optional Align As swMateAlign_e = swMateAlignCLOSEST, Optional Flip As Boolean = False) As SldWorks.Mate2
    
    Dim swMate As SldWorks.Mate2
    
    Dim error As swAddMateError_e
        
    Set swMate = swAssem.AddMate5(MateType, Align, Flip, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, False, False, 0, error)
    
    If swMate Is Nothing Then
        Debug.Print "Mate Error:" & error
    End If
    
    Set AddMate = swMate
    
End Function

Private Function AddComponent(swAssem As SldWorks.AssemblyDoc, compPath As String, configname As String) As SldWorks.Component2
    Set AddComponent = swAssem.AddComponent4(compPath, configname, 0#, 0#, 0#)
End Function

Private Function AskQuestion(message As String, Optional icon As swMessageBoxIcon_e = swMbQuestion, Optional button As swMessageBoxBtn_e = swMbYesNoCancel) As swMessageBoxResult_e

    AskQuestion = swApp.SendMsgToUser2(message, icon, button)

End Function

Private Sub EndWithError(message As String, Optional icon As swMessageBoxIcon_e = swMbInformation, Optional button As swMessageBoxBtn_e = swMbOk)

        swApp.SendMsgToUser2 message, button, icon
        End ' Nuke Button
        
End Sub

Private Function AddToCollection(ob As Object, col As Collection, Optional key As String = "") As Boolean
    
    Dim count As Long
    count = col.count
    
    On Error Resume Next
    
        If key = "" Then
        
            col.Add ob
        
        Else
        
            col.Add ob, key
        
        End If
        
    On Error GoTo 0
    
    AddToCollection = count <> col.count
    
End Function

Private Function ProcessSketchCollection(col As VBA.Collection) As VBA.Collection


    Dim swFeature As SldWorks.Feature
    Dim swSketch As SldWorks.Sketch
    
    Set ProcessSketchCollection = New VBA.Collection
    
    If col.count = 0 Then
        
        Exit Function
        
    End If
    
    Dim totalPointCount As Long
    Dim pointCount As Long
    
    Dim i As Long
    For i = 1 To col.count
    
        Set swFeature = col(i)
        
        Set swSketch = swFeature.GetSpecificFeature2
        
        pointCount = swSketch.GetUserPointsCount
        
        If pointCount > 0 Then
        
            ProcessSketchCollection.Add swFeature
            totalPointCount = totalPointCount + pointCount
        
        End If
        
    Next
    
    ProcessSketchCollection.Add totalPointCount, "Count"

End Function

Private Function TryProcessPlaneCollection(coll As VBA.Collection _
                                         , Optional assPlaneKey As String = "AP" _
                                         , Optional compPlaneKey As String = "CP" _
                                         , Optional compKey As String = "C" _
                                         ) As Boolean

    Dim swComponent As SldWorks.Component2
    
    Dim swEntity1 As SldWorks.Entity
    Dim swEntity2 As SldWorks.Entity
    
    Set swEntity1 = coll(1)
    Set swEntity2 = coll(2)
    
    coll.Remove 2
    coll.Remove 1
    
    
    Set swComponent = swEntity1.GetComponent

    If swComponent Is Nothing Then
    
        coll.Add swEntity1, assPlaneKey
        
    Else
    
        coll.Add swEntity1, compPlaneKey
        coll.Add swComponent, compKey
    
    End If

    
    Set swComponent = swEntity2.GetComponent

    On Error GoTo ProcessError

    If swComponent Is Nothing Then
    
        coll.Add swEntity2, assPlaneKey

    Else
        
        coll.Add swEntity2, compPlaneKey
        coll.Add swComponent, compKey
        
    End If
    
    On Error GoTo 0
    
    TryProcessPlaneCollection = True
    
    Exit Function

ProcessError:
    
    TryProcessPlaneCollection = False
    On Error GoTo 0

End Function

Private Sub Updates(swModel As SldWorks.ModelDoc2, update As Boolean)

    'Dim swView As SldWorks.ModelView
    'Set swView = swModel.ActiveView
    'swView.EnableGraphicsUpdate = update
    swModel.FeatureManager.EnableFeatureTree = update
    swModel.FeatureManager.EnableFeatureTreeWindow = update
    
End Sub

