Attribute VB_Name = "DualAssemCut1"
Dim swApp As SldWorks.SldWorks

Dim swModel As SldWorks.ModelDoc2

Dim swAssem As SldWorks.AssemblyDoc

Dim swSketchFeature As SldWorks.Feature

Dim swCompFeature1 As SldWorks.Component2
Dim swCompFeature2 As SldWorks.Component2

Dim swSelMgr As SldWorks.SelectionMgr

Dim swSelData As SelectData

    Dim selCount As Long

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swAssem = swModel
    
    Set swSelMgr = swModel.SelectionManager
    
    If ValidateSelections() Then
    
        Dim SketchName As String
        baseName = Replace(Expression:=swSketchFeature.Name, Find:="_", Replace:="-", Count:=-1)
        
        Set swSelData = swSelMgr.CreateSelectData
        swSelData.Mark = 128
    
        Dim swCutFeature1 As SldWorks.Feature
        Set swCutFeature1 = InsertFeature(swCompFeature1, False)
        
        If swCutFeature1 Is Nothing Then Exit Sub
        
        swCutFeature1.Name = GetComponentName(swCompFeature1) & "-" & baseName
        
        If selCount = 3 Then
        
            Dim swCutFeature2 As SldWorks.Feature
            Set swCutFeature2 = InsertFeature(swCompFeature2, True)
            
            If swCutFeature2 Is Nothing Then Exit Sub
            
            swCutFeature2.Name = GetComponentName(swCompFeature2) & "-" & baseName
            
            swCutFeature1.Select2 append:=False, Mark:=0
            swCutFeature2.Select2 append:=True, Mark:=0
          
            Dim swFeatureTreeFolder As Feature
            
            Set swFeatureTreeFolder = swModel.FeatureManager.InsertFeatureTreeFolder2(swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing)
            
            If Not swFeatureTreeFolder Is Nothing Then
                swFeatureTreeFolder.Name = baseName
            End If
        
        End If
        
        
        
        Dim UserResponse As swMessageBoxResult_e
        UserResponse = swApp.SendMsgToUser2("Dual Assembly Cut" & vbCr & "Looks like that worked?" & vbCr & "Yes to accept." & vbCr & "No to flip side to cut?" & vbCr & "Cancel to delete", swMessageBoxIcon_e.swMbQuestion, swMessageBoxBtn_e.swMbYesNoCancel)
        
        Select Case UserResponse
        
            Case swMessageBoxResult_e.swMbHitNo
            
                FlipSideToCut swCutFeature1
                If selCount = 3 Then FlipSideToCut swCutFeature2
            
            Case swMessageBoxResult_e.swMbHitYes
            
            Case swMessageBoxResult_e.swMbHitCancel
            
                swCutFeature1.Select2 append:=False, Mark:=0
                If selCount = 3 Then swCutFeature2.Select2 append:=True, Mark:=0
                swModel.Extension.DeleteSelection2 0
        
        End Select
        
    Else
    
        Call swApp.SendMsgToUser2("Please Select a Sketch and 1 or 2 Components.", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk)
    
    End If

End Sub

Private Function GetComponentName(swComponent As SldWorks.Component2) As String

    If (swComponent.ComponentReference <> "") Then
        GetComponentName = swComponent.ComponentReference
    Else
        GetComponentName = swComponent.ReferencedConfiguration
    End If
    
End Function



Private Sub FlipSideToCut(swFeature As SldWorks.Feature)

    Dim swExtrudeFeatureData As SldWorks.ExtrudeFeatureData2
    Set swExtrudeFeatureData = swFeature.GetDefinition
    
    Call swExtrudeFeatureData.AccessSelections(swModel, Nothing)
    
    swExtrudeFeatureData.FlipSideToCut = IIf(swExtrudeFeatureData.FlipSideToCut, False, True)
    
    Call swFeature.ModifyDefinition(swExtrudeFeatureData, swModel, Nothing)

End Sub


Private Function InsertFeature(swComponent As SldWorks.Component2, FlipSide As Boolean) As SldWorks.Feature

    Call swSketchFeature.Select2(append:=False, Mark:=0)
    
    Call swComponent.Select4(append:=True, Data:=swSelData, showpopup:=False)

    Set InsertFeature = swModel.FeatureManager.FeatureCut4 _
        (Sd:=True, Flip:=FlipSide, Dir:=True _
        , T1:=swEndConditions_e.swEndCondThroughAll, T2:=swEndConditions_e.swEndCondThroughAll _
        , D1:=1#, D2:=1#, Dchk1:=False, Dchk2:=False, Ddir1:=False, Ddir2:=False, Dang1:=0#, Dang2:=0# _
        , OffsetReverse1:=False, OffsetReverse2:=False, TranslateSurface1:=False, TranslateSurface2:=False _
        , NormalCut:=False, UseFeatScope:=True, UseAutoSelect:=False, AssemblyFeatureScope:=True, AutoSelectComponents:=False _
        , PropagateFeatureToParts:=False, T0:=swStartConditions_e.swStartSketchPlane, StartOffset:=0#, FlipStartOffset:=False, OptimizeGeometry:=False)

End Function

Private Function ValidateSelections() As Boolean

    ValidateSelections = False
    
    Set swSketchFeature = Nothing
    Set swCompFeature1 = Nothing
    Set swCompFeature2 = Nothing
    

    selCount = swSelMgr.GetSelectedObjectCount2(-1)

    If selCount = 2 Or selCount = 3 Then
      
        Dim i As Long
        
        For i = 1 To selCount
        
            If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelSKETCHES Then
            
                Set swSketchFeature = swSelMgr.GetSelectedObject6(i, -1)
                
            End If
            
            If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelCOMPONENTS Then
            
                If swCompFeature1 Is Nothing Then
                
                    Set swCompFeature1 = swSelMgr.GetSelectedObject6(i, -1)
                    
                Else
                
                    Set swCompFeature2 = swSelMgr.GetSelectedObject6(i, -1)
                    
                End If
                
            End If
                
        Next i
        
        If swSketchFeature Is Nothing Then Exit Function
                    
        If swCompFeature1 Is Nothing Then Exit Function
        
        If selCount = 3 And swCompFeature2 Is Nothing Then Exit Function
        
        ValidateSelections = True
              
    End If

End Function

