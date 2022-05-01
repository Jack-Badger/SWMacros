Attribute VB_Name = "SelectedDimensionsReadOnl"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim SelectedDimensions As VBA.Collection
    Set SelectedDimensions = CollectSelectedDimensions(swModel.SelectionManager)

    If SelectedDimensions.Count < 1 Then Exit Sub
    
    Dim i As Long
    For i = 1 To SelectedDimensions.Count
    
        Dim swDimension As SldWorks.Dimension
        Set swDimension = SelectedDimensions(i)
        swDimension.ReadOnly = True
        
    Next i
    
End Sub


Private Function CollectSelectedDimensions(swSelMgr As SldWorks.SelectionMgr) As VBA.Collection

    Set CollectSelectedDimensions = New VBA.Collection
    
    Dim selCount As Long
    selCount = swSelMgr.GetSelectedObjectCount2(-1)
    
    If selCount = 0 Then
    
        Exit Function
    
    End If
    
    Dim i As Long
    For i = 1 To selCount
    
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelDIMENSIONS Then
        
            Dim swDispDim As SldWorks.DisplayDimension
            Set swDispDim = swSelMgr.GetSelectedObject6(i, -1)
            
            Dim swDimension As SldWorks.Dimension
            Set swDimension = swDispDim.GetDimension
            
            Dim key As String
            key = swDimension.FullName
            
            On Error Resume Next
            
                CollectSelectedDimensions.Add swDimension, key

            On Error GoTo 0
            
        End If
        
    Next i

End Function
