Attribute VB_Name = "ReadOnlyFirstFeatureDim1"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    Dim count As Long
    count = swSelMgr.GetSelectedObjectCount2(Mark:=-1)
    
    If count = 0 Then Exit Sub
        
    Dim i As Long
    For i = 1 To count
    
        Dim swFeature As SldWorks.Feature
        Set swFeature = swSelMgr.GetSelectedObject6(Index:=i, Mark:=-1)
        
        Dim swDispDim As SldWorks.DisplayDimension
        Set swDispDim = swFeature.GetFirstDisplayDimension
        
        If Not swDispDim Is Nothing Then
        
            Dim swDimension As SldWorks.Dimension
            Set swDimension = swDispDim.GetDimension
            
            swDimension.ReadOnly = True
            
        End If
        
    Next i
    
End Sub

