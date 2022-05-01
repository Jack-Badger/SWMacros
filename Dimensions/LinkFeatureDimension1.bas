Attribute VB_Name = "LinkFeatureDimension1"
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
    
    Dim linkName As String
    linkName = InputBox("Please Enter Link Name", "Link First Dimension of Features")
    If linkName = "" Then Exit Sub
    
    Dim i As Long
    For i = 1 To count
    
        Dim swFeature As SldWorks.Feature
        Set swFeature = swSelMgr.GetSelectedObject6(Index:=i, Mark:=-1)
        
        Dim swDispDim As SldWorks.DisplayDimension
        Set swDispDim = swFeature.GetFirstDisplayDimension
        
        If Not swDispDim Is Nothing Then
        
            If swDispDim.IsLinked Then swDispDim.Unlink
                  
            Dim linkError As swLinkDimensionError_e
            linkError = swDispDim.SetLinkedText(linkName)
        
        End If
        
    Next i
    
End Sub
