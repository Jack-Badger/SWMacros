Attribute VB_Name = "LinkSelectedDimensions1"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then EndWithError "No Document!"

    Dim DimCollection As VBA.Collection
    Set DimCollection = CollectUniqueDimensions(swModel.SelectionManager)
        
    If DimCollection.count = 0 Then EndWithError "No Dimensions Selected!"
    
    Dim linkName As String
    
    linkName = InputBox("Please Enter Link Name", "Link Selected Dimensions")
    If linkName = "" Then Exit Sub ' User probly cancelled
      
    Dim i As Long
    For i = 1 To DimCollection.count
        
        Dim swDispDim As SldWorks.DisplayDimension
        Set swDispDim = DimCollection(i)
        
        If swDispDim.IsLinked Then swDispDim.Unlink
              
        Dim linkError As swLinkDimensionError_e
        linkError = swDispDim.SetLinkedText(linkName)
        
    Next i
    
End Sub

Private Sub EndWithError(message As String, Optional icon As swMessageBoxIcon_e = swMbInformation)

    swApp.SendMsgToUser2 message, icon, swMbOk
    
    End 'Nuke

End Sub

Private Function CollectUniqueDimensions(swSelMgr As SldWorks.SelectionMgr) As VBA.Collection

    Set CollectUniqueDimensions = New Collection

    Dim count As Long
    count = swSelMgr.GetSelectedObjectCount

    If count = 0 Then Exit Function
              
    Dim i As Long
    For i = 1 To count
    
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelDIMENSIONS Then
        
            Dim swDispDim As SldWorks.DisplayDimension
            Set swDispDim = swSelMgr.GetSelectedObject6(i, -1)
        
            On Error Resume Next
            CollectUniqueDimensions.Add Item:=swDispDim, Key:=swDispDim.GetDimension.FullName
            On Error GoTo 0
            
        End If
        
    Next i
      
End Function

