Attribute VB_Name = "FlipMarkedForDrawing1"
' SOLIDWORKS
' This macro will flip (toggle) 'Marked for Drawing' on all your dimensions.
' Works in a part file or an assembly but doesn't go down into your components
' Warning - I'm a joiner not a programmer, so use at your own risk!
' https://forum.solidworks.com/people/1-2SZI4TC


Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr
Dim swFeature As SldWorks.Feature
Dim swDisplayDimension As SldWorks.DisplayDimension
Dim swDimension As SldWorks.Dimension
Dim colFeaturesWithDims As Collection
Dim colDispDims As Collection

Sub main()

    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    Set colFeaturesWithDims = New Collection
    Set swFeature = swModel.FirstFeature
    
    Do While Not swFeature Is Nothing
        If HasDisplayDimensions(swFeature) Then
            On Error Resume Next 'we don't care if already in the collection
                Call colFeaturesWithDims.Add(Item:=swFeature, key:=swFeature.Name)
            On Error GoTo 0
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop

    Dim i As Integer
    For i = 1 To colFeaturesWithDims.Count
        Debug.Print colFeaturesWithDims(i).Name
    Next
    
    Set colDispDims = New Collection
    For i = 1 To colFeaturesWithDims.Count
        Set swFeature = colFeaturesWithDims(i)
        Set swDisplayDimension = swFeature.GetFirstDisplayDimension

        Do While Not swDisplayDimension Is Nothing
            On Error Resume Next 'we don't care if already in the collection
                Call colDispDims.Add(Item:=swDisplayDimension, key:=swDisplayDimension.GetDimension.FullName)
            On Error GoTo 0
            Set swDisplayDimension = swFeature.GetNextDisplayDimension(swDisplayDimension)
        Loop
    Next
   
   Dim command As String
   
   command = InputBox("Mark For Drawing: 1" & vbCr & "Unmark for Drawing: 2" & vbCr & "Toggle Mark for Drawing: 3", "Mark for Drawing Option", "3")
   
   
    For i = 1 To colDispDims.Count
        Set swDisplayDimension = colDispDims(i)
        With swDisplayDimension
            Debug.Print .GetNameForSelection, .MarkedForDrawing
            Select Case command
                                Case "1": .MarkedForDrawing = True
                                Case "2": .MarkedForDrawing = False
                                Case "3": .MarkedForDrawing = IIf(.MarkedForDrawing, False, True)
            End Select
        End With
    Next
    
    Call swApp.SendMsgToUser2(colDispDims.Count & " dimension" & IIf(colDispDims.Count = 1, "", "s") & " " & CommandDescription(command), swMbInformation, swMbOk)
   
End Sub

Private Function HasDisplayDimensions(swFeature As SldWorks.Feature) As Boolean
    HasDisplayDimensions = Not swFeature.GetFirstDisplayDimension Is Nothing
End Function



Private Function CommandDescription(c As String) As String
               Select Case c
                                Case "1": CommandDescription = "Marked"
                                Case "2": CommandDescription = "UnMarked"
                                Case "3": CommandDescription = "Flipped"
            End Select
 
End Function
