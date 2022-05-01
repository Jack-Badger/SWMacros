Attribute VB_Name = "EmbedMacroFIle1"
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = swModel
    
    Dim sFeatureName As String
    
    sFeatureName = InputBox("Please enter name of macro feature you would like to embed." & vbCr & "Note. You must run this on the computer the macro resides.", "Embed Macro Feature")
    
    Dim swFeature As SldWorks.Feature
    Set swFeature = swPart.FeatureByName(sFeatureName)
    
    If swFeature Is Nothing Then
    
        swApp.SendMsgToUser2 "Did you spell that right?", swMessageBoxIcon_e.swMbWarning, swMessageBoxBtn_e.swMbOk
        Exit Sub
    
    End If
    
    
    Dim swMFD As SldWorks.MacroFeatureData
    
    Set swMFD = swFeature.GetDefinition
    
    swMFD.EmbedMacroFIle
    
    Dim bResult As Boolean
    
    bResult = swFeature.ModifyDefinition(swMFD, swModel, Nothing)
    swApp.SendMsgToUser2 "Success: " & bResult, swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
End Sub
