Attribute VB_Name = "FindReplaceInConfigNames1"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
    
        swApp.SendMsgToUser2 "No Model", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
        
        Exit Sub
        
    End If
    
    Dim vConfigNames As Variant
    
    Dim bResult As Boolean
    
    vConfigNames = GetAnySelectedConfigNames(swModel.SelectionManager, bResult)
    
    If Not bResult Then
    
        swApp.SendMsgToUser2 "No Configs preselected.  Apply to all?", swMessageBoxIcon_e.swMbQuestion, swMessageBoxBtn_e.swMbYesNo
        
        If swApp.SendMsgToUser2("No Configs preselected.  Apply to all?", swMbQuestion, swMbYesNo) = swMbHitYes Then
        
            vConfigNames = swModel.GetConfigurationNames
        
        Else
        
            Exit Sub
        
        End If
        
    End If
    
    Dim find As String
    find = InputBox("Enter string to find", "Replace chars in config names")
    
    Dim rep As String
    rep = InputBox("Enter replacement", "Replace chars in config names")

    If StrComp(find, rep, vbTextCompare) = 0 Then
        swApp.SendMsgToUser2 "No changes required", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
        Exit Sub
    End If

    Dim totalCount As Long
    totalCount = UBound(vConfigNames) + 1
    
    Dim swProgBar As SldWorks.UserProgressBar
    swApp.GetUserProgressBar swProgBar
    
    swProgBar.Start 0, totalCount - 1, "Find Replace in Config Names"
    
    Dim count As Long
    
    Dim i As Long
    For i = 0 To UBound(vConfigNames)
    
        Dim name As String
        name = vConfigNames(i)
        
        If InStr(1, name, find, vbTextCompare) > 0 Then
        
            Dim swConfig As SldWorks.Configuration
            Set swConfig = swModel.GetConfigurationByName(name)
            
            name = replace(name, find, rep)
            
            swConfig.name = name
                            
            If StrComp(swConfig.name, name, vbTextCompare) = 0 Then
            
                count = count + 1
                swConfig.Description = name
                
            End If
        
        End If
        
        swProgBar.UpdateProgress i
        
    Next i
    
    swProgBar.End
    
    Beep
    
    swApp.SendMsgToUser2 "Changed " & count & " of " & totalCount & " configuration names.", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
    
End Sub

Private Function GetAnySelectedConfigNames(swSelMgr As SldWorks.SelectionMgr, ByRef bResult As Boolean) As Variant

    bResult = False

    Dim totalSelectedObjectCount As Long
    totalSelectedObjectCount = swSelMgr.GetSelectedObjectCount2(-1)
    
    If totalSelectedObjectCount > 0 Then
    
        Dim vbCollection As VBA.Collection
        Set vbCollection = New VBA.Collection
    
        Dim i As Long
        For i = 1 To totalSelectedObjectCount
    
            If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelCONFIGURATIONS Then

                vbCollection.Add swSelMgr.GetSelectedObject6(i, -1).name
            
            End If
            
        Next i
        
        If vbCollection.count > 0 Then
        
            ReDim configNames(vbCollection.count - 1) As String
            
            For i = 0 To UBound(configNames)
            
                configNames(i) = vbCollection(i + 1)
            
            Next i
            
            bResult = True
            GetAnySelectedConfigNames = configNames
            
        End If
        
    End If

End Function
