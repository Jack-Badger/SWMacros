Attribute VB_Name = "SelectComponents1"
' Make some selections in an assembly
' Run the macro
' The selections are modified to the component level
' Optionally you can save the selections as a selection set
' A Possible Enhancement is to allow for Top level selection
' ie - select the subassembly (let me know if this is something useful)
' rob 07May21

'**********************
'Copyright(C) 2020 Xarial Pty Limited
'Reference: https://www.codestack.net/solidworks-api/document/assembly/components/move-to-folder/
'License: https://www.codestack.net/license/
'**********************
#If VBA7 Then
     Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
     Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        Exit Sub
    End If
    
    If swModel.GetType <> swDocumentTypes_e.swDocASSEMBLY Then
        Exit Sub
    End If
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    If swSelMgr.GetSelectedObjectCount2(-1) < 1 Then
        Exit Sub
    End If
    
    Dim vbCollection As VBA.Collection
    Set vbCollection = New VBA.Collection
    
    Dim i As Long
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
    
        Dim swComponent As SldWorks.Component2
        Set swComponent = swSelMgr.GetSelectedObjectsComponent4(i, -1)
        
        If Not swComponent Is Nothing Then
        
            On Error Resume Next
            
                vbCollection.Add swComponent, swComponent.Name2
            
            On Error GoTo 0
        
        End If
        
    Next i
    
    swModel.ClearSelection2 True
    
    Dim selectedCount As Long
    
    If vbCollection.Count > 0 Then
    
        For i = 1 To vbCollection.Count
        
            Set swComponent = vbCollection(i)
            
            If swComponent.Select4(True, Nothing, False) Then
            
                selectedCount = selectedCount + 1
                
            End If
            
        Next i
        
    End If
    
    Select Case swApp.SendMsgToUser2(selectedCount & "Components Selected." & vbCr _
                            & "Yes to create a Selection Set?" & vbCr _
                            & "No to create a Folder" & vbCr _
                            & "Cancel to Exit" _
                            , swMessageBoxIcon_e.swMbQuestion _
                            , swMessageBoxBtn_e.swMbYesNoCancel)
   
    
        Case swMessageBoxResult_e.swMbHitYes:
        
            Dim swSelectionSet As SldWorks.SelectionSet
    
            Dim status As Long
    
            Set swSelectionSet = swModel.Extension.SaveSelection(status)
            
        Case swMessageBoxResult_e.swMbHitNo:
        
            '**********************
            'Copyright(C) 2020 Xarial Pty Limited
            'Reference: https://www.codestack.net/solidworks-api/document/assembly/components/move-to-folder/
            'License: https://www.codestack.net/license/
            '**********************
            
                Const WM_COMMAND As Long = &H111
                Const CMD_ADD_TO_NEW_FOLDER As Long = 37922
                
                Dim swFrame As SldWorks.Frame
                    
                Set swFrame = swApp.Frame
                    
                SendMessage swFrame.GetHWnd(), WM_COMMAND, CMD_ADD_TO_NEW_FOLDER, 0
                    
    End Select
    
End Sub
