' Solidworks 2018
'
' Get Write Access For Dirty Files
'
' In an assembly document
'
'   Get Every Unsuppressed Component
'      Check Save Flag
'      If Dirty Attempt Gain Write Access
'
'   Display Affected Files To User
'
' This is useful with the setting to open referenced files with read only access.
' After running this macro you can see the files that are affected and save.
'
' rob@jackbadger.co.uk
' 5 7 23

Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then Exit Sub
    
    If swModel.GetType <> swDocumentTypes_e.swDocASSEMBLY Then Exit Sub
    
    Dim swAssem As SldWorks.AssemblyDoc
    Set swAssem = swModel
    
    Dim vComponents As Variant
    vComponents = swAssem.GetComponents(TopLevelOnly:=False)
    
    Dim success As New VBA.Collection
    Dim fail As New VBA.Collection
    
    Dim i As Long
    For i = 0 To UBound(vComponents)
    
        Dim swComponent As SldWorks.Component2
        Set swComponent = vComponents(i)
        
        If Not swComponent.GetSuppression = swComponentSuppressed Then
        
            Dim swComponentModel As SldWorks.ModelDoc2
            Set swComponentModel = swComponent.GetModelDoc2
            
            If swComponentModel.IsOpenedReadOnly And swComponentModel.GetSaveFlag Then
            
                If swComponentModel.SetReadOnlyState(False) Then
                
                    success.Add swComponentModel.GetTitle
                    
                Else
                
                    fail.Add swComponentModel.GetTitle
                
                End If
                
            End If
            
        End If
    
    Next i
    
    If success.Count > 0 Or fail.Count > 0 Then
    
        Dim message As String
        message = ""
        
        If success.Count > 0 Then
            message = "Write Access set for" & vbNewLine
        
            For i = 1 To success.Count
        
                message = message & success(i) & vbNewLine
        
            Next i
            
        End If
        
        If fail.Count > 0 Then
        
            message = message & "Not able to get write access" & vbNewLine
            
            For i = 1 To fail.Count
            
                message = message & fail(i) & vbNewLine
            
            Next i
        
        End If
    
        swApp.SendMsgToUser2 message, swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
    
    End If

End Sub
