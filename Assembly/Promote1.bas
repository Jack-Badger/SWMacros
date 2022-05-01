Attribute VB_Name = "Promote1"
Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swAssem As SldWorks.AssemblyDoc

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    Dim userInput As String
    
    userInput = VBA.InputBox("1 - Hide" & vbCr & "2 - Promote" & vbCr & "3 - Show" & vbCr & "4 - Cancel", "Child Component BOM Option (All Configs)")
    
    Dim swChildCompOption As swChildComponentInBOMOption_e
    
    Select Case userInput
    
        Case "1": swChildCompOption = swChildComponent_Hide
        Case "2": swChildCompOption = swChildComponent_Promote
        Case "3": swChildCompOption = swChildComponent_Show
        Case Else
            Exit Sub
            
    End Select
    
    Dim vConfigs As Variant
    
    vConfigs = swModel.GetConfigurationNames
    
    Dim swConfig As SldWorks.Configuration
       
    Dim i As Long
    
    For i = 0 To UBound(vConfigs)
    
        Set swConfig = swModel.GetConfigurationByName(vConfigs(i))
        
        swConfig.ChildComponentDisplayInBOM = swChildCompOption
    
    Next i
    
    swApp.SendMsgToUser2 i & " configs processed", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
    
End Sub
