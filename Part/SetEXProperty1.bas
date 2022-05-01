Attribute VB_Name = "SetEXProperty1"
' SetEXProperty.swp
' Macro examines Length Width and Thickness Configuration Specific Properties and if present calculates the EX stock thickness
' This is a very primitive macro and requires more work!

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swConfig As SldWorks.Configuration
Dim swPropMgr As SldWorks.CustomPropertyManager

Dim PropAddedCount As Long

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then Exit Sub
    
    If swModel.GetType <> swDocumentTypes_e.swDocPART Then
    
        Call swApp.SendMsgToUser2("This macro only works for part documents." & vbCr & "Pester Rob for this enhancement request!", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk)
    
    End If
    
    Updates False
            
    Set swConfig = swModel.GetActiveConfiguration
    
    Dim ActiveConfigName As String
    ActiveConfigName = swConfig.Name
    
    Dim vConfigNames As Variant
    
    vConfigNames = swModel.GetConfigurationNames
    
    Dim ConfigName As String
    
    Dim i As Long
    
    For i = 0 To UBound(vConfigNames)
    
        ConfigName = vConfigNames(i)
        
        Call swModel.ShowConfiguration2(ConfigName)
        
        Set swConfig = swModel.GetActiveConfiguration
        
        Call DoShiz(swConfig.CustomPropertyManager)
        
        Debug.Print ConfigName
    
    Next i
    
    Call swModel.ShowConfiguration2(ActiveConfigName)
    
    Updates True
    
    Call swApp.SendMsgToUser2("Added " & PropAddedCount & " Properties", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk)



End Sub


Private Sub DoShiz(swPropMgr As SldWorks.CustomPropertyManager)

    Dim ValOut As String
    Dim ResolvedValOut As String
    Dim WasResolved As Boolean
    Dim LinkToProperty As Boolean
    
    Dim swResult As swCustomInfoGetResult_e

    'swResult = swPropMgr.Get6("Length", False, ValOut, ResolvedValOut, WasResolved, LinkToProperty)

    'If swResult = swCustomInfoGetResult_NotPresent Then
    
        'Exit Sub
        
    'End If
                                                      
    'Dim Length As Double
    'Length = ResolvedValOut
    
    'swResult = swPropMgr.Get6("Width", False, ValOut, ResolvedValOut, WasResolved, LinkToProperty)

    'If swResult = swCustomInfoGetResult_NotPresent Then
    
        'Exit Sub
        
    'End If
                                                          
    'Dim Width As Double
    'Width = ResolvedValOut


    swResult = swPropMgr.Get6("Thickness", False, ValOut, ResolvedValOut, WasResolved, LinkToProperty)

    If swResult = swCustomInfoGetResult_NotPresent Then
    
        Exit Sub
        
    End If
                                                      
    Dim Thickness As Double
    Thickness = ResolvedValOut

    Dim EXValue As String
    
    Select Case True
    
        Case Thickness < 10#:
            EXValue = "1/2"""
            
        Case Thickness < 16#:
            EXValue = "3/4"""

        Case Thickness < 25#:
            EXValue = "1"""
            
        Case Thickness < 38#:
            EXValue = "1-1/2"""
            
        Case Thickness < 47#:
            EXValue = "2"""
            
        Case Thickness < 63#:
            EXValue = "2-1/2"""
            
        Case Thickness < 75#:
            EXValue = "3"""
            
        Case Thickness < 95#:
            EXValue = "4"""
            
        Case Else:
            EXValue = "Special"
    
    End Select

    Dim swAddResult As swCustomInfoAddResult_e

    swAddResult = swPropMgr.Add3("EX", swCustomInfoType_e.swCustomInfoText, EXValue, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
    
    If swAddResult = swCustomInfoAddResult_AddedOrChanged Then
    
        PropAddedCount = PropAddedCount + 1
    
        Debug.Print EXValue & " added"
    
    End If
  
End Sub

Private Sub Updates(update As Boolean)

    Dim swView As SldWorks.ModelView
    Set swView = swModel.ActiveView
    swView.EnableGraphicsUpdate = update
    swModel.FeatureManager.EnableFeatureTree = update
    swModel.FeatureManager.EnableFeatureTreeWindow = update
    
End Sub
