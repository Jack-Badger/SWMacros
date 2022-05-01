Attribute VB_Name = "CloneCustomPropsFromActiv"
Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim swConfig As SldWorks.Configuration
    Set swConfig = swModel.GetActiveConfiguration
    
    Dim sInitialConfig As String
    sInitialConfig = swConfig.name
        
    Dim swCusPropMgr As SldWorks.CustomPropertyManager
    Set swCusPropMgr = swConfig.CustomPropertyManager


    
    Dim count As Long
    count = swCusPropMgr.count
    
    If count > 0 Then
    
        ReDim sNames(1 To count) As String
        ReDim svalues(1 To count) As String
        
        Dim vNames As Variant
        vNames = swCusPropMgr.GetNames
        
        Dim i As Long
        For i = 1 To count
            sNames(i) = vNames(i - 1)
            svalues(i) = GetValue(sNames(i), swCusPropMgr)
        Next i
    
    End If
    
    Dim vConfigNames As Variant
    vConfigNames = swModel.GetConfigurationNames
    

    For i = 0 To UBound(vConfigNames)
    
        swModel.ShowConfiguration2 vConfigNames(i)
        Set swConfig = swModel.GetActiveConfiguration
        Set swCusPropMgr = swConfig.CustomPropertyManager
        
        DeleteAll swCusPropMgr
        
        Dim j As Long
        
        For j = 1 To count
        
            swCusPropMgr.Add3 sNames(j), swCustomInfoText, Replace(svalues(j), "@" & sInitialConfig, "@" & vConfigNames(i)), swCustomPropertyOnlyIfNew

        Next j
        
    Next i
    
    swModel.ShowConfiguration2 sInitialConfig
    
    
End Sub

Private Function GetValue(name As String, swCusPropMgr As SldWorks.CustomPropertyManager) As String

    Dim result As swCustomInfoGetResult_e
    Dim sVal As String
    Dim sResVal As String
    Dim bResolved As Boolean
    Dim bLinked As Boolean
    
    result = swCusPropMgr.Get6(name, True, sVal, sResVal, bResolved, bLinked)
    
    GetValue = sVal
    
End Function

Private Sub DeleteAll(swCusPropMgr As SldWorks.CustomPropertyManager)

    If swCusPropMgr.count > 0 Then
    
        Dim vNames As Variant
        
        
        
        vNames = swCusPropMgr.GetNames
        
        Dim i As Long
        
        For i = 0 To UBound(vNames)
        
            swCusPropMgr.Delete2 vNames(i)
        
        Next i
    
    End If

End Sub


