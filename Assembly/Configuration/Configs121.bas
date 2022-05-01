Attribute VB_Name = "Configs121"
' SOLIDWORKS 2016
' Configs121
' by 369
'
' Inspired by Best Practice for Configurations with External References
' http://help.solidworks.com/2016/english/SolidWorks/sldworks/hid_extern_refs_msg_warning_help.htm
'
' PreConditions
'                   An Assembly Document is open
'                   Some or No Components selected
'                   Some or No Configurations selected
'
' Post Conditions
'                   Every Highlighted Component
'                               or
'                   All showing Components (No pre-Selection)
'
'                   Every* Highlighted Configuration
'                               or
'                   All Configurations* (No pre-Selection)
'
'                   The component configurations will be created / switched to
'                   assembly configuration on all highlighted components
'
' Limitations*      many, notably derived configs traversal and diving into subs
'                   setting options etc...  just a brain fart really.
'
'                   Enjoyed writing a gosub routine first time since 1984 / C64

Option Explicit
    Dim swApp       As SldWorks.SldWorks
    Dim box         As BufferBox
    Dim components  As VBA.Collection
    Dim models      As VBA.Collection
    Dim configs     As VBA.Collection

Sub main()
    Set swApp = Application.SldWorks
    Set box = New BufferBox
    With swApp: Select Case True
       Case .ActiveDoc Is Nothing _
          , .ActiveDoc.GetType <> swDocASSEMBLY
             box.SendMsg "Please Open an Assembly Document"
       Case Else
           box.Silent = True
           Call ProcessAssembly(.ActiveDoc)
    End Select: End With
End Sub

Private Function ProcessAssembly(swAssem As SldWorks.ModelDoc2)
    Call ProcessSelections(swAssem)
    Call CreateConfigurations
    Call TelegraphConfigurations(swAssem)
End Function

Private Function TelegraphConfigurations(swAssem As SldWorks.AssemblyDoc)
    Dim i As Integer
    Dim j As Integer
    With swAssem
    For i = 1 To configs.Count
       .ShowConfiguration2 configs.Item(i).Name
        For j = 1 To components.Count
                 .GetComponentByID(components.Item(j)) _
                 .ReferencedConfiguration _
                = configs.Item(i).Name
                  Call .EditRebuild3
    Next: Next
    End With
End Function
               
Private Function ProcessSelections(swAssem As SldWorks.AssemblyDoc)
    
    With swAssem: Dim i As Integer
    
        GoSub GetComponentIDs
        GoSub ListComponents
        
        GoSub GetModels
        GoSub ListModels
           
        GoSub GetConfigurations
        GoSub ListConfigurations
            
    Exit Function
        
GetComponentIDs:
        GoSub AddPreselectedComponents
        If components.Count = 0 Then GoSub AddAllComponents
        Return
    
AddPreselectedComponents:
        Set components = CollectTheseTypes(.SelectionManager, swSelCOMPONENTS)
        Return
    
AddAllComponents:
        Set components = CollectAllComponentIDs(swAssem)
        If components.Count = 0 Then
            box.SendMsg "No Components!"
            Exit Function
        End If
        Return
        
GetModels:
        Set models = CollectUniqueModels(components, swAssem)
        Return

GetConfigurations:
        Set configs = CollectTheseTypes(.SelectionManager, swSelCONFIGURATIONS)
        If configs.Count = 0 Then Set configs = CollectTopLevelConfigs(swAssem)
        Return
    
ListComponents:
        box.SendMsg "Selected Components: " & components.Count
        box.BufferSize = 10
        For i = 1 To components.Count
            box.Add .GetComponentByID(components.Item(i)).Name2
        Next
        box.Flush
        Return
        
ListModels:
        box.SendMsg "unique models: " & models.Count
        box.BufferSize = 10
        For i = 1 To models.Count
            box.Add models.Item(i).GetTitle
        Next
        box.Flush
        Return
        Return
    
ListConfigurations:
        box.SendMsg "Selected Configurations: " & configs.Count
        box.BufferSize = 10
        For i = 1 To configs.Count
            box.Add configs.Item(i).Name
        Next
        box.Flush
        Return
        
    End With
End Function
    
Private Function _
        CollectTheseTypes(swSelMgr As SldWorks.SelectionMgr, selType As Long) _
     As VBA.Collection
    Set CollectTheseTypes = New VBA.Collection
   With swSelMgr
        Dim i As Integer
        For i = 1 To .GetSelectedObjectCount2(Mark:=-1)
            If .GetSelectedObjectType3(Index:=i, Mark:=-1) = selType Then
                Select Case selType
                    Case swSelCONFIGURATIONS
                         CollectTheseTypes.Add .GetSelectedObject6(Index:=i, Mark:=-1)
                    Case swSelCOMPONENTS
                         CollectTheseTypes.Add .GetSelectedObject6(Index:=i, Mark:=-1).GetID
            End Select
            End If
        Next
    End With
End Function
    
Private Function _
        CollectAllComponentIDs(swAssem As SldWorks.AssemblyDoc) _
     As VBA.Collection
    Set CollectAllComponentIDs = New VBA.Collection
   With swAssem
        Dim vComponents As Variant
        vComponents = .GetComponents(ToplevelOnly:=True)
        Dim i As Integer
        For i = 0 To .GetComponentCount(ToplevelOnly:=True) - 1
            Dim swComponent As SldWorks.Component2
            Set swComponent = vComponents(i)
            If swComponent.Visible Then CollectAllComponentIDs.Add swComponent.GetID
        Next
    End With
End Function

Private Function _
        CollectUniqueModels(components As VBA.Collection, swAssem As SldWorks.AssemblyDoc) _
     As VBA.Collection
    Set CollectUniqueModels = New VBA.Collection
    Dim i As Integer: For i = 1 To components.Count
        Dim swModel As SldWorks.ModelDoc2
        Set swModel = swAssem.GetComponentByID(components.Item(i)).GetModelDoc2
        On Error Resume Next
            CollectUniqueModels.Add Item:=swModel, Key:=swModel.GetTitle
        On Error GoTo 0
    Next
End Function

Private Function _
        CollectTopLevelConfigs(swModel As SldWorks.ModelDoc2) _
     As VBA.Collection
    Set CollectTopLevelConfigs = New VBA.Collection
    With swModel
        Dim vConfigNames As Variant
        vConfigNames = .GetConfigurationNames
        Dim i As Integer
        For i = 0 To .GetConfigurationCount - 1
            Dim swConfiguration As SldWorks.configuration
            Set swConfiguration = .GetConfigurationByName(vConfigNames(i))
             If swConfiguration.IsDerived = False Then CollectTopLevelConfigs.Add swConfiguration
        Next
    End With
End Function

Private Function CreateConfigurations()
    Dim m As Integer: For m = 1 To models.Count
    Dim c As Integer: For c = 1 To configs.Count
        Call models.Item(m) _
            .AddConfiguration3 _
            (configs.Item(c).Name, "", "", 0)
    Next c: Next m
End Function


