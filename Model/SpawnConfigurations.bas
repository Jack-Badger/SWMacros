' Under Development!

Option Explicit

Dim swApp As SldWorks.SldWorks

Enum NoSelectionsOption_e

    Default = 0
    ReturnEmpty = 1
    ReturnAll = 2
    AskUser = 3

End Enum

Enum SpawnFormat_e
    
    Default = 0
    Prefix = 1
    Suffix = 2
    ReplaceText = 3

End Enum

Const PartsAndAssembliesFilter = "SOLIDWORKS Parts/Assems (*.sldprt; *.sldasm)|*.sldprt;*.sldasm|All Files (*.*)|*.*|"

Sub main()

    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim NoSelectionsOption As NoSelectionsOption_e
    NoSelectionsOption = AskUser
    
    If swModel Is Nothing Then
    
        Dim path As String
        path = GetOpenFileName(title:="Select File to Spawn Configurations", fileFilter:=PartsAndAssembliesFilter)
        
        If path = "" Then Exit Sub
        
        Set swModel = OpenDocumentSilentReadonly(path, silent:=True, readOnly:=False, setWorkDir:=True, resetWorkDir:=False, activate:=True)
            
        If swModel Is Nothing Then
        
            SendMessage "Could not open " & path
            
            Exit Sub
        
        End If
        
        If SendMessage _
            ("SPAWN CONFIGURATIONS" & vbNewLine & vbNewLine _
            & "To apply to all configurations" & vbNewLine & vbNewLine _
            & "  * Hit Yes" & vbNewLine & vbNewLine _
            & "To select configurations" & vbNewLine & vbNewLine _
            & "  * Hit No" & vbNewLine _
            & "  * Make your selections" & vbNewLine _
            & "  * Run this macro again" _
            , swMbQuestion _
            , swMbYesNo) = swMbHitNo Then
            
                Exit Sub
            
        End If
            
        NoSelectionsOption = ReturnAll
            
    End If
    
    Select Case swModel.GetType
    
        Case swDocASSEMBLY
        
            Dim SelectedComponents As VBA.Collection
            Set SelectedComponents = SiftSelectionsIntoCollections(swModel.SelectionManager, swSelCONFIGURATIONS, swSelCOMPONENTS)
        
            If SelectedComponents.count > 0 Then

            End If
        
        Case swDocPART
            
                Dim configNames As VBA.Collection
                Set configNames = CollectConfigurationNames(swModel, NoSelectionsOption)
                
                If configNames.count = 0 Then Exit Sub
                
                Dim deleteBase As Boolean
                
                Select Case SendMessage("Would you like to DELETE the BASE config?", swMbQuestion, swMbYesNoCancel)
                
                    Case swMbHitYes: deleteBase = True
                        
                    Case swMbHitNo: deleteBase = False
                    
                    Case swMbHitCancel: Exit Sub
                    
                End Select
                
                Dim spawnNames As VBA.Collection
                Set spawnNames = AskUserForSpawnNames("", "Enter Many Spawn Text")
                
                If spawnNames.count = 0 Then Exit Sub
                
                Dim format As SpawnFormat_e
                format = SpawnFormat_e.Default
                
                Dim placeHolder As String
                Dim separator As String
                
                If format = SpawnFormat_e.Default Then format = AskForFormat()
                
                Select Case format
                
                    Case Prefix, Suffix
                    
                            separator = InputBox("Please Enter Separator", "Spawn Configurations Separator", Default:="-")
                    
                    Case ReplaceText
                    
                    placeHolder = InputBox("Please Enter Text To Find", "Spawn Configurations Find Replace")
                    
                End Select


                On Error GoTo EnableUpdates
                
                    Updates swModel, False
                    
                        SpawnConfigurations swModel, configNames, spawnNames, format, deleteBase, separator, placeHolder
                    
                    Updates swModel, True
                    
                On Error GoTo 0

        Case Else
        
            Exit Sub
            
    End Select

    Exit Sub
    
EnableUpdates:

    Updates swModel, True
    
    SendMessage "summit went wrong" & Err.Description

End Sub

Private Function SiftSelectionsIntoCollections(swSelectionManager As SldWorks.SelectionMgr, ParamArray keys() As Variant) As VBA.Collection

    Dim topCollection As VBA.Collection
    Set topCollection = New VBA.Collection

    Set SiftSelectionsIntoCollections = topCollection

    If swSelectionManager Is Nothing Then Exit Function

    If UBound(keys) = 0 Then Exit Function

    Dim i As Long
    For i = 0 To UBound(keys)

        Dim subCollection As VBA.Collection
        Set subCollection = New VBA.Collection

        topCollection.Add subCollection, CStr(keys(i))

    Next i

    With swSelectionManager

        Dim selectionCount As Long
        selectionCount = .GetSelectedObjectCount2(-1)
        
        If selectionCount = 0 Then Exit Function
        
        
        For i = 1 To selectionCount
        
            Dim key As swSelectType_e
            key = .GetSelectedObjectType3(i, -1)
            
            'Dim subCollection As VBA.Collection
            Set subCollection = topCollection(CStr(key))

            subCollection.Add .GetSelectedObject6(i, -1)

            Set subCollection = Nothing

        Next i

    End With
    
End Function

Private Function AskForFormat() As SpawnFormat_e

    Select Case SendMessage("Would you like to REPLACE text?", swMbQuestion, swMbYesNo)
    
        Case swMbHitYes: AskForFormat = ReplaceText
        
        Case swMbHitNo
        
            Select Case SendMessage("YES for PREFIX" & vbNewLine & "NO for SUFFIX", swMbQuestion, swMbYesNo)
            
                Case swMbHitYes: AskForFormat = Prefix
            
                Case swMbHitNo: AskForFormat = Suffix
            
            End Select

    End Select

End Function


Private Function SpawnConfigurations(swModel As SldWorks.ModelDoc2, baseNames As VBA.Collection, spawnNames As VBA.Collection, format As SpawnFormat_e, Optional deleteBase As Boolean = True, Optional separator As String = "-", Optional placeHolder As String = "")

    Dim total As Long
    total = baseNames.count * spawnNames.count
    
    Dim swProgressBar As SldWorks.UserProgressBar

    Dim title As String
    title = "Spawning * Configs"
    
    title = Replace(title, "*", CStr(total))
    
    If total > 10 Then
    
            If swApp.GetUserProgressBar(swProgressBar) Then
            
                swProgressBar.Start 0, total, title
                
            End If
    End If

    Dim i As Long
    For i = 1 To baseNames.count

        Dim j As Long
        For j = 1 To spawnNames.count

            swModel.ShowConfiguration2 baseNames(i)

            Dim newName As String
            newName = CreateNewConfigurationName(baseNames(i), spawnNames(j), format, separator, placeHolder)

            swModel.AddConfiguration3 newName, "", "", 0

        Next j
        
        If deleteBase Then

            Dim swConfiguration As SldWorks.Configuration
            Set swConfiguration = swModel.GetConfigurationByName(baseNames(i))

            swConfiguration.Select2 False, Nothing

            swModel.EditDelete

        End If
        
        If Not swProgressBar Is Nothing Then
        
            Dim processedCount As Long
            
            processedCount = i * spawnNames.count
            
            swProgressBar.UpdateProgress processedCount
            
            'title = Replace(title, "*", CStr(processedCount))
        
            'swProgressBar.UpdateTitle title
            
        End If

    Next i
    
    If Not swProgressBar Is Nothing Then
    
        swProgressBar.End
    
    End If

End Function

Private Function CreateNewConfigurationName(baseName As String, spawnName As String, format As SpawnFormat_e, Optional separator As String = "-", Optional placeHolder As String = "")

    Dim name As String

    Select Case format

        Case Prefix
            
            name = spawnName & separator & baseName
        
        Case Suffix
        
            name = baseName & separator & spawnName
        
        Case ReplaceText
        
            name = Replace(baseName, find:=placeHolder, Replace:=spawnName)
        
        Case Else
        
            Err.Raise vbObjectError, "", "Unimplemented Exception: Default format"

    End Select

    CreateNewConfigurationName = name

End Function

Private Function AskUserForSpawnNames(prompt As String, Optional title As String = "") As VBA.Collection

    Dim spawnNames As New VBA.Collection
    
    Dim name As String
    
    Dim count As Long
    
    Do
        
        count = count + 1
        
        name = VBA.InputBox(CStr(count) & " " & prompt, title)
        
        If name = "" Then Exit Do
        
        'ToDO maybe - error checking of special characters.
        '
        'if not IsValid(name) then
        '
        '    Err.Raise vbObjectError, blah, blah
        '
        'End If

        spawnNames.Add name

    Loop

    Set AskUserForSpawnNames = spawnNames

End Function

Private Function CollectConfigurationNames _
    (swModel As SldWorks.ModelDoc2 _
    , Optional noSelectionOption As NoSelectionsOption_e = AskUser _
    ) As VBA.Collection

    Dim configurationsToSpawn As New VBA.Collection
    
    AddNamesOfSelectionsToCollection _
        swSelectionManager:=swModel.SelectionManager _
        , swSelectType:=swSelCONFIGURATIONS _
        , vbCollection:=configurationsToSpawn
    
    If configurationsToSpawn.count = 0 Then

        If noSelectionOption = AskUser Then
        
            Dim configNames As String
            configNames = GetConfigNamesList(swModel, 3, vbNewLine)

            Dim question As String
            question = "There are no selections.  Would you like to apply to all configurations?"
            
            question = question & vbNewLine & vbNewLine & configNames
        
            Dim swMessageBoxResult As swMessageBoxResult_e
            swMessageBoxResult = SendMessage(question, swMbQuestion, swMbYesNo)
            
            noSelectionOption = IIf(swMessageBoxResult = swMbHitYes, ReturnAll, ReturnEmpty)
        
        End If
        
        If noSelectionOption = ReturnAll Then
        
            AddItemsToCollection vItems:=swModel.GetConfigurationNames _
                               , vbCollection:=configurationsToSpawn
            
        End If
    
    End If

    Set CollectConfigurationNames = configurationsToSpawn

End Function

Private Function GetConfigNamesList(swModel As SldWorks.ModelDoc2, Optional max As Long = 10, Optional separator As String = vbNewLine) As String

    Dim vNames As Variant
    vNames = swModel.GetConfigurationNames

    Dim configCount As Long
    configCount = swModel.GetConfigurationCount

    Dim nameList As String

    Dim i As Long
    For i = 0 To UBound(vNames)

        If i >= max Then
        
            Dim overflowMessage As String
            overflowMessage = "[" & CStr(configCount - max) & " more]"
            
            nameList = nameList & separator & overflowMessage

            Exit For

        End If
        
        If i > 0 Then nameList = nameList & separator

        nameList = nameList & vNames(i)

    Next i

    GetConfigNamesList = nameList

End Function



Private Sub AddNamesOfSelectionsToCollection(swSelectionManager As SldWorks.SelectionMgr, swSelectType As swSelectType_e, vbCollection As VBA.Collection)

    With swSelectionManager

        Dim count As Long
        count = .GetSelectedObjectCount2(-1)
    
        If count = 0 Then Exit Sub
    
        Dim i As Long
        For i = 1 To count
    
            Dim name As String
    
            If .GetSelectedObjectType3(i, -1) = swSelectType Then
    
                Select Case swSelectType

                    Case swSelCONFIGURATIONS
                    
                        Dim swConfiguration As SldWorks.Configuration
                        Set swConfiguration = .GetSelectedObject6(i, -1)
                        
                        name = swConfiguration.name
                    
                    Case Else
                    
                        Err.Raise vbObjectError + 1, "SpawnConfigurations1.AddNamesToCollection", "Unimplemented Exception! swSelectType_e: " & CStr(swSelectType)
                    
                End Select
    
                vbCollection.Add name
    
            End If
    
        Next i
    
    End With

End Sub

Private Sub AddItemsToCollection(vItems As Variant, vbCollection As VBA.Collection)

    Dim i As Long
    For i = 0 To UBound(vItems)

'        Dim item As String
'        item = vItems(i)

        vbCollection.Add vItems(i)

    Next i

End Sub

Private Sub Updates(swModel As SldWorks.ModelDoc, update As Boolean)

    Dim swView As SldWorks.ModelView
    Set swView = swModel.ActiveView
    swView.EnableGraphicsUpdate = update
    swModel.FeatureManager.EnableFeatureTree = update
    swModel.FeatureManager.EnableFeatureTreeWindow = update
    
End Sub

Private Function SendMessage _
    (message As String _
    , Optional icon As swMessageBoxIcon_e = swMessageBoxIcon_e.swMbInformation _
    , Optional buttons As swMessageBoxBtn_e = swMessageBoxBtn_e.swMbOk _
    ) As swMessageBoxResult_e

    SendMessage = swApp.SendMsgToUser2(message, icon, buttons)

End Function

Private Function GetOpenFileName _
    (Optional title As String = "Open File" _
    , Optional initialFileName As String = "" _
    , Optional fileFilter As String = "SOLIDWORKS Files (*.sldprt; *.sldasm; *.slddrw)|*.sldprt;*.sldasm;*.slddrw|Filter name (*.fil)|*.fil|All Files (*.*)|*.*|" _
    , Optional configname As String = "" _
    , Optional displayName As String = "" _
    ) As String

    Dim notUsed As Long

    GetOpenFileName = swApp.GetOpenFileName(title, initialFileName, fileFilter, notUsed, configname, displayName)

End Function

Public Function OpenDocumentSilentReadonly _
    (path As String _
    , Optional silent As Boolean = True _
    , Optional readOnly As Boolean = True _
    , Optional setWorkDir As Boolean = True _
    , Optional resetWorkDir As Boolean = True _
    , Optional activate As Boolean = False _
    ) As SldWorks.ModelDoc2

    If resetWorkDir Then
        Dim workingDirectory As String
        workingDirectory = swApp.GetCurrentWorkingDirectory
    End If

    Dim RightSlash As Long
    RightSlash = InStrRev(path, "\")
    
    Dim Folder As String
    Folder = Left(path, RightSlash - 1)
    
    Debug.Print Folder
    
    swApp.SetCurrentWorkingDirectory Folder

    Dim swDocSpec As SldWorks.DocumentSpecification
    Set swDocSpec = swApp.GetOpenDocSpec(path)
    
    swDocSpec.silent = silent
    swDocSpec.readOnly = readOnly
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.OpenDoc7(swDocSpec)
    
    If Not swModel Is Nothing Then
    
        If activate Then
        
            Dim swActivateDocError As swActivateDocError_e
        
            swApp.ActivateDoc3 swModel.GetTitle, True, swRebuildOnActivation_e.swDontRebuildActiveDoc, swActivateDocError
        
            If swActivateDocError <> swGenericActivateError Then
            
                Debug.Print path, "swGenericActivateError"
                
            End If
        
        End If
    
    End If

    If resetWorkDir Then
    
        If Not swApp.SetCurrentWorkingDirectory(workingDirectory) Then
            Debug.Print "Did not reset working directory", workingDirectory
        End If

    End If
    
    Set OpenDocumentSilentReadonly = swModel

End Function


