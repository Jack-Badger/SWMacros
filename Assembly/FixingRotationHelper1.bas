Attribute VB_Name = "FixingRotationHelper1"
' Allows the first fixing in a folder to determine the angle of all fixings in same folder

Option Explicit

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swAssModel As SldWorks.ModelDoc2
    Set swAssModel = swApp.ActiveDoc
    
    If swAssModel Is Nothing Then
        Exit Sub
    End If
    
    If Not swAssModel.GetType = swDocumentTypes_e.swDocASSEMBLY Then
        Exit Sub
    End If
    
    Dim Folders As Collection
    Set Folders = GetSelectedFolders(swAssModel.SelectionManager)
    
    If Folders.count = 0 Then
    
        If swApp.SendMsgToUser2("No folders selected." & vbCr & "Would you like to use All folders?" _
                            , swMessageBoxIcon_e.swMbQuestion _
                            , swMessageBoxBtn_e.swMbYesNo) _
                            = swMessageBoxResult_e.swMbHitYes Then
        
            Set Folders = GetAllFolders(swAssModel)
        
        End If
        
    End If
    
    If Folders.count = 0 Then 'Last Chance
        
        Exit Sub
        
    End If
    
    Dim matePlane As Long
    
    matePlane = GetDefaultPlane
    
    If matePlane = 0 Then
    
        Exit Sub
    
    End If
    
    Dim randomAngle As Double
    randomAngle = GetRandomAngle
    
    Dim positionOnly As Boolean
    positionOnly = swApp.SendMsgToUser2("Position Only?", swMessageBoxIcon_e.swMbQuestion, swMessageBoxBtn_e.swMbYesNo) = swMessageBoxResult_e.swMbHitYes
    
    Updates swAssModel, False
    
    swApp.SendMsgToUser2 "Added " _
                        & ProcessFolders(swAssModel, Folders, matePlane, randomAngle, positionOnly) _
                        & " Mates." _
                        , swMessageBoxIcon_e.swMbInformation _
                        , swMessageBoxBtn_e.swMbOk
    
    Updates swAssModel, True
    
End Sub

Private Function ProcessFolders(swModel As SldWorks.ModelDoc2, Folders As VBA.Collection, matePlane As Long, angleDeviation As Double, positionOnly As Boolean) As Long

    Dim totalCount As Long
    
    Dim i As Long
    For i = 1 To Folders.count
    
        totalCount = totalCount + ProcessFolder(swModel, Folders(i), matePlane, angleDeviation, positionOnly)
    
    Next i
    
    ProcessFolders = totalCount
    
End Function

Private Function ProcessFolder(swAssModel As SldWorks.ModelDoc2, swFolderFeature As SldWorks.Feature, matePlane As Long, angleDeviation As Double, positionOnly As Boolean) As Long

    Dim count As Long
    
    Dim swAssem As SldWorks.AssemblyDoc
    Set swAssem = swAssModel
    
    Dim swFeatFolder As SldWorks.FeatureFolder
    Set swFeatFolder = swFolderFeature.GetSpecificFeature2
    
    If swFeatFolder.GetFeatureCount < 2 Then
    
        ProcessFolder = 0
    
        Exit Function
    
    End If
    
    Dim vFeatures As Variant
    
    vFeatures = swFeatFolder.GetFeatures
    
    Dim swFirstFeature As SldWorks.Feature
    Set swFirstFeature = vFeatures(0)
    
    Dim swEntity As SldWorks.Entity
    Set swEntity = swFirstFeature
    
    If swEntity.GetType <> swSelectType_e.swSelCOMPONENTS Then
    
        ProcessFolder = 0
    
        Exit Function
    
    End If
    
    Debug.Print swEntity.GetType
    Debug.Print swFirstFeature.Name
    
    Dim swFirstComponent As SldWorks.Component2
    Set swFirstComponent = swAssem.GetComponentByName(swFirstFeature.Name)
    swFirstComponent.Select4 False, Nothing, False
    swAssem.FixComponent
    
    Dim swFirstModel As SldWorks.ModelDoc2
    Set swFirstModel = swFirstComponent.GetModelDoc2
    
    Dim FirstTitle As String
    FirstTitle = swFirstModel.GetTitle
    
    Dim swFirstModelMatePlane As SldWorks.Feature
    Set swFirstModelMatePlane = GetMatePlane(swFirstComponent.GetModelDoc2, matePlane)
    
    Dim matePlaneName As String
    matePlaneName = swFirstModelMatePlane.Name
    
    Set swFirstModelMatePlane = swFirstComponent.GetCorresponding(swFirstModelMatePlane)
            
    Randomize
    
    Dim i As Long
    
    For i = 1 To UBound(vFeatures)
    
        Dim swSlaveFeature As SldWorks.Feature
        Dim swSlaveEntity As SldWorks.Entity
        Dim swSlaveComponent As SldWorks.Component2
        Dim swSlaveModel As SldWorks.ModelDoc2
        Dim swSlaveMatePlane As SldWorks.Feature
        
        Set swSlaveFeature = vFeatures(i)
        Set swSlaveEntity = swSlaveFeature
        
        If swSlaveEntity.GetType <> swSelectType_e.swSelCOMPONENTS Then
    
            GoTo continueNext
    
        End If
        
        Set swSlaveComponent = swAssem.GetComponentByName(swSlaveFeature.Name)
        

        
        Set swSlaveModel = swSlaveComponent.GetModelDoc2
        
        If StrComp(swSlaveModel.GetTitle, FirstTitle) <> 0 Then
        
            GoTo continueNext
        
        End If
        
        Select Case swSlaveModel.GetType
        
            Case swDocumentTypes_e.swDocASSEMBLY:
            
                Dim swSlaveAssem As SldWorks.AssemblyDoc
                Set swSlaveAssem = swSlaveModel
                Set swSlaveMatePlane = swSlaveAssem.FeatureByName(matePlaneName)
            
            Case swDocumentTypes_e.swDocPART:
            
                Dim swSlavePart As SldWorks.PartDoc
                Set swSlavePart = swSlaveModel
                Set swSlaveMatePlane = swSlavePart.FeatureByName(matePlaneName)
        
            Case Else:
            
                GoTo continueNext
        
        End Select
        
            Set swSlaveMatePlane = swSlaveComponent.GetCorresponding(swSlaveMatePlane)
        
            swFirstModelMatePlane.Select2 False, 1
            swSlaveMatePlane.Select2 True, 1

            
            Dim angle As Double
            angle = Rnd * angleDeviation
            
            
            If Rnd >= 0.5 Then
            
                angle = 360 - angle
                
            End If
                 
            Debug.Print angle
                 
            If Not AddMate(swAssem, swMateANGLE, swMateAlignALIGNED, False, ToRadians(angle), positionOnly) Is Nothing Then
            
                count = count + 1
            
            End If
            

            
continueNext:
    Next i
    
        swFirstComponent.Select4 False, Nothing, False
        swAssem.UnfixComponent
        Dim vMates As Variant
        vMates = swFirstComponent.GetMates
        
        Dim swMateFeat As SldWorks.Feature
        
        For i = 0 To UBound(vMates)
        
            Set swMateFeat = vMates(i)
            
            If swMateFeat.IsSuppressed Then
            
                swMateFeat.SetSuppression2 swFeatureSuppressionAction_e.swUnSuppressFeature, swInConfigurationOpts_e.swAllConfiguration, Empty
            
            End If
            
        Next i
        
    
    ProcessFolder = count

End Function

Private Function AddMate(swAssem As SldWorks.AssemblyDoc, Optional MateType As swMateType_e = swMateCOINCIDENT, Optional Align As swMateAlign_e = swMateAlignCLOSEST, Optional Flip As Boolean = False, Optional angle As Double = 0#, Optional positionOnly As Boolean = True) As SldWorks.Mate2
    
    Dim swMate As SldWorks.Mate2
    
    Dim error As swAddMateError_e
        
    Set swMate = swAssem.AddMate5(MateType, Align, Flip, 0#, 0#, 0#, 0#, 0#, angle, angle, angle, positionOnly, False, 0, error)
    
    If swMate Is Nothing Then
        Debug.Print "Mate Error:" & error
    End If
    
    Set AddMate = swMate
    
End Function


Private Function GetMatePlane(swModel As SldWorks.ModelDoc2, num As Long) As SldWorks.Feature

    Dim swFeature As SldWorks.Feature
    Set swFeature = swModel.FirstFeature
    
    Dim swEntity As SldWorks.Entity
    Set swEntity = swFeature
    
    Do While Not swEntity.GetType = swSelectType_e.swSelDATUMPLANES
    
        Set swFeature = swFeature.GetNextFeature
        Set swEntity = swFeature
    Loop
    
    Debug.Print swFeature.Name
    
    Dim i As Long
    For i = 1 To num - 1
    
        Set swFeature = swFeature.GetNextFeature
    
    Next i
    
    Set GetMatePlane = swFeature
    
End Function

Private Function GetDefaultPlane() As Long

    Dim userResponse As Long
    
    Do While Not (userResponse > 0 And userResponse < 4)
    
    On Error GoTo CancelHandler
        userResponse = VBA.InputBox("Please Enter Default Plane number (1-3)", "Default Plane Selection")
    On Error GoTo 0
    Loop
    
    GetDefaultPlane = userResponse

    Exit Function

CancelHandler:
    GetDefaultPlane = 0
    On Error GoTo 0

End Function

Private Function GetRandomAngle() As Double

    Dim userResponse As Double
    userResponse = 999.9
    
    Do While Not (userResponse >= 0# And userResponse <= 360#)
    
    On Error GoTo CancelHandler
        userResponse = VBA.InputBox("Please Enter Angle deviation in degrees", "Random Angle Selection")
    On Error GoTo 0
    Loop
    
    GetRandomAngle = userResponse

    Exit Function

CancelHandler:
    GetRandomAngle = 0#
    On Error GoTo 0

End Function


Private Function GetAllFolders(swModel As SldWorks.ModelDoc2) As VBA.Collection

    Set GetAllFolders = New Collection
    
    Dim swFeature As SldWorks.Feature
    
    Set swFeature = swModel.FirstFeature
    
    Do While Not swFeature Is Nothing
    
        Dim typeName As String
        typeName = swFeature.GetTypeName2
    
        If StrComp(typeName, "FtrFolder") = 0 Then
        
            If InStrRev(swFeature.Name, "___EndTag___") = 0 Then
            
                GetAllFolders.Add swFeature
                'Debug.Print swFeature.Name
            
            End If
                    
        End If
        
        Set swFeature = swFeature.GetNextFeature
        
    Loop

End Function

Private Function GetSelectedFolders(swSelMgr As SldWorks.SelectionMgr) As VBA.Collection

    Set GetSelectedFolders = New Collection
    
    If swSelMgr.GetSelectedObjectCount > 0 Then
    
        Dim i As Long
        
        For i = 1 To swSelMgr.GetSelectedObjectCount
        
            If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelFTRFOLDER Then
            
                GetSelectedFolders.Add swSelMgr.GetSelectedObject6(i, -1)
            
            End If
        
        Next i
    
    End If
    
End Function

Private Function ToRadians(degrees As Double) As Double

    ToRadians = degrees / 57.2957795130823

End Function

Private Sub Updates(swModel As SldWorks.ModelDoc2, update As Boolean)

    'Dim swView As SldWorks.ModelView
    'Set swView = swModel.ActiveView
    'swView.EnableGraphicsUpdate = update
    swModel.FeatureManager.EnableFeatureTree = update
    swModel.FeatureManager.EnableFeatureTreeWindow = update
    
End Sub
