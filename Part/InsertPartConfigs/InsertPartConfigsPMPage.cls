VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InsertPartConfigsPMPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const ccDebug = True

Implements PropertyManagerPage2Handler9

Enum Control_ID

    messageGroup = 1&
        messageLabel = 2&

    transferGroup = 3&
        solidBodiesCheckBox = 4&
        surfaceBodiesCheckBox = 5&
        axesCheckBox = 6&
        planesCheckBox = 7&
        cosmeticThreadsCheckBox = 8&
        absorbedSketchesCheckBox = 9&
        unabsorbedSketchesCheckBox = 10&
        customPropertiesCheckBox = 11&
            customToFilePropertiesOption = 12&
            customToToCutListPropertiesOption = 13&
        cutlistPropertiesCheckBox = 14&
        coordinateSystemsCheckBox = 15&
        modelDimensionsCheckBox = 16&
        holeWizardDataCheckBox = 17&
        sheetMetalInformationCheckBox = 18&
            unlockedPropertiesCheckBox = 19&
    
    visualPropertiesGroup = 20&
        propagateVisualPropertiesCheckBox = 21&
        
    linkGroup = 22&
        breakLinkToOriginalPartCheckBox = 22&
         
    configurationsGroup = 23&
        configurationSelectionsALLCheckBox = 24&
        configurationSelectionsINCLUDEOption = 25&
        configurationSelectionsEXCLUDEOption = 26&
        configurationSelectionsLabel = 27&
        configurationsListBox = 28&
        
End Enum

Private longErrors As Long
Private boolStatus As Boolean

Private pm_Page As PropertyManagerPage2

    Private pm_messageGroup As PropertyManagerPageGroup
        Private pm_messageLabel As PropertyManagerPageLabel
        
    Private pm_transferGroup As PropertyManagerPageGroup
        Private pm_solidBodiesCheckBox As PropertyManagerPageCheckbox
        Private pm_surfaceBodiesCheckBox As PropertyManagerPageCheckbox
        Private pm_axesCheckBox As PropertyManagerPageCheckbox
        Private pm_planesCheckBox As PropertyManagerPageCheckbox
        Private pm_cosmeticThreadsCheckBox As PropertyManagerPageCheckbox
        Private pm_absorbedSketchesCheckBox As PropertyManagerPageCheckbox
        Private pm_unabsorbedSketchesCheckBox As PropertyManagerPageCheckbox
        Private pm_customPropertiesCheckBox As PropertyManagerPageCheckbox
            Private pm_customToFilePropertiesOption As PropertyManagerPageOption
            Private pm_customToToCutListPropertiesOption As PropertyManagerPageOption
        Private pm_cutlistPropertiesCheckBox As PropertyManagerPageCheckbox
        Private pm_coordinateSystemsCheckBox As PropertyManagerPageCheckbox
        Private pm_modelDimensionsCheckBox As PropertyManagerPageCheckbox
        Private pm_holeWizardDataCheckBox As PropertyManagerPageCheckbox
        Private pm_sheetMetalInformationCheckBox As PropertyManagerPageCheckbox
            Private pm_unlockedPropertiesCheckBox As PropertyManagerPageCheckbox
    
    Private pm_visualPropertiesGroup As PropertyManagerPageGroup
        Private pm_propagateVisualPropertiesCheckBox As PropertyManagerPageCheckbox
    
    Private pm_linkGroup As PropertyManagerPageGroup
        Private pm_breakLinkToOriginalPartCheckBox As PropertyManagerPageCheckbox
    
    Private pm_configurationsGroup As PropertyManagerPageGroup
        Private pm_configurationSelectionsALLCheckBox As PropertyManagerPageCheckbox
        Private pm_configurationSelectionsINCLUDEOption As PropertyManagerPageOption
        Private pm_configurationSelectionsEXCLUDEOption As PropertyManagerPageOption
        Private pm_configurationSelectionsLabel As PropertyManagerPageLabel
        Private pm_configurationsListBox As PropertyManagerPageListbox


Private Sub UpdateConfigurationsLabel()
    pm_configurationSelectionsLabel.Caption = GetSelectedConfigsCount & " of " & pm_configurationsListBox.ItemCount & " configs selected"
End Sub

Private Function GetSelectedConfigsCount() As Long
   GetSelectedConfigsCount = IIf(pm_configurationSelectionsINCLUDEOption.Checked _
                                , pm_configurationsListBox.GetSelectedItemsCount _
                                , pm_configurationsListBox.ItemCount - pm_configurationsListBox.GetSelectedItemsCount)
End Function

Private Function GetSelectedConfigs() As Variant
    
    If pm_configurationsListBox.GetSelectedItemsCount = 0 Then ' No selections
        GetSelectedConfigs = IIf(pm_configurationSelectionsEXCLUDEOption.Checked _
                                 , configNames _
                                 , vbNull) ' All or nothing!
    Else
        ReDim configurationsToInsert(GetSelectedConfigsCount - 1) As String
        Dim index_configToInsert As Integer
        index_configToInsert = 0
        
        Dim selections As Variant
        selections = pm_configurationsListBox.GetSelectedItems
        Dim index_selections As Integer
        Rem Whilst not specified in the API, it is observed selections are returned in reverse order
        index_selections = UBound(selections) ' index backwards
            
        Dim index_configNames As Integer
        For index_configNames = 0 To UBound(configNames)
        
            Debug.Print "index_configNames " & index_configNames _
                      , "selections(index_selections) " & selections(index_selections)
                      
            If selections(index_selections) = index_configNames Then
                
                If pm_configurationSelectionsINCLUDEOption.Checked Then
                    configurationsToInsert(index_configToInsert) = configNames(index_configNames)
                    index_configToInsert = index_configToInsert + 1
                    Debug.Print "config+"
                End If
                
                If index_selections > 0 Then ' non-optimal but simple
                    index_selections = index_selections - 1
                    Debug.Print "selection-"
                End If
          
            Else
                
                If pm_configurationSelectionsEXCLUDEOption.Checked Then
                    configurationsToInsert(index_configToInsert) = configNames(index_configNames)
                    index_configToInsert = index_configToInsert + 1
                    Debug.Print "config+"
                End If
    
            End If
    
        Next index_configNames
        
        GetSelectedConfigs = configurationsToInsert
    End If
End Function

Private Function GetInsertOptions() As Long
    
    Dim options As Long
    options = 0
    
    If pm_solidBodiesCheckBox.Checked Then options = options + swInsertPartImportSolids
    If pm_surfaceBodiesCheckBox.Checked Then options = options + swInsertPartImportSurfaces
    If pm_axesCheckBox.Checked Then options = options + swInsertPartImportAxes
    If pm_planesCheckBox.Checked Then options = options + swInsertPartImportPlanes
    If pm_cosmeticThreadsCheckBox.Checked Then options = options + swInsertPartImportCosmeticThreads
    If pm_absorbedSketchesCheckBox.Checked Then options = options + swInsertPartImportAbsorbedSketchs
    If pm_unabsorbedSketchesCheckBox.Checked Then options = options + swInsertPartImportUnabsorbedSketchs
    If pm_customPropertiesCheckBox.Checked Then
        options = options + swInsertPartImportCustomProperties
        If pm_customToFilePropertiesOption.Checked Then options = options + swInsertPartImportCustomToFileProperties
        If pm_customToToCutListPropertiesOption.Checked Then options = options + swInsertPartImportCustomToCutlistProperties
    End If
    If pm_cutlistPropertiesCheckBox.Checked Then options = options + swInsertPartImportCutListProperties
    If pm_coordinateSystemsCheckBox.Checked Then options = options + swInsertPartImportCoordinateSystem
    If pm_modelDimensionsCheckBox.Checked Then options = options + swInsertPartImportModelDimensions
    If pm_holeWizardDataCheckBox.Checked Then options = options + swInsertPartImportHoleWzdData
    If pm_sheetMetalInformationCheckBox.Checked Then
        options = options + swInsertPartImportSMInfo
        If pm_unlockedPropertiesCheckBox.Checked Then options = options + swInsertPartImportIndProps 'guess this is right
    End If
    If pm_breakLinkToOriginalPartCheckBox.Checked Then options = swInsertPartBreakLink

    #If ccDebug Then
        Debug.Print "GetInsertOptions", options
    #End If
    
    GetInsertOptions = options
    
End Function

Private Sub Class_Initialize()
    #If ccDebug Then
        Debug.Print "Class_Initialize"
    #End If
   
    'swPropertyManagerPageOptions_e
    
    Set pm_Page = swApp _
       .CreatePropertyManagerPage( _
        Title:="Insert Configurations of Part" _
      , Handler:=Me _
      , errors:=longErrors _
      , options:=swPropertyManagerOptions_LockedPage _
               + swPropertyManagerOptions_PushpinButton _
               + swPropertyManager_CancelButton _
               + swPropertyManager_OkayButton _
               + swPropertyManagerOptions_CanEscapeCancel _
        )
       pm_Page.SetTitleBitmap2 (swApp.GetCurrentMacroPathFolder & "\InsertPartConfigsTitle.bmp")
       
'    pm_Page.SetMessage3 "I really wanted a yellow background! :)" _
                    , swMessageBoxVisible, swMessageBoxMaintainExpandState, "Message"
    
    Set pm_messageGroup = pm_Page _
       .AddGroupBox( _
                       Control_ID.messageGroup _
                     , Caption:="Message" _
                     , options:=swGroupBoxOptions_Visible _
                              + swGroupBoxOptions_Expanded _
                    )
                    
      Set pm_messageLabel = pm_messageGroup _
       .AddControl2( _
                       Control_ID.messageLabel _
                     , swControlType_Label _
                     , Caption:="This Macro will allow you to insert multiple configurations of a part.  I have tried to make it similar to the native Insert Part property manager page, but there are some limitations, notably Locate Part is absent.  It is also untested, please use at your own risk!" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                    )
    pm_messageLabel.BackgroundColor = vbYellow
     
    Set pm_transferGroup = pm_Page _
       .AddGroupBox( _
                       Control_ID.transferGroup _
                     , Caption:="Transfer" _
                     , options:=swGroupBoxOptions_Visible _
                              + swGroupBoxOptions_Expanded _
                    )
                     
    Set pm_solidBodiesCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.solidBodiesCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Solid bodies" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                    )
                   
    Set pm_surfaceBodiesCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.surfaceBodiesCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Surface bodies" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
    
    Set pm_axesCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.axesCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Axes" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )

    Set pm_planesCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.planesCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Planes" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
                   
    Set pm_cosmeticThreadsCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.cosmeticThreadsCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Cosmetic threads" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
    
    Set pm_absorbedSketchesCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.absorbedSketchesCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Absorbed sketches" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
    
    Set pm_unabsorbedSketchesCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.unabsorbedSketchesCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Unabsorbed sketches" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
    
    Set pm_customPropertiesCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.customPropertiesCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Custom properties" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
    pm_customPropertiesCheckBox.Checked = False
                        
    Set pm_customToFilePropertiesOption = pm_transferGroup _
       .AddControl2( _
                       Control_ID.customToFilePropertiesOption _
                     , swControlType_Option _
                     , Caption:="File properties" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_Indent _
                     , options:=swControlOptions_Enabled _
                    )
    pm_customToFilePropertiesOption.Style = swPropMgrPageOptionStyle_FirstInGroup
    pm_customToFilePropertiesOption.Checked = False

    Set pm_customToToCutListPropertiesOption = pm_transferGroup _
       .AddControl2( _
                       Control_ID.customToToCutListPropertiesOption _
                     , swControlType_Option _
                     , Caption:="Cut list properties" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_Indent _
                     , options:=swControlOptions_Enabled _
                    )
      pm_customToToCutListPropertiesOption.Checked = True

    Set pm_cutlistPropertiesCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.cutlistPropertiesCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Cut list properties" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )

    Set pm_coordinateSystemsCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.coordinateSystemsCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Coordinate systems" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )

    Set pm_modelDimensionsCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.modelDimensionsCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Model dimensions" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )

    Set pm_holeWizardDataCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.holeWizardDataCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Hole wizard data" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )

    Set pm_sheetMetalInformationCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.sheetMetalInformationCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Sheet metal information" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                    )

    Set pm_unlockedPropertiesCheckBox = pm_transferGroup _
       .AddControl2( _
                       Control_ID.unlockedPropertiesCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Unlocked properties" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_Indent _
                     , options:=swControlOptions_Visible _
                   )

    Set pm_visualPropertiesGroup = pm_Page _
       .AddGroupBox( _
                       Control_ID.visualPropertiesGroup _
                     , Caption:="Visual Properties" _
                     , options:=swGroupBoxOptions_Visible _
                              + swGroupBoxOptions_Expanded _
                    )
    
    Set pm_propagateVisualPropertiesCheckBox = pm_visualPropertiesGroup _
       .AddControl2( _
                       Control_ID.propagateVisualPropertiesCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Propagate from original part" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
 
    Set pm_linkGroup = pm_Page _
       .AddGroupBox( _
                       Control_ID.linkGroup _
                     , Caption:="Link" _
                     , options:=swGroupBoxOptions_Visible _
                              + swGroupBoxOptions_Expanded _
                    )

    Set pm_breakLinkToOriginalPartCheckBox = pm_linkGroup _
       .AddControl2( _
                       Control_ID.breakLinkToOriginalPartCheckBox _
                     , swControlType_Checkbox _
                     , Caption:="Break link to original part" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
 
    Set pm_configurationsGroup = pm_Page _
       .AddGroupBox( _
                       Control_ID.configurationsGroup _
                     , Caption:="Configurations" _
                     , options:=swGroupBoxOptions_Visible _
                              + swGroupBoxOptions_Expanded _
                    )
    
      Set pm_configurationSelectionsINCLUDEOption = pm_configurationsGroup _
       .AddControl2( _
                       Control_ID.configurationSelectionsINCLUDEOption _
                     , swControlType_Option _
                     , Caption:="Include" _
                     , Tip:="Only selected configs will be used" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
      pm_configurationSelectionsINCLUDEOption.Style = swPropMgrPageOptionStyle_FirstInGroup
      pm_configurationSelectionsINCLUDEOption.Checked = False
      
      Set pm_configurationSelectionsEXCLUDEOption = pm_configurationsGroup _
       .AddControl2( _
                       Control_ID.configurationSelectionsEXCLUDEOption _
                     , swControlType_Option _
                     , Caption:="Exclude" _
                     , Tip:="All but the selected configs will be used" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
      pm_configurationSelectionsEXCLUDEOption.Checked = True

                   
      Set pm_configurationSelectionsLabel = pm_configurationsGroup _
       .AddControl2( _
                       Control_ID.configurationSelectionsLabel _
                     , swControlType_Label _
                     , Caption:="initialising!" _
                     , Tip:="" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
                 
    Set pm_configurationsListBox = pm_configurationsGroup _
       .AddControl2( _
                       Control_ID.configurationsListBox _
                     , swControlType_Listbox _
                     , Caption:="" _
                     , Tip:="CTRL-click to deselect last item!" _
                     , LeftAlign:=swControlAlign_LeftEdge _
                     , options:=swControlOptions_Visible _
                              + swControlOptions_Enabled _
                   )
    
    With pm_configurationsListBox
                .AddItems (configNames)
                .Height = 200
                .Style = swPropMgrPageListBoxStyle_MultipleItemSelect
    End With
                    
End Sub
Public Function Show() As Long
    #If ccDebug Then
        Debug.Print "Show"
    #End If

    Show = pm_Page.Show2(options:=0)
     
End Function

Private Sub PropertyManagerPage2Handler9_AfterActivation()
    #If ccDebug Then
        Debug.Print "AfterActivation"
    #End If
    Call UpdateConfigurationsLabel
End Sub

Private Sub PropertyManagerPage2Handler9_AfterClose()
    #If ccDebug Then
        Debug.Print "AfterClose"
    #End If

    ' Destroy the class
    Set pm = Nothing

End Sub


Private Sub PropertyManagerPage2Handler9_OnClose(ByVal Reason As Long)
    #If ccDebug Then
        Debug.Print "OnClose", Reason
    #End If

    If Reason = swPropertyManagerPageClose_Cancel Then

    ElseIf Reason = swPropertyManagerPageClose_Okay Then
        If GetSelectedConfigsCount = 0 Then
        
            ' no config selected..  not sure what the best behaviour is, just do this for now
             ' but maybe I should look at other ideas, especially if there is only 1 config
             MsgBox "No Configurations Selected"
        Else
            
            Dim options As Long: options = GetInsertOptions
            
            If options = 0 Then 'avoid a SW crash
                MsgBox "Nothing Selected to Insert"
                
            Else 'Lets DO this already
                Call InsertPart_UpdateExFileRefs_Loop(GetSelectedConfigs, options)
            End If
        End If
    End If
End Sub

Private Function PropertyManagerPage2Handler9_OnHelp() As Boolean
    #If ccDebug Then
        Debug.Print "OnHelp"
    #End If
    PropertyManagerPage2Handler9_OnHelp = False
End Function
Private Function PropertyManagerPage2Handler9_OnSubmitSelection(ByVal ID As Long, ByVal Selection As Object, ByVal SelType As Long, ItemText As String) As Boolean
    #If ccDebug Then
        Debug.Print "OnSubmitSelection", ID, Selection, SelType, ItemText
    #End If
    PropertyManagerPage2Handler9_OnSubmitSelection = True
End Function
Private Function PropertyManagerPage2Handler9_OnPreviousPage() As Boolean
    #If ccDebug Then
        Debug.Print "OnPreviousPage"
    #End If
End Function
Private Function PropertyManagerPage2Handler9_OnNextPage() As Boolean
    #If ccDebug Then
        Debug.Print "OnNextPage"
    #End If
End Function
Private Function PropertyManagerPage2Handler9_OnPreview() As Boolean
    #If ccDebug Then
        Debug.Print "OnPreview"
    #End If
End Function
Private Sub PropertyManagerPage2Handler9_OnWhatsNew()
    #If ccDebug Then
        Debug.Print "OnWhatsNew"
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnUndo()
    #If ccDebug Then
        Debug.Print "OnUndo"
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnRedo()
    #If ccDebug Then
        Debug.Print "OnRedo"
    #End If
End Sub
Private Function PropertyManagerPage2Handler9_OnTabClicked(ByVal ID As Long) As Boolean
    #If ccDebug Then
        Debug.Print "OnTabClicked", ID
    #End If
End Function
Private Sub PropertyManagerPage2Handler9_OnTextboxChanged(ByVal ID As Long, ByVal Text As String)
    #If ccDebug Then
        Debug.Print "OnTextboxChanged", ID, Text
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnGroupExpand(ByVal ID As Long, ByVal Expanded As Boolean)
    #If ccDebug Then
        Debug.Print "OnGroupExpand", ID, Expanded
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnGroupCheck(ByVal ID As Long, ByVal Checked As Boolean)
    #If ccDebug Then
        Debug.Print "OnGroupCheck", ID, Checked
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnCheckboxCheck(ByVal ID As Long, ByVal Checked As Boolean)
    #If ccDebug Then
        Debug.Print "OnCheckboxCheck", ID, Checked
    #End If
    
    Select Case ID
           Case Control_ID.customPropertiesCheckBox
                
                pm_customToFilePropertiesOption.Visible = Checked
                pm_customToToCutListPropertiesOption.Visible = Checked
    End Select
    
End Sub
Private Sub PropertyManagerPage2Handler9_OnOptionCheck(ByVal ID As Long)
    #If ccDebug Then
        Debug.Print "OnOptionCheck", ID
    #End If
    
    Select Case ID
           Case Control_ID.configurationSelectionsINCLUDEOption _
              , Control_ID.configurationSelectionsEXCLUDEOption
                
                Call UpdateConfigurationsLabel
    End Select
    
End Sub
Private Sub PropertyManagerPage2Handler9_OnPopupMenuItem(ByVal ID As Long)
    #If ccDebug Then
        Debug.Print "OnPopupMenuItem", ID
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnPopupMenuItemUpdate(ByVal ID As Long, retVal As Long)
    #If ccDebug Then
        Debug.Print "OnPopupMenuItemUpdate", ID, retVal
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnButtonPress(ByVal ID As Long)
    #If ccDebug Then
        Debug.Print "OnButtonPress", ID
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnNumberboxChanged(ByVal ID As Long, ByVal Value As Double)
    #If ccDebug Then
        Debug.Print "OnNumberboxChanged", ID, Value
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnComboboxEditChanged(ByVal ID As Long, ByVal Text As String)
    #If ccDebug Then
        Debug.Print "OnComboboxEditChanged", ID, Text
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnComboboxSelectionChanged(ByVal ID As Long, ByVal Item As Long)
    #If ccDebug Then
        Debug.Print "OnComboboxSelectionChanged", ID, Item
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnListboxSelectionChanged(ByVal ID As Long, ByVal Item As Long)
    #If ccDebug Then
        Debug.Print "OnListboxSelectionChanged", ID, Item
    #End If
    Call UpdateConfigurationsLabel
End Sub
Private Sub PropertyManagerPage2Handler9_OnSelectionboxCalloutCreated(ByVal ID As Long)
    #If ccDebug Then
        Debug.Print "OnSelectionboxCalloutCreated", ID
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnSelectionboxCalloutDestroyed(ByVal ID As Long)
    #If ccDebug Then
        Debug.Print "OnSelectionboxCalloutDestroyed", ID
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnSelectionboxFocusChanged(ByVal ID As Long)
    #If ccDebug Then
        Debug.Print "OnSelectionboxFocusChanged", ID
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnSelectionboxListChanged(ByVal ID As Long, ByVal count As Long)
    #If ccDebug Then
        Debug.Print "OnSelectionboxListChanged", ID, count
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnSliderPositionChanged(ByVal ID As Long, ByVal Value As Double)
    #If ccDebug Then
        Debug.Print "OnSliderPositionChanged", ID, Value
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnSliderTrackingCompleted(ByVal ID As Long, ByVal Value As Double)
    #If ccDebug Then
        Debug.Print "OnSliderTrackingCompleted", ID, Value
    #End If
End Sub
Private Function PropertyManagerPage2Handler9_OnActiveXControlCreated(ByVal ID As Long, ByVal Status As Boolean) As Long
    #If ccDebug Then
        Debug.Print "OnActiveXControlCreated", ID, Status
    #End If
End Function
Private Function PropertyManagerPage2Handler9_OnKeystroke(ByVal Wparam As Long, ByVal Message As Long, ByVal Lparam As Long, ByVal ID As Long) As Boolean
    #If ccDebug Then
        Debug.Print "OnKeystroke", Wparam, Message, Lparam, ID
    #End If
End Function
Private Sub PropertyManagerPage2Handler9_OnGainedFocus(ByVal ID As Long)
    #If ccDebug Then
        Debug.Print "OnGainedFocus", ID
    #End If
End Sub
Private Sub PropertyManagerPage2Handler9_OnLostFocus(ByVal ID As Long)
    #If ccDebug Then
        Debug.Print "OnLostFocus", ID
    #End If
End Sub
Public Function PropertyManagerPage2Handler9_OnWindowFromHandleControlCreated(ByVal ID As Long, ByVal Status As Boolean) As Long
    #If ccDebug Then
        Debug.Print "OnWindowFromHandleControlCreated", ID, Status
    #End If
End Function
Public Sub PropertyManagerPage2Handler9_OnListBoxRMBUp(ByVal ID As Long, ByVal posX As Long, ByVal posY As Long)
    #If ccDebug Then
        Debug.Print "OnListBoxRMBUp", ID, posX, posY
    #End If
End Sub
Public Sub PropertyManagerPage2Handler9_OnNumberboxTrackingCompleted(ByVal ID As Long, ByVal Value As Double)
    #If ccDebug Then
        Debug.Print "OnNumberboxTrackingCompleted", ID, Value
    #End If
End Sub

