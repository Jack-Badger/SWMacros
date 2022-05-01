Attribute VB_Name = "InsertPartConfigs1"
Option Explicit
Public swApp As SldWorks.SldWorks
Public activePart As SldWorks.ModelDoc2
Dim derivedPart As SldWorks.PartDoc
Public pm As InsertPartConfigsPMPage
Public filename As String
Private insertFeature As SldWorks.Feature
Public configNames As Variant
Private configName As Variant
Private derivedPartInitialConfiguration As SldWorks.Configuration
Public hasCutlistProperties As Boolean
Public hasSheetMetalProperties As Boolean

Dim boolStatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub Main()
    Set swApp = Application.SldWorks
    Set activePart = swApp.activeDoc
    Select Case True
           Case (activePart Is Nothing), (activePart.GetType <> swDocPART)
                Call swApp.SendMsgToUser2("This macro requires an active part", swMbStop, swMbOk)
           Case Else
                If GetDerivedPartFromUserAndShiz Then
                    Set pm = New InsertPartConfigsPMPage
                        pm.Show
                End If
     End Select
End Sub

Private Function GetDerivedPartFromUserAndShiz() As Boolean

    GetDerivedPartFromUserAndShiz = False 'Assume the worst
    
    filename = swApp.GetOpenFileName("Open Part", "", "Part Files (*.sldprt)|*.sldprt", vbNull, "", "")
    
    If filename = "" Then Exit Function 'User probly hit Cancel

    Set derivedPart = swApp.OpenDoc6(filename, swDocPART, swOpenDocOptions_Silent, "", longstatus, longwarnings)
    
    
    If derivedPart Is Nothing Then 'There seems to be a lot more to error checking than I'm prepared to go into, lets just see how this fares?

        Select Case longstatus
           Case swFileLoadError_e.swFileNotFoundError
                MsgBox "Unable to locate the file"
           Case swFileLoadError_e.swFileWithSameTitleAlreadyOpen 'Prolly need to look at this one.. I did it's OK
                MsgBox "A document with the same name is already open"
           Case swFileLoadError_e.swFutureVersion
                MsgBox "The document was saved in a future version of SOLIDWORKS"
           Case swFileLoadError_e.swGenericError
                MsgBox "An unspecified error was encountered"
           Case swFileLoadError_e.swInvalidFileTypeError
                MsgBox "File type argument is not valid"
           Case swFileLoadError_e.swLiquidMachineDoc
                MsgBox "File encrypted by Liquid Machines"
           Case swFileLoadError_e.swLowResourcesError
                MsgBox "File is open and blocked because the system memory is low, or the number of GDI handles has exceeded the allowed maximum"
           Case swFileLoadError_e.swNoDisplayData
                MsgBox "File contains no display data"
           Case Else
                MsgBox "Well that didn't work!"
        End Select
        
        Exit Function
        
    Else ' likely we have a valid PartDoc
    
            Set derivedPartInitialConfiguration = derivedPart.ConfigurationManager.ActiveConfiguration
            configNames = derivedPart.GetConfigurationNames
            hasCutlistProperties = derivedPart.IsWeldment
            hasSheetMetalProperties = (derivedPart.GetBendState <> swSMBendStateNone)
    
            Call swApp.ActivateDoc3(Name:=activePart.GetTitle, UseUserPreferences:=False, Option:=swDontRebuildActiveDoc, errors:=longstatus)
            
            If longstatus = swActivateDocError_e.swGenericActivateError Then
                MsgBox " An unspecified error was encountered, and the document was not activated"
            Else
                GetDerivedPartFromUserAndShiz = True 'HOORAY!
            End If
       
    End If

End Function

Public Sub InsertPart_UpdateExFileRefs_Loop(selectedConfigs As Variant, insertOptions As Long) ' updateOptions As Long)
 
    For Each configName In selectedConfigs
         Set insertFeature = activePart.InsertPart2(filename, insertOptions)
        Call insertFeature.UpdateExternalFileReferences(swExternalFileReferencesNamedConfig _
                                                      , CStr(configName) _
                                                      , swExternalFileReferencesUpdateNone)
    Next configName
    
    Set insertFeature = Nothing: configName = vbNull

End Sub

'Public Sub SwitchConfig_InsertPart_Loop(ByVal selectedConfigs As Variant, ByVal insertOptions As Long)
'Dim configName As Variant
'For Each configName In selectedConfigs
'    Set insertFeature = activePart.InsertPart2(filename, insertOptions)
'Next configName
'Set insertFeature = Nothing

'End Sub



