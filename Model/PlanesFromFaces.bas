Attribute VB_Name = "PlanesFromFaces"
' SOLIDWORKS 2016
' PlanesFromFaces
' by 369
'
' Inspired because I'm sick of SW not remembering my default choice
'
' PreConditions
'                   A ModelDoc is Active
'                   Some Faces Selected
'
' Post Conditions
'                   A coincident plane is created for every highlighted planar face
'
' Limitations*
'
' TODO              Some way of naming them efficiently
'                   All Faces option
'                   Select by body option

Option Explicit
    Dim swApp       As SldWorks.SldWorks
    Dim faces  As VBA.Collection

Sub main()
    Set swApp = Application.SldWorks
    With swApp: Select Case True
       Case .ActiveDoc Is Nothing
       Case .ActiveDoc.GetType = swDocumentTypes_e.swDocASSEMBLY _
          , .ActiveDoc.GetType = swDocumentTypes_e.swDocPART
                Call ProcessModel(.ActiveDoc)
    End Select: End With
End Sub

Private Function ProcessModel(swModel As SldWorks.ModelDoc2)
    Call ProcessSelections(swModel)
    Call CreatePlanes(swModel.FeatureManager)
End Function
               
Private Function ProcessSelections(swModel As SldWorks.ModelDoc2): With swModel

GoSub GetFaces
Exit Function

GetFaces:
        Set faces = CollectTheseTypes(.SelectionManager, swSelFACES)
        'TODO MAYB ?? If faces.Count = 0 Then Set faces = CollectAllFaces(swModel)
        Return
        
End With: End Function
    
Private Function _
        CollectTheseTypes(swSelMgr As SldWorks.SelectionMgr, selType As Long) _
     As VBA.Collection
   With swSelMgr
    
    Set CollectTheseTypes = New VBA.Collection
        Dim i As Integer
        For i = 1 To .GetSelectedObjectCount2(Mark:=-1)
            If .GetSelectedObjectType3(index:=i, Mark:=-1) = selType Then
                Select Case selType
                    'Case swSelCONFIGURATIONS
                         'CollectTheseTypes.Add .GetSelectedObject6(index:=i, Mark:=-1)
                    'Case swSelCOMPONENTS
                         'CollectTheseTypes.Add .GetSelectedObject6(index:=i, Mark:=-1).GetID
                    Case swSelFACES
                         CollectTheseTypes.Add .GetSelectedObject6(index:=i, Mark:=-1)
            End Select
            End If
        Next
    End With
End Function

Private Function CreatePlanes(swFeatMgr As SldWorks.FeatureManager)
' Marks for selection
' 0 = First reference entity
' 1 = Second reference entity
' 2 = Third reference entity

' swRefPlaneReferenceConstraints_e Enum
' swRefPlaneReferenceConstraint_Angle 16 or 0x10
' swRefPlaneReferenceConstraint_Coincident 4 or 0x4
' swRefPlaneReferenceConstraint_Distance 8 or 0x8
' swRefPlaneReferenceConstraint_MidPlane 128 or 0x80
' swRefPlaneReferenceConstraint_OptionFlip 256 or 0x100
' swRefPlaneReferenceConstraint_OptionOriginOnCurve 512 or 0x200
' swRefPlaneReferenceConstraint_OptionProjectAlongSketchNormal 2056 or 0x800
' swRefPlaneReferenceConstraint_OptionProjectToNearestLocation 1028 or 0x400
' swRefPlaneReferenceConstraint_Parallel 1 or 0x1
' swRefPlaneReferenceConstraint_ParallelToScreen 4096 or 0x1000
' swRefPlaneReferenceConstraint_Perpendicular 2 or 0x2
' swRefPlaneReferenceConstraint_Project 64 or 0x40
' swRefPlaneReferenceConstraint_Tangent 32 or 0x20

    Dim swFeature As SldWorks.Feature
    Dim swFace As SldWorks.Entity
    
    Dim vNormal As Variant
    Dim i As Integer: For i = faces.Count To 1 Step -1
        Set swFace = faces.Item(i)
        vNormal = swFace.Normal
        If vNormal(0) = 0 And vNormal(1) = 0 And vNormal(2) = 0 Then
            ' ignore non planar surface
        Else
            If swFace.Select2(append:=False, Mark:=0) Then
             Set swFeature = swFeatMgr.InsertRefPlane(swRefPlaneReferenceConstraint_Coincident, 0, 0, 0, 0, 0)
            End If
        End If
    Next
End Function


