Attribute VB_Name = "NameFacesStuff1"
' NameFacesStuff.swp
' rob 26 Jan 2021
'
' Gives a face standard name

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Const StandardFaceName = "Face2,Edge2,Face1,Edge1,End2,End1"

Sub main()

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        Exit Sub
    End If
    
    If swModel.GetType = swDocumentTypes_e.swDocPART Then
        Dim swPart As SldWorks.PartDoc
        Set swPart = swModel
    Else
        Exit Sub
    End If
    
    Dim vBodies As Variant
    
    vBodies = swPart.GetBodies2(BodyType:=swBodyType_e.swSheetBody, BVisibleOnly:=False)
    
    Dim swbody As SldWorks.Body2
    
    Dim i As Long
    For i = LBound(vBodies) To UBound(vBodies)
        Set swbody = vBodies(i)
        
        If swbody.name = "Box" Then Exit For
        
    Next i
    
    
    Debug.Print swbody.name
    

    Dim swEntity As SldWorks.Entity
    
    Dim vFaceNames As Variant
    vFaceNames = Split(StandardFaceName, ",")
    
    Dim swFace As SldWorks.Face2
    Set swFace = swbody.GetFirstFace
    
    Dim changedNamesCount As Long
    
    changedNamesCount = 0
    
    For i = 0 To 5
    
        Set swEntity = swFace
           
        swEntity.Select False
        swPart.DeleteEntityName swEntity
        If swPart.SetEntityName(swEntity, vFaceNames(i)) Then
            changedNamesCount = changedNamesCount + 1
        End If
        
        Set swFace = swFace.GetNextFace
        
        Debug.Print swEntity.ModelName
    Next i
    
    swApp.SendMsgToUser2 "Created " & changedNamesCount & " face names!", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
    
    Dim userResponse As swMessageBoxResult_e
    
    userResponse = swApp.SendMsgToUser2("Create Selection Set?", swMessageBoxIcon_e.swMbQuestion, swMessageBoxBtn_e.swMbYesNo)
    
    If userResponse = swMbHitYes Then
    
        selectfacesandcreateselectionset swbody, "Face1,Face2,Edge1,Edge2,End1,End2"
    
    End If
    

End Sub
    
Private Sub selectfacesandcreateselectionset(swbody As SldWorks.Body2, namesCSV As String)

    
    Dim nameArray() As String
    
    nameArray = Split(namesCSV, ",")
    
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = swModel
    
    Dim vName As Variant
    
    Dim swEntity As SldWorks.Entity
    
    swModel.ClearSelection2 All:=True
    
    For Each vName In nameArray
    
        Set swEntity = swPart.GetEntityByName(CStr(vName), swSelectType_e.swSelFACES)
    
        swEntity.Select appendFlag:=True
    Next
    
    Dim swSelectionSet As SldWorks.SelectionSet
    
    Dim status As Long
    
    Set swSelectionSet = swModel.Extension.SaveSelection(status)
        

End Sub


