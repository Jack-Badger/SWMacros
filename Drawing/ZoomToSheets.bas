Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel
    
    Dim vSheetNames As Variant
    vSheetNames = swDraw.GetSheetNames
    
    Dim i As Long
    
    For i = 0 To UBound(vSheetNames)
    
        swDraw.ActivateSheet vSheetNames(i)
        swModel.Extension.ViewZoomToSheet
    
    Next i

End Sub
