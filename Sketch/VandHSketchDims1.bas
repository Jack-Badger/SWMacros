Attribute VB_Name = "VandHSketchDims1"
' VandHSketchDims1.1
' SW2016
' Adds Horizontal and Vertical Dimensions to two sketch points in an active sketch
' If somethings not right this macro just exits.
' you must be in an active sketch and have exactly 2 sketch points selected
' it positions the dims centrally to the top and trys to choose the best side
' inspired by https://forum.solidworks.com/thread/228935
' by 369 6feb19

Option Explicit

Public SketchPoints(1 To 2) As SldWorks.SketchPoint

Sub main()

    Dim swmodel As SldWorks.ModelDoc2
    Set swmodel = Application.SldWorks.ActiveDoc
    
    If swmodel Is Nothing Then Exit Sub
  With swmodel
    If .SketchManager.ActiveSketch Is Nothing Then Exit Sub
    
   With .SelectionManager
    
    If .GetSelectedObjectCount <> 2 Then Exit Sub
    
    If .GetSelectedObjectType3(1, -1) <> swSelSKETCHPOINTS _
    Or .GetSelectedObjectType3(2, -1) <> swSelSKETCHPOINTS Then Exit Sub
    
    Set SketchPoints(1) = .GetSelectedObject6(1, -1)
    Set SketchPoints(2) = .GetSelectedObject6(2, -1)
    
    End With ' .SelectionManager
    
    Call .AddHorizontalDimension2(AvgX, MaxY, 0#)
    
    Call SketchPoints(1).Select(appendflag:=False)
    Call SketchPoints(2).Select(appendflag:=True)
    
    Call .AddVerticalDimension2(IIf((Left), MinX, MaxX), AvgY, 0#)
    
    .ClearSelection2 True
   End With ' swModel
End Sub

Private Function TopSkPt() As SldWorks.SketchPoint: Set TopSkPt = IIf(SketchPoints(1).y > SketchPoints(2).y, SketchPoints(1), SketchPoints(2)): End Function
Private Function BtmSkPt() As SldWorks.SketchPoint: Set BtmSkPt = IIf(SketchPoints(1).y < SketchPoints(2).y, SketchPoints(1), SketchPoints(2)): End Function
Private Function MaxX() As Double:  MaxX = IIf(SketchPoints(1).x > SketchPoints(2).x, SketchPoints(1).x, SketchPoints(2).x): End Function
Private Function MinX() As Double:  MinX = IIf(SketchPoints(1).x < SketchPoints(2).x, SketchPoints(1).x, SketchPoints(2).x): End Function
Private Function MaxY() As Double:  MaxY = IIf(SketchPoints(1).y > SketchPoints(2).y, SketchPoints(1).y, SketchPoints(2).y): End Function
Private Function AvgX() As Double:  AvgX = (SketchPoints(1).x + SketchPoints(2).x) / 2: End Function
Private Function AvgY() As Double:  AvgY = (SketchPoints(1).y + SketchPoints(2).y) / 2: End Function
Private Function Left() As Boolean: Left = IIf(TopSkPt.x < BtmSkPt.x, True, False): End Function

