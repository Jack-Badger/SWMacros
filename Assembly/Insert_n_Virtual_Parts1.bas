Attribute VB_Name = "Insert_n_Virtual_Parts1"
'SOLIDWORKS 2016
' Inserts and fixes/unfixes virtual components in an assembly
' by rob e

' additional code for key detection based on
' https://forum.solidworks.com/thread/213575 by Simon Turner

' keycode reference https://msdn.microsoft.com/en-us/library/windows/desktop/dd375731(v=vs.85).aspx
' what is &H? reference http://www.vbforums.com/showthread.php?398006-RESOLVED-amp-H-Values
' also for further study on getting *all* keys see
' https://stackoverflow.com/questions/13127762/what-is-keys0-after-getkeyboardstate-keys0-in-vba


Option Explicit

#If VBA7 Then
    Declare PtrSafe Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
#Else
    Declare Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
#End If

'Please change these as you please
Const DEFAULT_COUNT = 10
Const MAX_COUNT = 50

'but not these
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swAssem As SldWorks.AssemblyDoc
Dim swComponent As SldWorks.Component2

Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

'https://msdn.microsoft.com/en-us/library/windows/desktop/dd375731(v=vs.85).aspx

Public Enum Key_Code
      SHIFT = &H10
       CTRL = &H11
       KEY0 = &H30
       KEY1 = &H31
       KEY2 = &H32
       KEY3 = &H33
       KEY4 = &H34
       KEY5 = &H35
       KEY6 = &H36
       KEY7 = &H37
       KEY8 = &H38
       KEY9 = &H39
    NUMPAD0 = &H60
    NUMPAD1 = &H61
    NUMPAD2 = &H62
    NUMPAD3 = &H63
    NUMPAD4 = &H64
    NUMPAD5 = &H65
    NUMPAD6 = &H66
    NUMPAD7 = &H67
    NUMPAD8 = &H68
    NUMPAD9 = &H69

End Enum


Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    Select Case True
           Case (swModel Is Nothing) _
              , (swModel.GetType <> swDocASSEMBLY)
                'Debug.Print "QUIT"
           Case Else
                'Debug.Print "OK"
                Set swAssem = swModel
                Call InsertLoop
                swModel.ClearSelection2 True
    End Select
    'Debug.Print "End"
End Sub

Private Function InsertLoop() As Integer
    With swAssem
        Dim count As Integer: count = GetCount()
        If count > 0 Then

            Dim c As Integer
            For c = 1 To count
                'Debug.Print "+"
               .InsertNewVirtualPart Nothing, swComponent
                swComponent.Select appendflag:=False
               .FixComponent
               '.UnfixComponent
            Next

        End If
    End With
    InsertLoop = count
End Function

Private Function GetCount() As Integer
    ' returns the value of the key currently being pressed
    ' if 0 is pressed counts as 10
    ' if SHIFT is pressed displays an Input Box to User
    ' ***REMOVED-if CTRL is pressed the result is doubled ***
    ' if no keys pressed returns DEFAULT
    ' Return value is clipped by:  1 <= result <= MAX

    GetCount = DEFAULT_COUNT
    Select Case True
           Case (GetKeyState(Key_Code.SHIFT) < 0): GetCount = AskUser
           Case (GetKeyState(Key_Code.KEY1) < 0), (GetKeyState(Key_Code.NUMPAD1) < 0): GetCount = 1
           Case (GetKeyState(Key_Code.KEY2) < 0), (GetKeyState(Key_Code.NUMPAD2) < 0): GetCount = 2
           Case (GetKeyState(Key_Code.KEY3) < 0), (GetKeyState(Key_Code.NUMPAD3) < 0): GetCount = 3
           Case (GetKeyState(Key_Code.KEY4) < 0), (GetKeyState(Key_Code.NUMPAD4) < 0): GetCount = 4
           Case (GetKeyState(Key_Code.KEY5) < 0), (GetKeyState(Key_Code.NUMPAD5) < 0): GetCount = 5
           Case (GetKeyState(Key_Code.KEY6) < 0), (GetKeyState(Key_Code.NUMPAD6) < 0): GetCount = 6
           Case (GetKeyState(Key_Code.KEY7) < 0), (GetKeyState(Key_Code.NUMPAD7) < 0): GetCount = 7
           Case (GetKeyState(Key_Code.KEY8) < 0), (GetKeyState(Key_Code.NUMPAD8) < 0): GetCount = 8
           Case (GetKeyState(Key_Code.KEY9) < 0), (GetKeyState(Key_Code.NUMPAD9) < 0): GetCount = 9
           Case (GetKeyState(Key_Code.KEY0) < 0), (GetKeyState(Key_Code.NUMPAD0) < 0): GetCount = 10
          'Case (GetKeyState(Key_Code.CTRL) < 0): GetCount = GetCount + GetCount 'not a good idea
           Case Else: 'Debug.Print "No Keys"
    End Select
    'Debug.Print "GetCount"; GetCount
End Function

Private Function AskUser() As Integer
    Dim sInput As String
    sInput = InputBox("Please Choose from 1 to " & MAX_COUNT & "." & vbCrLf & _
                      " (To change defaults edit the macro)" _
                    , "Insert Virtual Parts", DEFAULT_COUNT)

    Select Case True
           Case Not (IsNumeric(sInput)): AskUser = 0
           Case (sInput > MAX_COUNT): AskUser = MAX_COUNT
           Case (sInput < 1): AskUser = 1
           Case Else: AskUser = sInput
    End Select
End Function


