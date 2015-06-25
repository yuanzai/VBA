Attribute VB_Name = "ClickModule"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10

Sub ClickOnce(ByVal x As Long, y As Long)
    SetCursorPos x, y ' Sets the cursor position by x, y coordinates
    LeftClick 1 ' Click once
End Sub


Sub LeftClick(ByVal times As Long, Optional ByVal delay_in_ms As Long = 10)
    Dim i As Long
    For i = 1 To times
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        If times > 1 Then Sleep delay_in_ms
    Next
End Sub
