Module MouseControl
    Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Int32, ByVal dx As Int32, ByVal dy As Int32, ByVal cButtons As Int32, ByVal dwExtraInfo As Int32)
    Private Declare Sub GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI)
    Private Const MOUSEEVENTF_LEFTDOWN = &H2
    Private Const MOUSEEVENTF_LEFTUP = &H4
    Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Private Const MOUSEEVENTF_MIDDLEUP = &H40
    Private Const MOUSEEVENTF_MIDDLEWHEEL = &H800
    Private Const MOUSEEVENTF_RIGHTDOWN = &H8
    Private Const MOUSEEVENTF_RIGHTUP = &H10
    Private Const MOUSEEVENTF_MOVE = &H1
    Private Const MOUSEEVENTF_ABSOLUTE = &H8000

    Private Structure POINTAPI
        Dim x As Int32
        Dim y As Int32
    End Structure

    Sub SetMouseLocation(ByVal Location As Point) '移動滑鼠到指定作標
        Call mouse_event(MOUSEEVENTF_MOVE Or MOUSEEVENTF_ABSOLUTE, Location.X * 65535 / Screen.PrimaryScreen.Bounds.Width + 1, Location.Y * 65535 / Screen.PrimaryScreen.Bounds.Height + 1, 0, 0)
    End Sub

    Function GetMouseLocation() As Point '取得滑鼠坐標
        Dim MousePoint As POINTAPI
        GetCursorPos(MousePoint)
        Return New Point(MousePoint.x, MousePoint.y)
    End Function

    Sub LeftButtonDown()
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    End Sub

    Sub LeftButtonUp()
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Sub MiddleButtonDown()
        Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)
    End Sub

    Sub MiddleButtonUp()
        Call mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
    End Sub

    Sub MiddleWhell(ByVal Value As Integer)
        Call mouse_event(MOUSEEVENTF_MIDDLEWHEEL, 0, 0, Value, 0)
    End Sub

    Sub RightButtonDown() '滑鼠右鍵按下
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    End Sub

    Sub RightButtonUp() '滑鼠右鍵彈起
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
    End Sub

End Module
