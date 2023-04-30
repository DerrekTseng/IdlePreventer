Imports System.Runtime.InteropServices

''' <summary>
''' A static class for handling and simulating physical mouse input.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Mouse
    Private Sub New()
    End Sub

#Region "Public methods"

#Region "IsButtonDown()"
    ''' <summary>
    ''' Checks whether the specified mouse button is currently held down.
    ''' </summary>
    ''' <param name="Button">The mouse button to check.</param>
    ''' <remarks></remarks>
    Public Shared Function IsButtonDown(ByVal Button As MouseButtons) As Boolean
        Dim Key As Keys = Keys.None
        Select Case Button
            Case MouseButtons.Left : Key = Keys.LButton
            Case MouseButtons.Middle : Key = Keys.MButton
            Case MouseButtons.Right : Key = Keys.RButton
            Case MouseButtons.XButton1 : Key = Keys.XButton1
            Case MouseButtons.XButton2 : Key = Keys.XButton2
            Case Else
                Throw New ArgumentException("Invalid mouse button " & Button.ToString() & "!", "Button")
        End Select

        Return Keyboard.IsKeyDown(Key)
    End Function
#End Region

#Region "IsButtonUp()"
    ''' <summary>
    ''' Checks whether the specified mouse button is currently NOT held down.
    ''' </summary>
    ''' <param name="Button">The mouse button to check.</param>
    ''' <remarks></remarks>
    Public Shared Function IsButtonUp(ByVal Button As MouseButtons) As Boolean
        Return Mouse.IsButtonDown(Button) = False
    End Function
#End Region

#Region "PressButton()"
    ''' <summary>
    ''' Simulates a mouse button click.
    ''' </summary>
    ''' <param name="Button">The button to press.</param>
    ''' <remarks></remarks>
    Public Shared Sub PressButton(ByVal Button As MouseButtons)
        Mouse.SetButtonState(Button, True)
        Mouse.SetButtonState(Button, False)
    End Sub
#End Region

#Region "SetButtonState()"
    ''' <summary>
    ''' Simulates a mouse button being pushed down or released.
    ''' </summary>
    ''' <param name="Button">The button which to simulate.</param>
    ''' <param name="MouseDown">Whether to push down or release the mouse button.</param>
    ''' <remarks></remarks>
    Public Shared Sub SetButtonState(ByVal Button As MouseButtons, ByVal MouseDown As Boolean)
        Dim InputList As NativeMethods.INPUT() = _
            New NativeMethods.INPUT(1 - 1) {Mouse.GetMouseClickInputStructure(Button, MouseDown)}
        NativeMethods.SendInput(CType(InputList.Length, UInteger), InputList, Marshal.SizeOf(GetType(NativeMethods.INPUT)))
    End Sub
#End Region

#End Region

#Region "Internal methods"

#Region "GetMouseClickInputStructure()"
    ''' <summary>
    ''' Constructs a native mouse INPUT structure for click events that can be passed to SendInput().
    ''' </summary>
    ''' <param name="Button">The button of the event.</param>
    ''' <param name="MouseDown">Whether to push down or release the mouse button.</param>
    ''' <remarks></remarks>
    Private Shared Function GetMouseClickInputStructure(ByVal Button As MouseButtons, ByVal MouseDown As Boolean) As NativeMethods.INPUT
        Dim Position As Point = Cursor.Position
        Dim MouseFlags As NativeMethods.MOUSEEVENTF
        Dim MouseData As UInteger = 0

        Select Case Button
            Case MouseButtons.Left : MouseFlags = If(MouseDown, NativeMethods.MOUSEEVENTF.LEFTDOWN, NativeMethods.MOUSEEVENTF.LEFTUP)
            Case MouseButtons.Middle : MouseFlags = If(MouseDown, NativeMethods.MOUSEEVENTF.MIDDLEDOWN, NativeMethods.MOUSEEVENTF.MIDDLEUP)
            Case MouseButtons.Right : MouseFlags = If(MouseDown, NativeMethods.MOUSEEVENTF.RIGHTDOWN, NativeMethods.MOUSEEVENTF.RIGHTUP)
            Case MouseButtons.XButton1
                MouseFlags = If(MouseDown, NativeMethods.MOUSEEVENTF.XDOWN, NativeMethods.MOUSEEVENTF.XUP)
                MouseData = NativeMethods.MouseXButton.XBUTTON1
            Case MouseButtons.XButton2
                MouseFlags = If(MouseDown, NativeMethods.MOUSEEVENTF.XDOWN, NativeMethods.MOUSEEVENTF.XUP)
                MouseData = NativeMethods.MouseXButton.XBUTTON2
            Case Else
                Throw New ArgumentException("Invalid mouse button " & Button.ToString() & "!", "Button")
        End Select

        Dim MouseInput As New NativeMethods.MOUSEINPUT With {
            .dx = Position.X,
            .dy = Position.Y,
            .mouseData = MouseData,
            .dwFlags = MouseFlags,
            .time = 0,
            .dwExtraInfo = UIntPtr.Zero
        }

        Dim Union As New NativeMethods.INPUTUNION With {.mi = MouseInput}
        Dim Input As New NativeMethods.INPUT With {
            .type = NativeMethods.INPUTTYPE.MOUSE,
            .U = Union
        }

        Return Input
    End Function
#End Region

#End Region

End Class