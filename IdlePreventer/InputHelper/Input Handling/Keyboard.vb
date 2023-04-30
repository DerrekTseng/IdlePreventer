Imports System.Runtime.InteropServices

''' <summary>
''' A static class for handling and simulating physical keyboard input.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Keyboard
    Private Sub New()
    End Sub

#Region "Public methods"

#Region "IsKeyDown()"
    ''' <summary>
    ''' Checks whether the specified key is currently held down.
    ''' </summary>
    ''' <param name="Key">The key to check.</param>
    ''' <remarks></remarks>
    Public Shared Function IsKeyDown(ByVal Key As Keys) As Boolean
        Dim Modifiers As Keys() = Internal.ExtractModifiers(Key)
        For Each Modifier As Keys In Modifiers
            If (NativeMethods.GetAsyncKeyState(Modifier) And Constants.KeyDownBit) <> Constants.KeyDownBit Then
                Return False
            End If
        Next

        If Key = Keys.None Then Return True 'All modifiers are held down, no more keys left to check.

        Return (NativeMethods.GetAsyncKeyState(Key) And Constants.KeyDownBit) = Constants.KeyDownBit
    End Function
#End Region

#Region "IsKeyUp()"
    ''' <summary>
    ''' Checks whether the specified key is currently NOT held down.
    ''' </summary>
    ''' <param name="Key">The key to check.</param>
    ''' <remarks></remarks>
    Public Shared Function IsKeyUp(ByVal Key As Keys) As Boolean
        Return Keyboard.IsKeyDown(Key) = False
    End Function
#End Region

#Region "PressKey()"
    ''' <summary>
    ''' Simulates a keystroke.
    ''' </summary>
    ''' <param name="Key">The key to press.</param>
    ''' <param name="HardwareKey">Whether to simulate the keystroke using its virtual key code or its hardware scan code.</param>
    ''' <remarks></remarks>
    Public Shared Sub PressKey(ByVal Key As Keys, Optional ByVal HardwareKey As Boolean = False)
        Keyboard.SetKeyState(Key, True, HardwareKey)
        Keyboard.SetKeyState(Key, False, HardwareKey)
    End Sub
#End Region

#Region "SetKeyState()"
    ''' <summary>
    ''' Simulates a key being pushed down or released.
    ''' </summary>
    ''' <param name="Key">The key which to simulate.</param>
    ''' <param name="KeyDown">Whether to push down or release the key.</param>
    ''' <param name="HardwareKey">Whether to simulate the event using the key's virtual key code or its hardware scan code.</param>
    ''' <remarks></remarks>
    Public Shared Sub SetKeyState(ByVal Key As Keys, ByVal KeyDown As Boolean, Optional ByVal HardwareKey As Boolean = False)
        Dim InputList As New List(Of NativeMethods.INPUT)
        Dim Modifiers As Keys() = Internal.ExtractModifiers(Key)

        For Each Modifier As Keys In Modifiers
            InputList.Add(Keyboard.GetKeyboardInputStructure(Modifier, KeyDown, HardwareKey))
        Next
        InputList.Add(Keyboard.GetKeyboardInputStructure(Key, KeyDown, HardwareKey))

        NativeMethods.SendInput(CType(InputList.Count, UInteger), InputList.ToArray(), Marshal.SizeOf(GetType(NativeMethods.INPUT)))
    End Sub
#End Region

#End Region

#Region "Internal methods"

#Region "GetKeyboardInputStructure()"
    ''' <summary>
    ''' Constructs a native keyboard INPUT structure that can be passed to SendInput().
    ''' </summary>
    ''' <param name="Key">The key to send.</param>
    ''' <param name="KeyDown">Whether to send a KeyDown/KeyUp stroke.</param>
    ''' <param name="HardwareKey">Whether to send the key's hardware scan code instead of its virtual key code.</param>
    ''' <remarks></remarks>
    Private Shared Function GetKeyboardInputStructure(ByVal Key As Keys, ByVal KeyDown As Boolean, ByVal HardwareKey As Boolean) As NativeMethods.INPUT
        Dim KeyboardInput As New NativeMethods.KEYBDINPUT With {
            .wVk = If(HardwareKey = False, (Key And UShort.MaxValue), 0),
            .wScan = If(HardwareKey,
                        NativeMethods.MapVirtualKeyEx(CType(Key, UInteger), 0, NativeMethods.GetKeyboardLayout(0)),
                        0) And UShort.MaxValue,
            .time = 0,
            .dwFlags = If(HardwareKey, NativeMethods.KEYEVENTF.SCANCODE, 0) Or
                        If(KeyDown = False, NativeMethods.KEYEVENTF.KEYUP, 0),
            .dwExtraInfo = UIntPtr.Zero
        }

        Dim Union As New NativeMethods.INPUTUNION With {.ki = KeyboardInput}
        Dim Input As New NativeMethods.INPUT With {
            .type = NativeMethods.INPUTTYPE.KEYBOARD,
            .U = Union
        }

        Return Input
    End Function
#End Region

#End Region

End Class