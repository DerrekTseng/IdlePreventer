Imports System.Runtime.InteropServices
Imports System.ComponentModel

Namespace Hooks

    Public NotInheritable Class LocalKeyboardHook
        Implements IDisposable

        Public Event KeyDown As EventHandler(Of InputHelper.EventArgs.KeyboardHookEventArgs)

        Public Event KeyUp As EventHandler(Of InputHelper.EventArgs.KeyboardHookEventArgs)

        Private hHook As IntPtr = IntPtr.Zero

        Private HookProcedureDelegate As New NativeMethods.KeyboardProc(AddressOf HookCallback)

        Private Modifiers As ModifierKeys = ModifierKeys.None

        Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
            Dim Block As Boolean = False

            If nCode >= NativeMethods.HookCode.HC_ACTION Then
                Dim KeyCode As Keys = CType(wParam.ToInt32(), Keys)
                Dim KeyFlags As New NativeMethods.QWORD(lParam.ToInt64())
                Dim ScanCode As Byte = BitConverter.GetBytes(KeyFlags)(2) 'The scan code is the third byte in the integer (bits 16-23).
                Dim Extended As Boolean = (KeyFlags.LowWord.High And NativeMethods.KeyboardFlags.KF_EXTENDED) = NativeMethods.KeyboardFlags.KF_EXTENDED
                Dim AltDown As Boolean = (KeyFlags.LowWord.High And NativeMethods.KeyboardFlags.KF_ALTDOWN) = NativeMethods.KeyboardFlags.KF_ALTDOWN
                Dim KeyUp As Boolean = (KeyFlags.LowWord.High And NativeMethods.KeyboardFlags.KF_UP) = NativeMethods.KeyboardFlags.KF_UP

                'Set the ALT modifier if the KF_ALTDOWN flag is set.
                If AltDown = True _
                    AndAlso Internal.IsModifier(KeyCode, ModifierKeys.Alt) = False _
                     AndAlso (Me.Modifiers And ModifierKeys.Alt) <> ModifierKeys.Alt Then

                    Me.Modifiers = Me.Modifiers Or ModifierKeys.Alt
                End If

                'Raise KeyDown/KeyUp event.
                If KeyUp = False Then
                    Dim HookEventArgs As New InputHelper.EventArgs.KeyboardHookEventArgs(KeyCode, ScanCode, Extended, InputHelper.Hooks.KeyState.Down, Me.Modifiers)

                    If Internal.IsModifier(KeyCode, ModifierKeys.Control) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Control
                    If Internal.IsModifier(KeyCode, ModifierKeys.Shift) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Shift
                    If Internal.IsModifier(KeyCode, ModifierKeys.Alt) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Alt
                    If Internal.IsModifier(KeyCode, ModifierKeys.Windows) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Windows

                    RaiseEvent KeyDown(Me, HookEventArgs)
                    Block = HookEventArgs.Block
                Else
                    'Must be done before creating the HookEventArgs during KeyUp.
                    If Internal.IsModifier(KeyCode, ModifierKeys.Control) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Control
                    If Internal.IsModifier(KeyCode, ModifierKeys.Shift) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Shift
                    If Internal.IsModifier(KeyCode, ModifierKeys.Alt) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Alt
                    If Internal.IsModifier(KeyCode, ModifierKeys.Windows) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Windows

                    Dim HookEventArgs As New InputHelper.EventArgs.KeyboardHookEventArgs(KeyCode, ScanCode, Extended, InputHelper.Hooks.KeyState.Up, Me.Modifiers)

                    RaiseEvent KeyUp(Me, HookEventArgs)
                    Block = HookEventArgs.Block
                End If
            End If

            Return If(Block, New IntPtr(1), NativeMethods.CallNextHookEx(hHook, nCode, wParam, lParam))
        End Function
        Public Sub New()
            Me.New(NativeMethods.GetCurrentThreadId())
        End Sub

        Public Sub New(ByVal ThreadID As UInteger)
            hHook = NativeMethods.SetWindowsHookEx(NativeMethods.HookType.WH_KEYBOARD, HookProcedureDelegate, IntPtr.Zero, ThreadID)
            If hHook = IntPtr.Zero Then
                Dim Win32Error As Integer = Marshal.GetLastWin32Error()
                Throw New Win32Exception(Win32Error, "Failed to create local keyboard hook! (" & Win32Error & ")")
            End If
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
                If hHook <> IntPtr.Zero Then NativeMethods.UnhookWindowsHookEx(hHook)
            End If
            Me.disposedValue = True
        End Sub

        Protected Overrides Sub Finalize()
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(False)
            MyBase.Finalize()
        End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class
End Namespace