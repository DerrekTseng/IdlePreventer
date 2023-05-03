Imports System.Runtime.InteropServices
Imports System.ComponentModel

Namespace Hooks

    Public NotInheritable Class KeyboardHook
        Implements IDisposable

        Public Event KeyDown As EventHandler(Of InputHelper.EventArgs.KeyboardHookEventArgs)

        Public Event KeyUp As EventHandler(Of InputHelper.EventArgs.KeyboardHookEventArgs)

        Private hHook As IntPtr = IntPtr.Zero

        Private HookProcedureDelegate As New NativeMethods.LowLevelKeyboardProc(AddressOf HookCallback)

        Private Modifiers As ModifierKeys = ModifierKeys.None

        Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
            Dim Block As Boolean = False

            If nCode >= NativeMethods.HookCode.HC_ACTION AndAlso _
                (wParam = NativeMethods.KeyMessage.WM_KEYDOWN OrElse _
                 wParam = NativeMethods.KeyMessage.WM_KEYUP OrElse _
                 wParam = NativeMethods.KeyMessage.WM_SYSKEYDOWN OrElse _
                 wParam = NativeMethods.KeyMessage.WM_SYSKEYUP) Then

                Dim KeystrokeInfo As NativeMethods.KBDLLHOOKSTRUCT = _
                    CType(Marshal.PtrToStructure(lParam, GetType(NativeMethods.KBDLLHOOKSTRUCT)), NativeMethods.KBDLLHOOKSTRUCT)

                Select Case wParam

                    Case NativeMethods.KeyMessage.WM_KEYDOWN, NativeMethods.KeyMessage.WM_SYSKEYDOWN
                        Dim HookEventArgs As InputHelper.EventArgs.KeyboardHookEventArgs = Me.CreateEventArgs(True, KeystrokeInfo)

                        If Internal.IsModifier(HookEventArgs.KeyCode, ModifierKeys.Control) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Control
                        If Internal.IsModifier(HookEventArgs.KeyCode, ModifierKeys.Shift) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Shift
                        If Internal.IsModifier(HookEventArgs.KeyCode, ModifierKeys.Alt) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Alt
                        If Internal.IsModifier(HookEventArgs.KeyCode, ModifierKeys.Windows) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Windows

                        RaiseEvent KeyDown(Me, HookEventArgs)
                        Block = HookEventArgs.Block

                    Case NativeMethods.KeyMessage.WM_KEYUP, NativeMethods.KeyMessage.WM_SYSKEYUP
                        Dim KeyCode As Keys = CType(KeystrokeInfo.vkCode And Integer.MaxValue, Keys)

                        'Must be done before creating the HookEventArgs during KeyUp.
                        If Internal.IsModifier(KeyCode, ModifierKeys.Control) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Control
                        If Internal.IsModifier(KeyCode, ModifierKeys.Shift) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Shift
                        If Internal.IsModifier(KeyCode, ModifierKeys.Alt) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Alt
                        If Internal.IsModifier(KeyCode, ModifierKeys.Windows) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Windows

                        Dim HookEventArgs As InputHelper.EventArgs.KeyboardHookEventArgs = Me.CreateEventArgs(False, KeystrokeInfo)

                        RaiseEvent KeyUp(Me, HookEventArgs)
                        Block = HookEventArgs.Block

                End Select
            End If

            Return If(Block, New IntPtr(1), NativeMethods.CallNextHookEx(hHook, nCode, wParam, lParam))
        End Function

        Private Function CreateEventArgs(ByVal KeyDown As Boolean, ByVal KeystrokeInfo As NativeMethods.KBDLLHOOKSTRUCT) As InputHelper.EventArgs.KeyboardHookEventArgs
            Dim KeyCode As Keys = CType(KeystrokeInfo.vkCode And Integer.MaxValue, Keys)
            Dim ScanCode As UInteger = KeystrokeInfo.scanCode
            Dim Extended As Boolean = (KeystrokeInfo.flags And NativeMethods.LowLevelKeyboardHookFlags.LLKHF_EXTENDED) = NativeMethods.LowLevelKeyboardHookFlags.LLKHF_EXTENDED
            Dim AltDown As Boolean = (KeystrokeInfo.flags And NativeMethods.LowLevelKeyboardHookFlags.LLKHF_ALTDOWN) = NativeMethods.LowLevelKeyboardHookFlags.LLKHF_ALTDOWN
            Dim Injected As Boolean = (KeystrokeInfo.flags And NativeMethods.LowLevelKeyboardHookFlags.LLKHF_INJECTED) = NativeMethods.LowLevelKeyboardHookFlags.LLKHF_INJECTED
            Dim InjectedAtLowerIL As Boolean = (KeystrokeInfo.flags And NativeMethods.LowLevelKeyboardHookFlags.LLKHF_LOWER_IL_INJECTED) = NativeMethods.LowLevelKeyboardHookFlags.LLKHF_LOWER_IL_INJECTED

            If AltDown = True _
                AndAlso Internal.IsModifier(KeyCode, ModifierKeys.Alt) = False _
                 AndAlso (Me.Modifiers And ModifierKeys.Alt) <> ModifierKeys.Alt Then

                Me.Modifiers = Me.Modifiers Or ModifierKeys.Alt
            End If

            Return New InputHelper.EventArgs.KeyboardHookEventArgs(KeyCode, ScanCode, Extended, If(KeyDown, InputHelper.Hooks.KeyState.Down, InputHelper.Hooks.KeyState.Up), Me.Modifiers, Injected, InjectedAtLowerIL)
        End Function
        Public Sub New()

        End Sub

        Public Sub StartHookking()
            If hHook = IntPtr.Zero Then
                hHook = NativeMethods.SetWindowsHookEx(NativeMethods.HookType.WH_KEYBOARD_LL, HookProcedureDelegate, NativeMethods.GetModuleHandle(Nothing), 0)
                If hHook = IntPtr.Zero Then
                    Dim Win32Error As Integer = Marshal.GetLastWin32Error()
                    Throw New Win32Exception(Win32Error, "Failed to create keyboard hook! (" & Win32Error & ")")
                End If
            End If
        End Sub

        Public Sub StopHookking()
            If hHook <> IntPtr.Zero Then
                NativeMethods.UnhookWindowsHookEx(hHook)
                hHook = IntPtr.Zero
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