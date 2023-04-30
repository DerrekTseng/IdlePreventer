Imports System.Runtime.InteropServices
Imports System.ComponentModel

Namespace Hooks

    Public NotInheritable Class MouseHook
        Implements IDisposable

        Public Event MouseDown As EventHandler(Of InputHelper.EventArgs.MouseHookEventArgs)

        Public Event MouseUp As EventHandler(Of InputHelper.EventArgs.MouseHookEventArgs)

        Public Event MouseMove As EventHandler(Of InputHelper.EventArgs.MouseHookEventArgs)

        Public Event MouseWheel As EventHandler(Of InputHelper.EventArgs.MouseHookEventArgs)

        Private hHook As IntPtr = IntPtr.Zero
        Private HookProcedureDelegate As New NativeMethods.LowLevelMouseProc(AddressOf HookCallback)

        Private LeftClickTimeStamp As Integer = 0
        Private MiddleClickTimeStamp As Integer = 0
        Private RightClickTimeStamp As Integer = 0
        Private X1ClickTimeStamp As Integer = 0
        Private X2ClickTimeStamp As Integer = 0

        Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
            Dim Block As Boolean = False

            If nCode >= NativeMethods.HookCode.HC_ACTION AndAlso _
                (wParam = NativeMethods.MouseMessage.WM_LBUTTONDOWN OrElse _
                 wParam = NativeMethods.MouseMessage.WM_LBUTTONUP OrElse _
                 wParam = NativeMethods.MouseMessage.WM_MBUTTONDOWN OrElse _
                 wParam = NativeMethods.MouseMessage.WM_MBUTTONUP OrElse _
                 wParam = NativeMethods.MouseMessage.WM_RBUTTONDOWN OrElse _
                 wParam = NativeMethods.MouseMessage.WM_RBUTTONUP OrElse _
                 wParam = NativeMethods.MouseMessage.WM_XBUTTONDOWN OrElse _
                 wParam = NativeMethods.MouseMessage.WM_XBUTTONUP OrElse _
                 wParam = NativeMethods.MouseMessage.WM_MOUSEWHEEL OrElse _
                 wParam = NativeMethods.MouseMessage.WM_MOUSEHWHEEL OrElse _
                 wParam = NativeMethods.MouseMessage.WM_MOUSEMOVE) Then

                Dim MouseEventInfo As NativeMethods.MSLLHOOKSTRUCT = _
                    CType(Marshal.PtrToStructure(lParam, GetType(NativeMethods.MSLLHOOKSTRUCT)), NativeMethods.MSLLHOOKSTRUCT)

                Dim Injected As Boolean = (MouseEventInfo.flags And NativeMethods.LowLevelMouseHookFlags.LLMHF_INJECTED) = NativeMethods.LowLevelMouseHookFlags.LLMHF_INJECTED
                Dim InjectedAtLowerIL As Boolean = (MouseEventInfo.flags And NativeMethods.LowLevelMouseHookFlags.LLMHF_LOWER_IL_INJECTED) = NativeMethods.LowLevelMouseHookFlags.LLMHF_LOWER_IL_INJECTED

                Select Case wParam
                    Case NativeMethods.MouseMessage.WM_LBUTTONDOWN
                        Dim DoubleClick As Boolean = (Environment.TickCount - LeftClickTimeStamp) <= NativeMethods.GetDoubleClickTime()
                        LeftClickTimeStamp = Environment.TickCount

                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(MouseButtons.Left, InputHelper.Hooks.KeyState.Down, DoubleClick,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.None, 0, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseDown(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_LBUTTONUP
                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(MouseButtons.Left, InputHelper.Hooks.KeyState.Up, False,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.None, 0, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseUp(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_MBUTTONDOWN
                        Dim DoubleClick As Boolean = (Environment.TickCount - MiddleClickTimeStamp) <= NativeMethods.GetDoubleClickTime()
                        MiddleClickTimeStamp = Environment.TickCount

                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(MouseButtons.Middle, InputHelper.Hooks.KeyState.Down, DoubleClick,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.None, 0, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseDown(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_MBUTTONUP
                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(MouseButtons.Middle, InputHelper.Hooks.KeyState.Up, False,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.None, 0, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseUp(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_RBUTTONDOWN
                        Dim DoubleClick As Boolean = (Environment.TickCount - RightClickTimeStamp) <= NativeMethods.GetDoubleClickTime()
                        RightClickTimeStamp = Environment.TickCount

                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(MouseButtons.Right, InputHelper.Hooks.KeyState.Down, DoubleClick,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.None, 0, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseDown(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_RBUTTONUP
                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(MouseButtons.Right, InputHelper.Hooks.KeyState.Up, False,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.None, 0, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseUp(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_XBUTTONDOWN
                        Dim IsXButton2 As Boolean = (New NativeMethods.DWORD(MouseEventInfo.mouseData).High = 2)
                        Dim DoubleClick As Boolean = (Environment.TickCount - If(IsXButton2, X2ClickTimeStamp, X1ClickTimeStamp)) <= NativeMethods.GetDoubleClickTime()

                        If IsXButton2 = True Then X2ClickTimeStamp = Environment.TickCount _
                                             Else X1ClickTimeStamp = Environment.TickCount

                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(If(IsXButton2, MouseButtons.XButton2, MouseButtons.XButton1), InputHelper.Hooks.KeyState.Down, DoubleClick,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.None, 0, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseDown(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_XBUTTONUP
                        Dim IsXButton2 As Boolean = (New NativeMethods.DWORD(MouseEventInfo.mouseData).High = 2)
                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(If(IsXButton2, MouseButtons.XButton2, MouseButtons.XButton1), InputHelper.Hooks.KeyState.Up, False,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.None, 0, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseUp(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_MOUSEWHEEL
                        Dim Delta As Integer = New NativeMethods.DWORD(MouseEventInfo.mouseData).SignedHigh
                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(MouseButtons.None, InputHelper.Hooks.KeyState.Up, False,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.Vertical, Delta, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseWheel(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_MOUSEHWHEEL
                        Dim Delta As Integer = New NativeMethods.DWORD(MouseEventInfo.mouseData).SignedHigh
                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(MouseButtons.None, InputHelper.Hooks.KeyState.Up, False,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.Horizontal, Delta, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseWheel(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_MOUSEMOVE
                        Dim HookEventArgs As New InputHelper.EventArgs.MouseHookEventArgs(MouseButtons.None, InputHelper.Hooks.KeyState.Up, False,
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), InputHelper.Hooks.ScrollDirection.None, 0, Injected, InjectedAtLowerIL)
                        RaiseEvent MouseMove(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                End Select
            End If

            Return If(Block, New IntPtr(1), NativeMethods.CallNextHookEx(hHook, nCode, wParam, lParam))
        End Function
        Public Sub New()
            hHook = NativeMethods.SetWindowsHookEx(NativeMethods.HookType.WH_MOUSE_LL, HookProcedureDelegate, NativeMethods.GetModuleHandle(Nothing), 0)
            If hHook = IntPtr.Zero Then
                Dim Win32Error As Integer = Marshal.GetLastWin32Error()
                Throw New Win32Exception(Win32Error, "Failed to create mouse hook! (" & Win32Error & ")")
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