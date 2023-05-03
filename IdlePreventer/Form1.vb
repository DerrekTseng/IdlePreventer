Imports System.ComponentModel
Imports System.Threading

Public Class Form1
    Dim Generator As System.Random = New System.Random()

    Dim MyMouseHook As New Hooks.MouseHook
    Dim MyKeyboardHook As New Hooks.KeyboardHook
    Dim IdelSecond As Integer
    ReadOnly AFK As Integer = 60

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NotifyIcon1.Visible = False
        AddHandler MyMouseHook.MouseDown, AddressOf MouseHook_MouseEvent
        AddHandler MyMouseHook.MouseUp, AddressOf MouseHook_MouseEvent
        AddHandler MyMouseHook.MouseMove, AddressOf MouseHook_MouseEvent
        AddHandler MyMouseHook.MouseWheel, AddressOf MouseHook_MouseEvent
        AddHandler MyKeyboardHook.KeyDown, AddressOf KeyboardHook_KeyEvent
        AddHandler MyKeyboardHook.KeyUp, AddressOf KeyboardHook_KeyEvent
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Interlocked.Exchange(IdelSecond, 0)

        Timer1.Enabled = Not Timer1.Enabled

        If Timer1.Enabled Then
            MyMouseHook.StartHookking()
            MyKeyboardHook.StartHookking()
            Button1.Text = "Stop"
        Else
            MyMouseHook.StopHookking()
            MyKeyboardHook.StopHookking()
            Button1.Text = "Start"
        End If

    End Sub

    Private Sub MouseHook_MouseEvent(sender As Object, e As InputHelper.EventArgs.MouseHookEventArgs)
        Interlocked.Exchange(IdelSecond, 0)
    End Sub

    Private Sub KeyboardHook_KeyEvent(sender As Object, e As InputHelper.EventArgs.KeyboardHookEventArgs)
        Interlocked.Exchange(IdelSecond, 0)
    End Sub

    Private Function RandomPoint() As Point
        Return New Point(Generator.Next(0, Screen.PrimaryScreen.Bounds.Width), Generator.Next(0, Screen.PrimaryScreen.Bounds.Height))
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Interlocked.Increment(IdelSecond)
        Dim IdelSecondValue As Integer = IdelSecond
        If (IdelSecondValue >= AFK) Then
            Dim SavedLocation = MouseControl.GetMouseLocation()
            MouseControl.SetMouseLocation(RandomPoint())
            System.Threading.Thread.Sleep(10)
            MouseControl.SetMouseLocation(SavedLocation)
            Interlocked.Exchange(IdelSecond, 0)
        End If
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If sender.WindowState = FormWindowState.Minimized Then
            Me.Hide()
            NotifyIcon1.Visible = True
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
        NotifyIcon1.Visible = False
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Me.Show()
        NotifyIcon1.Visible = False
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Timer1.Enabled = False
        MyMouseHook.StopHookking()
        MyKeyboardHook.StopHookking()
    End Sub
End Class
