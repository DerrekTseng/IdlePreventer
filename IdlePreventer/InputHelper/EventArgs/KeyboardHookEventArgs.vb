Namespace InputHelper.EventArgs
    Public Class KeyboardHookEventArgs
        Inherits System.EventArgs

        Private _extended As Boolean
        Private _injected As Boolean
        Private _injectedAtLowerIL As Boolean
        Private _keyCode As Keys
        Private _keyState As Hooks.KeyState
        Private _modifiers As ModifierKeys
        Private _scanCode As UInteger

        Public Property Block As Boolean = False

        Public ReadOnly Property Extended As Boolean
            Get
                Return _extended
            End Get
        End Property


        Public ReadOnly Property Injected As Boolean
            Get
                Return _injected
            End Get
        End Property

        Public ReadOnly Property InjectedAtLowerIL As Boolean
            Get
                Return _injectedAtLowerIL
            End Get
        End Property

        Public ReadOnly Property KeyCode As Keys
            Get
                Return _keyCode
            End Get
        End Property

        Public ReadOnly Property KeyState As Hooks.KeyState
            Get
                Return _keyState
            End Get
        End Property


        Public ReadOnly Property Modifiers As ModifierKeys
            Get
                Return _modifiers
            End Get
        End Property

        Public ReadOnly Property ScanCode As UInteger
            Get
                Return _scanCode
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return String.Format("{{KeyCode: {0}, ScanCode: {1}, Extended: {2}, KeyState: {3}, Modifiers: {4}, Injected: {5}, InjectedAtLowerIL: {6}}}",
                                 Me.KeyCode, Me.ScanCode, Me.Extended, Me.KeyState, Me.Modifiers, Me.Injected, Me.InjectedAtLowerIL)
        End Function

        Public Sub New(ByVal KeyCode As Keys,
                       ByVal ScanCode As UInteger,
                       ByVal Extended As Boolean,
                       ByVal KeyState As Hooks.KeyState,
                       ByVal Modifiers As ModifierKeys)
            Me._keyCode = KeyCode
            Me._scanCode = ScanCode
            Me._extended = Extended
            Me._keyState = KeyState
            Me._modifiers = Modifiers
        End Sub

        Public Sub New(ByVal KeyCode As Keys,
                       ByVal ScanCode As UInteger,
                       ByVal Extended As Boolean,
                       ByVal KeyState As Hooks.KeyState,
                       ByVal Modifiers As ModifierKeys,
                       ByVal Injected As Boolean,
                       ByVal InjectedAtLowerIL As Boolean)
            Me._keyCode = KeyCode
            Me._scanCode = ScanCode
            Me._extended = Extended
            Me._keyState = KeyState
            Me._modifiers = Modifiers
            Me._injected = Injected
            Me._injectedAtLowerIL = InjectedAtLowerIL
        End Sub
    End Class
End Namespace