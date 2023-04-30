Namespace InputHelper.EventArgs
    Public Class MouseHookEventArgs
        Inherits System.EventArgs

        Private _button As MouseButtons
        Private _buttonState As Hooks.KeyState
        Private _delta As Integer
        Private _doubleClick As Boolean
        Private _injected As Boolean
        Private _injectedAtLowerIL As Boolean
        Private _location As Point
        Private _scrollDirection As Hooks.ScrollDirection
        Public Property Block As Boolean = False

        Public ReadOnly Property Button As MouseButtons
            Get
                Return _button
            End Get
        End Property

        Public ReadOnly Property ButtonState As Hooks.KeyState
            Get
                Return _buttonState
            End Get
        End Property

        Public ReadOnly Property Delta As Integer
            Get
                Return _delta
            End Get
        End Property

        Public ReadOnly Property DoubleClick As Boolean
            Get
                Return _doubleClick
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
        Public ReadOnly Property Location As Point
            Get
                Return _location
            End Get
        End Property
        Public ReadOnly Property ScrollDirection As Hooks.ScrollDirection
            Get
                Return _scrollDirection
            End Get
        End Property
        Public Overrides Function ToString() As String
            Return String.Format("{{Button: {0}, State: {1}, DoubleClick: {2}, Location: {3}, Scroll: {4}, Delta: {5}, Injected: {6}, InjectedAtLowerIL: {7}}}",
                                 Me.Button, Me.ButtonState, Me.DoubleClick, Me.Location, Me.ScrollDirection, Me.Delta, Me.Injected, Me.InjectedAtLowerIL)
        End Function

        Public Sub New(ByVal Button As MouseButtons,
                       ByVal ButtonState As Hooks.KeyState,
                       ByVal DoubleClick As Boolean,
                       ByVal Location As Point,
                       ByVal ScrollDirection As Hooks.ScrollDirection,
                       ByVal Delta As Integer)
            Me._button = Button
            Me._buttonState = ButtonState
            Me._doubleClick = DoubleClick
            Me._location = Location
            Me._scrollDirection = ScrollDirection
            Me._delta = Delta
        End Sub

        Public Sub New(ByVal Button As MouseButtons,
                       ByVal ButtonState As Hooks.KeyState,
                       ByVal DoubleClick As Boolean,
                       ByVal Location As Point,
                       ByVal ScrollDirection As Hooks.ScrollDirection,
                       ByVal Delta As Integer,
                       ByVal Injected As Boolean,
                       ByVal InjectedAtLowerIL As Boolean)
            Me._button = Button
            Me._buttonState = ButtonState
            Me._doubleClick = DoubleClick
            Me._location = Location
            Me._scrollDirection = ScrollDirection
            Me._delta = Delta
            Me._injected = Injected
            Me._injectedAtLowerIL = InjectedAtLowerIL
        End Sub
    End Class
End Namespace