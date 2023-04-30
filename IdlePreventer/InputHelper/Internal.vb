Imports System.Windows.Forms

''' <summary>
''' A class holding internal methods and fields related to InputHelper.
''' </summary>
''' <remarks></remarks>
Friend Class Internal

#Region "ExtractModifiers()"
    ''' <summary>
    ''' Extracts any .NET modifiers from the specified key combination and returns them as native virtual key code keys.
    ''' </summary>
    ''' <param name="Key">The key combination to extract the modifiers from (if any).</param>
    ''' <remarks></remarks>
    Public Shared Function ExtractModifiers(ByRef Key As Keys) As Keys()
        Dim Modifiers As New List(Of Keys)

        If (Key And Keys.Control) = Keys.Control Then
            Key = Key And Not Keys.Control
            Modifiers.Add(Keys.ControlKey)
        End If

        If (Key And Keys.Shift) = Keys.Shift Then
            Key = Key And Not Keys.Shift
            Modifiers.Add(Keys.ShiftKey)
        End If

        If (Key And Keys.Alt) = Keys.Alt Then
            Key = Key And Not Keys.Alt
            Modifiers.Add(Keys.Menu)
        End If

        Return Modifiers.ToArray()
    End Function
#End Region

#Region "IsModifier()"
    ''' <summary>
    ''' Checks whether the specified key is any Left or Right version of the specified modifier.
    ''' </summary>
    ''' <param name="Key">The key to check.</param>
    ''' <param name="Modifier">The modifier to check for.</param>
    ''' <remarks></remarks>
    Public Shared Function IsModifier(ByVal Key As Keys, ByVal Modifier As ModifierKeys) As Boolean
        Select Case Modifier
            Case ModifierKeys.Control
                Return _
                    Key = Keys.Control OrElse _
                    Key = Keys.ControlKey OrElse _
                    Key = Keys.LControlKey OrElse _
                    Key = Keys.RControlKey
            Case ModifierKeys.Shift
                Return _
                    Key = Keys.Shift OrElse _
                    Key = Keys.ShiftKey OrElse _
                    Key = Keys.LShiftKey OrElse _
                    Key = Keys.RShiftKey
            Case ModifierKeys.Alt
                Return _
                    Key = Keys.Alt OrElse _
                    Key = Keys.Menu OrElse _
                    Key = Keys.LMenu OrElse _
                    Key = Keys.RMenu
            Case ModifierKeys.Windows
                Return _
                    Key = Keys.LWin OrElse _
                    Key = Keys.RWin
        End Select
        Throw New ArgumentOutOfRangeException("Modifier", CType(Modifier, Integer) & " is not a valid modifier key!")
    End Function
#End Region

End Class
