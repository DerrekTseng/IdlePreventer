Imports System.Text

''' <summary>
''' A class holding various InputHelper constants.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Constants
    Private Sub New()
    End Sub

    ''' <summary>
    ''' The Alt Gr key.
    ''' </summary>
    ''' <remarks></remarks>
    Public Const AltGr As Keys = Keys.Control Or Keys.Alt

    ''' <summary>
    ''' The bit value used to check if a key is held down (masked with the return value from GetAsyncKeyState()).
    ''' </summary>
    ''' <remarks></remarks>
    Public Const KeyDownBit As Integer = &H8000

    ''' <summary>
    ''' Windows-1252 encoding.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared ReadOnly Windows1252 As Encoding = Encoding.GetEncoding("windows-1252")

End Class
