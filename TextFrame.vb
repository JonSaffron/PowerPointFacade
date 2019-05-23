Public Class TextFrame
    Private ReadOnly _textFrame As Object

    Friend Sub New(ByVal textFrame As Object)
        Me._textFrame = textFrame
    End Sub

    Public ReadOnly Property TextRange As TextRange
        Get
            Return New TextRange(Me._textFrame.TextRange)
        End Get
    End Property

    Public Sub DeleteText()
        Call Me._textframe.DeleteText()
    End Sub

    Public ReadOnly Property HasText As Boolean
        Get
            Return Me._textframe.HasText
        End Get
    End Property
End Class
