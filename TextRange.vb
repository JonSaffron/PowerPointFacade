Public Class TextRange
    Private ReadOnly _textRange As Object

    Friend Sub New(ByVal textRange As Object)
        Me._textRange = textRange
    End Sub

    Public Property Text As String
        Get
            Return Me._textRange.Text
        End Get
        Set(ByVal value As String)
            Me._textRange.Text = value
        End Set
    End Property

    Public ReadOnly Property ParagraphFormat As ParagraphFormat
        Get
            Return New ParagraphFormat(Me._textRange.ParagraphFormat)
        End Get
    End Property

    Public Function InsertAfter(ByVal newText As String) As TextRange
        Return New TextRange(Me._textRange.InsertAfter(newText))
    End Function

    Public ReadOnly Property Count As Integer
        Get
            Return Me._textRange.Count
        End Get
    End Property

    Public ReadOnly Property Start As Integer
        Get
            Return Me._textRange.Start
        End Get
    End Property

    Public ReadOnly Property Length As Integer
        Get
            Return Me._textRange.Length
        End Get
    End Property

    Public Function Paragraphs(ByVal startParagraph As Integer) As TextRange
        Return New TextRange(Me._textRange.Paragraphs(startParagraph))
    End Function

    Public Function Characters(ByVal startCharacter As Integer, ByVal countOfCharacters As Integer) As TextRange
        Return New TextRange(Me._textRange.Characters(startCharacter, countOfCharacters))
    End Function

    Public Sub Delete()
        Call Me._textRange.Delete()
    End Sub

    Public Sub [Select]()
        Call Me._textRange.Select()
    End Sub

    Public ReadOnly Property Parent As TextFrame
        Get
            Return New TextFrame(Me._textRange.Parent)
        End Get
    End Property

    Public ReadOnly Property Font As Font
        Get
            Return New Font(Me._textRange.Font)
        End Get
    End Property
End Class
