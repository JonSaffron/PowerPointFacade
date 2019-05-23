Public Class ParagraphFormat
    Private ReadOnly _paragraphFormat As Object

    Friend Sub New(ByVal paragraphFormat As Object)
        Me._paragraphFormat = paragraphformat
    End Sub

    Public ReadOnly Property Bullet As BulletFormat
        Get
            Return New BulletFormat(Me._paragraphFormat.Bullet)
        End Get
    End Property
End Class
