Public Class Font
    Private ReadOnly _font As Object

    Friend Sub New(ByVal font As Object)
        Me._font = font
    End Sub

    Public Property Bold As MsoTriState
        Get
            Return Me._font.Bold
        End Get
        Set(ByVal value As MsoTriState)
            Me._font.Bold = value
        End Set
    End Property

    Public Property Italic As MsoTriState
        Get
            Return Me._font.Italic
        End Get
        Set(ByVal value As MsoTriState)
            Me._font.Italic = value
        End Set
    End Property

    Public Property Underline As MsoTriState
        Get
            Return Me._font.Underline
        End Get
        Set(ByVal value As MsoTriState)
            Me._font.Underline = value
        End Set
    End Property

    Public Property Size As Single
        Get
            Return Me._font.Size
        End Get
        Set(ByVal value As Single)
            Me._font.Size = value
        End Set
    End Property
End Class
