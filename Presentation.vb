Public Class Presentation
    Private ReadOnly _presentation As Object

    Friend Sub New(ByVal presentation As Object)
        Me._presentation = presentation
    End Sub

    Public ReadOnly Property Slides As Slides
        Get
            Return New Slides(Me._presentation.Slides)
        End Get
    End Property

    Public Sub Save()
        Call Me._presentation.Save()
    End Sub

    Public Sub Close()
        Call Me._presentation.Close()
    End Sub

    Public Property Saved As Boolean
        Get
            Return Me._presentation.Saved
        End Get
        Set(ByVal value As Boolean)
            Me._presentation.Saved = value
        End Set
    End Property

    Public ReadOnly Property Name As String
        Get
            Return Me._presentation.Name
        End Get
    End Property

    Public ReadOnly Property FullName As String
        Get
            Return Me._presentation.FullName
        End Get
    End Property

    Public Function NewWindow() As DocumentWindow
        Return New DocumentWindow(Me._presentation.NewWindow())
    End Function
End Class
