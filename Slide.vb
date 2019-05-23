Public Class Slide
    Private ReadOnly _slide As Object

    Friend Sub New(ByVal slide As Object)
        Me._slide = slide
    End Sub

    Public ReadOnly Property Application As Application
        Get
            Return New Application(Me._slide.Application)
        End Get
    End Property

    Public ReadOnly Property Shapes As Shapes
        Get
            Return New Shapes(Me._slide.Shapes)
        End Get
    End Property

    Public Function Duplicate() As SlideRange
        Return New SlideRange(Me._slide.Duplicate())
    End Function

    Public ReadOnly Property SlideNumber As Integer
        Get
            Return Me._slide.SlideNumber
        End Get
    End Property

    Public Sub [Select]()
        Call Me._slide.Select()
    End Sub

    Public Property Name As String
        Get
            Return Me._slide.Name
        End Get
        Set(ByVal value As String)
            Me._slide.Name = value
        End Set
    End Property
End Class
