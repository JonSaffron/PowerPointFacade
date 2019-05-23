Public Class Shape
    Private ReadOnly _shape As Object

    Friend Sub New(ByVal shape As Object)
        Me._shape = shape
    End Sub

    Public ReadOnly Property Application As Application
        Get
            Return New Application(Me._shape.Application)
        End Get
    End Property

    Public ReadOnly Property Type As MsoShapeType
        Get
            Return Me._shape.Type
        End Get
    End Property

    Public ReadOnly Property HasTextFrame As Boolean
        Get
            Return Me._shape.HasTextFrame
        End Get
    End Property

    Public ReadOnly Property HasTable As Boolean
        Get
            Return Me._shape.HasTable
        End Get
    End Property

    Public ReadOnly Property TextFrame As TextFrame
        Get
            Return New TextFrame(Me._shape.TextFrame)
        End Get
    End Property

    Public ReadOnly Property Table As Table
        Get
            Dim result As Table
            If CDec(Application.Version) < 12 Then
                result = New Table(Me._shape.Table)
            Else
                result = New Table2(Me._shape.Table)
            End If
            Return result
        End Get
    End Property

    Public Property Top As Single
        Get
            Return Me._shape.Top
        End Get
        Set(ByVal value As Single)
            Me._shape.Top = value
        End Set
    End Property

    Public Property Left As Single
        Get
            Return Me._shape.Left
        End Get
        Set(ByVal value As Single)
            Me._shape.Left = value
        End Set
    End Property

    Public Property Width As Single
        Get
            Return Me._shape.Width
        End Get
        Set(ByVal value As Single)
            Me._shape.Width = value
        End Set
    End Property

    Public Property Height As Single
        Get
            Return Me._shape.Height
        End Get
        Set(ByVal value As Single)
            Me._shape.Height = value
        End Set
    End Property

    Public Sub Delete()
        Call Me._shape.Delete()
    End Sub

    Public ReadOnly Property Parent As Slide
        Get
            Return New Slide(Me._shape.Parent)
        End Get
    End Property

    Public Sub ScaleHeight(ByVal factor As Single, ByVal relativeToOriginalSize As Boolean)
        Call Me._shape.ScaleHeight(factor, relativeToOriginalSize)
    End Sub

    Public Sub ScaleHeight(ByVal factor As Single, ByVal relativeToOriginalSize As Boolean, ByVal scaleFrom As msoScaleFrom)
        Call Me._shape.ScaleHeight(factor, relativeToOriginalSize, scaleFrom)
    End Sub

    Public Sub ScaleWidth(ByVal factor As Single, ByVal relativeToOriginalSize As Boolean)
        Call Me._shape.ScaleWidth(factor, relativeToOriginalSize)
    End Sub

    Public Sub ScaleWidth(ByVal factor As Single, ByVal relativeToOriginalSize As Boolean, ByVal scaleFrom As msoScaleFrom)
        Call Me._shape.ScaleWidth(factor, relativeToOriginalSize, scaleFrom)
    End Sub

    Public Property LockAspectRatio As Boolean
        Get
            Return Me._shape.LockAspectRatio
        End Get
        Set(ByVal value As Boolean)
            Me._shape.LockAspectRatio = value
        End Set
    End Property

    Public Function Duplicate() As ShapeRange
        Return New ShapeRange(Me._shape.Duplicate())
    End Function

    Public ReadOnly Property Tags As Tags
        Get
            Return New Tags(Me._shape.Tags)
        End Get
    End Property

    Public Property Name As String
        Get
            Return Me._shape.Name
        End Get
        Set(ByVal value As String)
            Me._shape.Name = value
        End Set
    End Property
End Class
