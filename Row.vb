Public Class Row
    Private ReadOnly _row As Object

    Friend Sub New(ByVal row As Object)
        Me._row = row
    End Sub

    Public Sub Delete()
        Call Me._row.Delete()
    End Sub

    Public ReadOnly Property Cells As CellRange
        Get
            Return New CellRange(Me._row.Cells)
        End Get
    End Property

    Public Property Height As Single
        Get
            Return Me._row.Height
        End Get
        Set(ByVal value As Single)
            Me._row.Height = value
        End Set
    End Property
End Class
