Public Class Column
    Private ReadOnly _column As Object

    Friend Sub New(ByVal column As Object)
        Me._column = column
    End Sub

    Public Sub Delete()
        Call Me._column.Delete()
    End Sub

    Public ReadOnly Property Cells As CellRange
        Get
            Return New CellRange(Me._column.Cells)
        End Get
    End Property
End Class
