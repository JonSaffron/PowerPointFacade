Public Class Cell
    Private ReadOnly _cell As Object

    Friend Sub New(ByVal cell As Object)
        Me._cell = cell
    End Sub

    Public ReadOnly Property Shape As Shape
        Get
            Return New Shape(Me._cell.Shape)
        End Get
    End Property

    Public ReadOnly Property Parent As Table
        Get
            Return New Table(Me._cell.Parent)
        End Get
    End Property
End Class
