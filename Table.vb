Public Class Table
    ' ReSharper disable once InconsistentNaming - consistent within this library
    Protected ReadOnly _table As Object

    Friend Sub New(ByVal table As Object)
        Me._table = table
    End Sub

    Public ReadOnly Property Application As Application
        Get
            Return New Application(Me._table.Application)
        End Get
    End Property

    Public ReadOnly Property Columns As Columns
        Get
            Return New Columns(Me._table.Columns)
        End Get
    End Property

    Public ReadOnly Property Rows As Rows
        Get
            Return New Rows(Me._table.Rows)
        End Get
    End Property

    Public Function Cell(ByVal row As Integer, ByVal column As Integer) As Cell
        Return New Cell(Me._table.Cell(row, column))
    End Function

    Public ReadOnly Property Parent As Shape
        Get
            Return New Shape(Me._table.Parent)
        End Get
    End Property
End Class
