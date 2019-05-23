Public Class Table2
    Inherits Table

    Friend Sub New(ByVal table As Object)
        MyBase.New(table)

        Dim app As Application = New Application(Me._table.Application)
        Dim applicationVersion As Decimal = CDec(app.Version)
        If applicationVersion < 12 Then Throw New ApplicationException("Cannot instantiate the Table2 class against version " & applicationVersion & " of PowerPoint.")
    End Sub

    Public Property FirstCol As Boolean
        Get
            Return Me._table.FirstCol
        End Get
        Set(ByVal value As Boolean)
            Me._table.FirstCol = value
        End Set
    End Property

    Public Property FirstRow As Boolean
        Get
            Return Me._table.FirstRow
        End Get
        Set(ByVal value As Boolean)
            Me._table.FirstRow = value
        End Set
    End Property

    Public Property LastCol As Boolean
        Get
            Return Me._table.LastCol
        End Get
        Set(ByVal value As Boolean)
            Me._table.LastCol = value
        End Set
    End Property

    Public Property LastRow As Boolean
        Get
            Return Me._table.LastRow
        End Get
        Set(ByVal value As Boolean)
            Me._table.LastRow = value
        End Set
    End Property
End Class
