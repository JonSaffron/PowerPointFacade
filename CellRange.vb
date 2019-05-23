' CellRange is a 1 based collection of object Cell
Public Class CellRange
    Implements IEnumerable(Of Cell)
    Implements IEnumerator(Of Cell)
    Implements IDisposable

    Private ReadOnly _cellrange As Object

    Friend Sub New(ByVal cellrange As Object)
        Me._cellrange = cellrange
        Call Me.Reset()
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Cell
        Get
            Return New Cell(Me._cellrange.Item(index))
        End Get
    End Property

    Public ReadOnly Property Count As Integer
        Get
            Return Me._cellrange.Count
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfCell() As IEnumerator(Of Cell) Implements IEnumerable(Of Cell).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfCell As Cell Implements IEnumerator(Of Cell).Current
        Get
            Return Me.Item(Me._enumeratorPosition)
        End Get
    End Property

    Public ReadOnly Property Current As Object Implements IEnumerator.Current
        Get
            Return Me.Item(Me._enumeratorPosition)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        Me._enumeratorPosition += 1
        Return (Me._enumeratorPosition <= Me.Count)
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        Me._enumeratorPosition = 0
    End Sub
#End Region

#Region " IDisposable Support "
    Private _isDisposed As Boolean = False

    Protected Overridable Sub Dispose(ByVal isSafeToDisposeManagedResources As Boolean)
        If Not Me._isDisposed Then
            If isSafeToDisposeManagedResources Then
                ' no managed resources to free
            End If

            ' no shared unmanaged resources to free
        End If
        Me._isDisposed = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Call Me.Dispose(True)
        Call GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
