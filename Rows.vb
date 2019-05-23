' Rows is a 1 based collection of object Row
Public Class Rows
    Implements IEnumerable(Of Row)
    Implements IEnumerator(Of Row)
    Implements IDisposable

    Private ReadOnly _rows As Object

    Friend Sub New(ByVal rows As Object)
        Me._rows = rows
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._rows.Count
        End Get
    End Property

    Public Function Add() As Row
        Return Me.Add(-1)
    End Function

    Public Function Add(ByVal beforeRow As Integer) As Row
        Return New Row(Me._rows.Add(beforeRow))
    End Function

    Default Public ReadOnly Property Item(ByVal index As Integer) As Row
        Get
            Return New Row(Me._rows.Item(index))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfRow() As IEnumerator(Of Row) Implements IEnumerable(Of Row).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfRow As Row Implements IEnumerator(Of Row).Current
        Get
            Return Me.Item(Me._enumeratorPostion)
        End Get
    End Property

    Public ReadOnly Property Current As Object Implements IEnumerator.Current
        Get
            Return Me.Item(Me._enumeratorPostion)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        Me._enumeratorPostion += 1
        Return (Me._enumeratorPostion <= Me.Count)
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        Me._enumeratorPostion = 0
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
