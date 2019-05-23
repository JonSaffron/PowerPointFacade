' Tags is a 1 based collection of the type string
Public Class Tags
    Implements IEnumerable(Of String)
    Implements IEnumerator(Of String)
    Implements IDisposable

    Private ReadOnly _tags As Object

    Friend Sub New(ByVal tags As Object)
        Me._tags = tags
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._tags.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal tagName As String) As String
        Get
            Return Me._tags.Item(tagName)
        End Get
    End Property

    Public ReadOnly Property Name(ByVal index As Integer) As String
        Get
            Return Me._tags.Name(index)
        End Get
    End Property

    Public ReadOnly Property Value(ByVal index As Integer) As String
        Get
            Return Me._tags.Value(index)
        End Get
    End Property

    Public Sub Delete(ByVal tagName As String)
        Call Me._tags.Delete(tagName)
    End Sub

    Public Sub Add(ByVal tagName As String, ByVal tagValue As String)
        Call Me._tags.add(tagName, tagValue)
    End Sub

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfString() As IEnumerator(Of String) Implements IEnumerable(Of String).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfString As String Implements IEnumerator(Of String).Current
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

    ' IDisposable
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
        Call Dispose(True)
        Call GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
