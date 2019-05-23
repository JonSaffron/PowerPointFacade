' SlideRange is a 1 based collection of the object Slide
Public Class SlideRange
    Implements IEnumerable(Of Slide)
    Implements IEnumerator(Of Slide)
    Implements IDisposable

    Private ReadOnly _slideRange As Object

    Friend Sub New(ByVal slideRange As Object)
        Me._slideRange = slideRange
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._slideRange.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Slide
        Get
            Return New Slide(Me._sliderange.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal slideName As String) As Slide
        Get
            Return New Slide(Me._sliderange.Item(slideName))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfSlide() As IEnumerator(Of Slide) Implements IEnumerable(Of Slide).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfSlide As Slide Implements IEnumerator(Of Slide).Current
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
