' Presentations is a 1 based collection of object Presentation
Public Class Presentations
    Implements IEnumerable(Of Presentation)
    Implements IEnumerator(Of Presentation)
    Implements IDisposable

    Private ReadOnly _presentations As Object

    Friend Sub New(ByVal presentations As Object)
        Me._presentations = presentations
        Call Me.Reset()
    End Sub

    Public Function Open(ByVal filename As String, ByVal openReadOnly As Boolean, ByVal untitled As Boolean, ByVal withWindow As Boolean) As Presentation
        Return New Presentation(Me._presentations.open(filename, openReadOnly, untitled, withWindow))
    End Function

    Public ReadOnly Property Count As Integer
        Get
            Return Me._presentations.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Presentation
        Get
            Return New Presentation(Me._presentations.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal presentationName As String) As Presentation
        Get
            Return New Presentation(Me._presentations.item(presentationName))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfPresentation() As IEnumerator(Of Presentation) Implements IEnumerable(Of Presentation).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfPresentation As Presentation Implements IEnumerator(Of Presentation).Current
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
