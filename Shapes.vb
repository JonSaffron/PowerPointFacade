' Shapes is a 1 based collection of object Shape
Public Class Shapes
    Implements IEnumerable(Of Shape)
    Implements IEnumerator(Of Shape)
    Implements IDisposable

    Private ReadOnly _shapes As Object

    Friend Sub New(ByVal shapes As Object)
        Me._shapes = shapes
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._shapes.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Shape
        Get
            Return New Shape(Me._shapes.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal shapeName As String) As Shape
        Get
            Return New Shape(Me._shapes.Item(shapeName))
        End Get
    End Property

    Public Function AddPicture(ByVal filename As String, ByVal linkToFile As Boolean, ByVal saveWithDocument As Boolean, ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single) As Shape
        Return New Shape(Me._shapes.AddPicture(filename, linkToFile, saveWithDocument, left, top, width, height))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfShape() As IEnumerator(Of Shape) Implements IEnumerable(Of Shape).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfShape As Shape Implements IEnumerator(Of Shape).Current
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
