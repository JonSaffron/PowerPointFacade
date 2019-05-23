Public Class BulletFormat
    Private ReadOnly _bulletformat As Object

    Friend Sub New(ByVal bulletformat As Object)
        Me._bulletformat = bulletformat
    End Sub

    Public Property Type As PpBulletType
        Get
            Return Me._bulletformat.Type
        End Get
        Set(ByVal value As PpBulletType)
            Me._bulletformat.Type = value
        End Set
    End Property
End Class
