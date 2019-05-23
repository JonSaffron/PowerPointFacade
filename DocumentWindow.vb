Public Class DocumentWindow
    ' ReSharper disable once NotAccessedField.Local - todo...
    Private ReadOnly _documentWindow As Object

    Friend Sub New(ByVal documentWindow As Object)
        Me._documentWindow = documentWindow
    End Sub

End Class
