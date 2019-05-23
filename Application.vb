Public Class Application
    Private ReadOnly _app As Object

    Public Sub New()
        Try
            Dim typePowerPoint As Type = Type.GetTypeFromProgID("PowerPoint.Application")
            Me._app = Activator.CreateInstance(typePowerPoint)
        Catch ex As Exception
            Throw New ApplicationException("It was not possible to start PowerPoint - " & ex.Message, ex)
        End Try

        If CDec(Me.Version) < 10 Then
            Me._app = Nothing
            Throw New ApplicationException("The version of PowerPoint installed is too early to be supported.")
        End If
    End Sub

    Friend Sub New(ByVal application As Object)
        Me._app = application
    End Sub

    Public Property Visible As Boolean
        Get
            Return Me._app.Visible
        End Get
        Set(ByVal value As Boolean)
            Me._app.visible = value
        End Set
    End Property

    Public ReadOnly Property Presentations As Presentations
        Get
            Return New Presentations(Me._app.presentations)
        End Get
    End Property

    Public ReadOnly Property Version As String
        Get
            Return Me._app.Version
        End Get
    End Property

    Public Sub Quit()
        Call Me._app.Quit()
    End Sub

    Public Sub Run(ByVal macroName As String)
        Call Me._app.Run(macroName)
    End Sub

    Public Sub Run(ByVal macroName As String, ByVal ParamArray arrayOfParams() As Object)
        Call Me._app.Run(macroName, arrayOfParams)
    End Sub

    Public Property WindowState As PpWindowState
        Get
            Return Me._app.WindowState
        End Get
        Set(ByVal value As PpWindowState)
            Me._app.WindowState = value
        End Set
    End Property
End Class
