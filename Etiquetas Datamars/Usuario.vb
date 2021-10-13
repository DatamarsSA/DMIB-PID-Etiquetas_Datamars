Public Class Usuario

    Dim id As Integer
    Dim Nombre As String
    Dim grupoTrabajo() As String

    Sub New()

        ' Llamada necesaria para el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        Nombre = "More.."
        lblNombre.Text = Nombre
        Me.Visible = False

        grupoTrabajo = {""}
    End Sub

    Sub New(_id As Integer, _nombre As String, _grupos() As String)
        id = _id
        Nombre = _nombre
        grupoTrabajo = _grupos

        InitializeComponent()

        lblNombre.Text = Nombre

    End Sub

    Public Function estaEnGrupo(_grupo As String) As Boolean

        If _grupo Is Nothing Then Return True
        If _grupo = "" Then Return True

        For Each s As String In grupoTrabajo
            If s.ToUpper = _grupo.ToUpper Then Return True
        Next
        Return False
    End Function

    Private Sub lblNombre_Click(sender As Object, e As EventArgs) Handles lblNombre.Click
        'Al pulsar sobre el nombre del usuario lanzamos un evento de formulario, sino se haria un evento de usuario.
        RaiseEvent miClick(Me)
    End Sub

    Public Function getID() As Integer
        Return id
    End Function

    Public Function getNombre() As String
        Return Nombre
    End Function
    Public Sub cambiarLabel()
        If Nombre = "More.." Then
            Nombre = "Menos.."
        Else
            Nombre = "More.."
        End If

        lblNombre.Text = Nombre
    End Sub
    'Declaración del evento a nivel de clase.
    Event miClick(sender As Usuario)

End Class
