Imports GetDomain

Public Class frmUsuarios

    Sub New()

        ' Llamada necesaria para el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub
    ''Private Sub frmUsuarios_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
    ''    For i As Integer = panel.Controls.Count - 1 To 0 Step -1
    ''        Dim c As Usuario = panel.Controls(i)
    ''        RemoveHandler c.miClick, AddressOf Click_Usuario
    ''        panel.Controls.Remove(c)
    ''    Next
    ''End Sub

    Private Sub frmUsuarios_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CargarUsuarios()
        MostrarUsuarios(True)

    End Sub

    Sub MostrarUsuarios(Optional _todos As Boolean = False)

        For i As Integer = panel.Controls.Count - 1 To 0 Step -1
            Dim c As Usuario = panel.Controls(i)
            RemoveHandler c.miClick, AddressOf Click_Usuario
            panel.Controls.Remove(c)
        Next

        If _todos Then
            For Each u As Usuario In listausuarios
                AddHandler u.miClick, AddressOf Click_Usuario 'si se produce el evento mi click ejecuta el procedimiento click_usuario.
                panel.Controls.Add(u)
            Next

        Else
            Dim grupo As String = New ClasePrincipal().get
            For Each u As Usuario In listaUsuarios
                AddHandler u.miClick, AddressOf Click_Usuario
                If u.estaEnGrupo(grupo) Then panel.Controls.Add(u)
            Next

        End If
        panel.Controls.Add(listausuarios(listausuarios.Count - 1))

    End Sub


    Sub Click_Usuario(sender As Usuario)
        If sender.getID = 0 Then
            If sender.getNombre = "Más..." Then
                MostrarUsuarios(True)
            Else
                MostrarUsuarios(False)
            End If
            sender.cambiarLabel()
        Else
            usuarioActual = sender
            Me.DialogResult = DialogResult.OK
            Form1.Show()
            Me.Close()
        End If
    End Sub
End Class