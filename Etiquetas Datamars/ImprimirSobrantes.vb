Public Class ImprimirSobrantes
    Private Sub botonNo_Click(sender As Object, e As EventArgs) Handles botonNo.Click
        Form1.añadirDosPorciento = False
        Me.Close()
    End Sub

    Private Sub botonSi_Click(sender As Object, e As EventArgs) Handles botonSi.Click
        Form1.añadirDosPorciento = True
        Me.Close()
    End Sub
End Class