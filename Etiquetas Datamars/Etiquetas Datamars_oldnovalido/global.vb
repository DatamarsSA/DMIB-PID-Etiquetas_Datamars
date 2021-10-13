Module Globales
    Public usuarioactual As Usuario
    Public listausuarios As New List(Of Usuario)
    Public uspDatosPedidos As New DataSet1TableAdapters.conPedidosRefSapDesartTableAdapter
    Public uspRangoChip As New DataSet1TableAdapters.OrdenesGrabacionTableAdapter
    Public dt, dt1 As DataTable
    Public etiManual, histGrabado, reImprimir As Boolean
    Public IDhistorico, CajaInicial, TotalCajas, totalLotes, maxcajas, palets As Integer

    Sub CargarUsuarios()

        Dim aux As New DataSet1TableAdapters.uspBuscarUsuariosEtiquetasTableAdapter

        If Not listausuarios Is Nothing Then
            listausuarios.Clear()
            listausuarios = Nothing
        End If

        listausuarios = New List(Of Usuario)

        For Each r As DataRow In aux.GetData
            listausuarios.Add(New Usuario(r("id"), r("nombre"), {"Workgroup"}))
        Next

        listausuarios.Add(New Usuario)
    End Sub


End Module
