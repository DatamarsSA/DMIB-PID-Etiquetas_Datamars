
Public Class Form1


    Public uspDatosPedidos As New DataSet1TableAdapters.conPedidosRefSapDesartTableAdapter
    Public uspRangoChip As New DataSet1TableAdapters.OrdenesGrabacionTableAdapter
    Public dt, dt1 As DataTable
    Public etiManual, histGrabado, reImprimir As Boolean
    'Public IDhistorico, CajaInicial, TotalCajas, totalLotes As Integer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        etiManual = False
        histGrabado = False
        reImprimir = False
        txtAño.Value = Now.ToString("yy")
        txtAño.Minimum = Now.ToString("yy") - 1
        txtAño.Maximum = Now.ToString("yy") + 1

        numAñoCad.Value = Now.ToString("yy") + 5
        numAñoCad.Minimum = Now.ToString("yy") - 1 + 5
        numAñoCad.Maximum = Now.ToString("yy") + 1 + 5

        numMesCad.Value = Now.Month

        Dim cal As System.Globalization.Calendar = System.Globalization.DateTimeFormatInfo.CurrentInfo.Calendar
        txtSemana.Text = cal.GetWeekOfYear(Now, Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday).ToString
        'indicamos el numero de la caja Inicial
        CajaInicial = txtCaja.Value
        'indicamos el numero total de cajas que se van a imprimir, desde el numero inicial de caja.
        TotalCajas = (txtCaja.Value + numCajas.Value)
        cantLote = numLotes.Value
        cantCaja = numCajas.Value
        lotxCaja = numLotesPorCaja.Value
        jerxLote = numUnidadesPorlote.Value
        fecFabricacion = "20" & txtAño.Value & "-" & numMesCad.Value.ToString("00")
        ListBox1.SelectedIndex = -1
        ListBox2.SelectedIndex = -1
        'Bloqueamos los botones de imprimir hasta que no se introduzca un pedido.
        cmdImprimirCajas.Enabled = False
        cmdImprimirLotes.Enabled = False
        cmdImpManual.Enabled = False
        cmdImprimirPalets.Enabled = False
    End Sub
#Region "Funciones y Procedimientos "

    Function sacarSiglaPais(valor As Integer) As String
        'funcion que saca las siglas del pais donde va el microchip.
        Dim uspsigla As New DataSet1TableAdapters.CodigosPaisesTableAdapter
        Dim dt As DataTable
        Dim dato As String
        dt = uspsigla.GetData(valor)
        dato = dt.Rows(0)("A 3")
        Return dato
    End Function
    Private Sub activarBotones()
        cmdImprimirLotes.Enabled = True
        cmdImprimirCajas.Enabled = True
        cmdImpManual.Enabled = True
        cmdImprimirPalets.Enabled = True
    End Sub
    Private Sub cargarPedido()
        'procedimiento de busqueda de Pedido en la BBDD.
        Dim respuesta As Integer
        Dim etiquetasImpresas As String
        respuesta = 0
        If etiManual = True Then
            txtSemana.Enabled = False
            txtCaja.Enabled = False
            txtLote.Enabled = False
            numCajasPorPalet.Enabled = False
            numLotes.Enabled = False
            numCajas.Enabled = False
            numPalets.Enabled = False
            etiManual = False
        End If
        'recargamos el formulario y borramos los datos anteriores
        limpiar()
        'habilitamos los botones de impresión.
        activarBotones()
        txtCodPedido.Text = Replace(txtCodPedido.Text, "#", "")
        etiquetasImpresas = pedidoImpreso(txtCodPedido.Text)
        'si el pedido ya esta impreso preguntamos si queremos reimprimirlo
        If etiquetasImpresas <> "" Then respuesta = MsgBox("Pedido Impreso el " & etiquetasImpresas & ". ¿Quiere ReImprimilo?", MsgBoxStyle.YesNo)
        If etiquetasImpresas = "" Or etiquetasImpresas <> "" AndAlso respuesta = 6 Then
            'si el pedido es nuevo o Ya impreso y queremos modificarlo
            dt = uspDatosPedidos.GetData(txtCodPedido.Text)
            rellenarCamposPedido()
            If dt.Rows.Count > 0 Then
                If respuesta = 0 Then
                    'Para pedido nuevo
                    'calculamos la caja inicial
                    CajaInicial = ultimaCaja(txtSemana.Value, txtAño.Value)
                    txtCaja.Value = CajaInicial
                    TotalCajas = (CajaInicial + numCajas.Value)
                Else
                    'Reimprimir 
                    modificarPedidoImpreso(txtCodPedido.Text)
                End If


            Else
                MessageBox.Show("El Pedido No Existe. Compruebelo", "ATENCION", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                limpiar()
            End If
        Else
            txtCodPedido.Text = " "
            limpiar()
        End If
    End Sub
    Private Sub limpiar()
        Form1_Load(Nothing, Nothing)
    End Sub
    Sub grabarHistorico(tipoGrab As String, valor As Integer)
        'Procedimiento para grabar datos
        Dim uspInsertarHisto As New DataSet1TableAdapters.qryInsertHistorico
        Dim equipo As String
        Dim lote, caja As Boolean
        equipo = SystemInformation.ComputerName
        Dim grabCorrecta As Integer
        Select Case valor
            Case 0
                lote = True
                caja = False
            Case 1
                lote = False
                caja = True
            Case 2
                lote = True
                caja = True
        End Select

        If tipoGrab = "insertar" Then
            uspInsertarHisto.uspInsertarHistoricoEtiquetasDM(txtCodPedido.Text, txtSemana.Value, txtAño.Value, txtCaja.Value, numCajas.Value, numLotes.Value, lote, caja,
                                                         ListBox1.SelectedItem, ListBox2.SelectedItem, Now, equipo, usuarioactual.getID, reImprimir, grabCorrecta, IDhistorico)
        ElseIf tipoGrab = "actualizar" Then
            uspInsertarHisto.uspActualizarHistoricoEtiquetas(IDhistorico, lote, caja, Now, grabCorrecta)
        End If
        If grabCorrecta = 1 Then
            MsgBox("Se ha producido un error al grabar en la BB.DD.", MsgBoxStyle.Information)
        End If
    End Sub


    Sub mostrarEtiqCaja(valor As String)
        Dim imgEtiqueta As String()
        imgEtiqueta = vistaPreviaEtiqueta(1, valor)
        picEtiquetaCaja.ImageLocation = imgEtiqueta(0)
        picEtiquetaCaja.Width = Integer.Parse(imgEtiqueta(1))
        picEtiquetaCaja.Height = Integer.Parse(imgEtiqueta(2))
    End Sub
    Sub mostrarEtiqLote(valor As String)
        Dim imgEtiqueta As String()
        imgEtiqueta = vistaPreviaEtiqueta(0, valor)
        picEtiquetaLote.ImageLocation = imgEtiqueta(0)
        picEtiquetaLote.Width = Integer.Parse(imgEtiqueta(1))
        picEtiquetaLote.Height = Integer.Parse(imgEtiqueta(2))
    End Sub
    Public Sub rellenarCamposPedido()
        'procedimiento de rellenar los campos del pedido.

        If dt.Rows.Count > 0 Then
            If Not dt.Rows(0)("Pedido") Is Nothing Then txtNumPedido.Text = dt.Rows(0)("Pedido").ToString
            If Not dt.Rows(0)("Cliente") Is Nothing Then txtNombreCliente.Text = dt.Rows(0)("Cliente")
            If Not dt.Rows(0)("jeringasPorExpositor") Then numUnidadesPorlote.Value = dt.Rows(0)("jeringasPorExpositor")
            If Not dt.Rows(0)("jeringasPorCaja") Then numLotesPorCaja.Value = (dt.Rows(0)("jeringasPorCaja") / dt.Rows(0)("jeringasPorExpositor"))
            If Not dt.Rows(0)("refProducto") Is Nothing Then txtRefProducto.Text = dt.Rows(0)("refProducto")
            If Not dt.Rows(0)("Artículo") Is Nothing Then txtNombProducto.Text = dt.Rows(0)("Artículo")
            If Not dt.Rows(0)("pedidoSap") Is Nothing Then txtnumsap.Text = dt.Rows(0)("pedidoSap")
        Else
                txtNumPedido.Text = ""
            txtNombreCliente.Text = ""
            numUnidadesPorlote.Value = 10
            numLotesPorCaja.Value = 5
            txtRefProducto.Text = ""
            txtNombProducto.Text = ""
        End If
    End Sub
    Public Sub rellenarCamposRango()
        'procedimiento de rellenar los campos de Rango de Chip.
        Dim numchips As Integer

        If dt1.Rows.Count > 0 Then
            txtGlobalFirstCode.Text = dt1.Rows(0)("Inicio").ToString
            txtLastCodigoChip.Text = dt1.Rows(0)("Final").ToString
            numchips = (dt1.Rows(0)("Final") - dt1.Rows(0)("Inicio")) + 1
            numLotes.Value = numchips / dt.Rows(0)("jeringasPorExpositor")
            'variable donde guardamos la cantidad de lotes que compone el pedido. No se modifica.
            totalLotes = numLotes.Value
            numCajas.Value = Math.Round(numchips / dt.Rows(0)("jeringasPorCaja"))
            'variable donde guardamos la cantidad de cajas que compone el pedido. No se modifica.
            maxcajas = numCajas.Value
            'calculamos el numero de Palets a imprimir
            'variable donde guardamos la cantidad de palets que compone el pedido. No se modifica.
            palets = Math.Round(numCajas.Value / numCajasPorPalet.Value)
            If palets < 1 Then palets = 1 ' si es menor de uno lo redondeamos por que siempre se tiene que imprimir 1.
            numPalets.Value = palets
            txtSiglaPais.Text = sacarSiglaPais(Integer.Parse(Strings.Left(txtGlobalFirstCode.Text, 3)))
        Else
            txtGlobalFirstCode.Text = ""
            txtLastCodigoChip.Text = ""
            numLotes.Value = 1
            totalLotes = numLotes.Value
            numCajas.Value = 1
            maxcajas = numCajas.Value
            numPalets.Value = 1
            palets = numPalets.Value
            txtSiglaPais.Text = ""
        End If
    End Sub
    Public Sub calcularChipInicial(valor As Byte)
        Dim chipInicial, tempChipInicial As Long
        Dim loteInicial As Integer
        'convertimos el numero de chip inicial, caja inicial y lote inicial a numero para poder calcular.
        chipInicial = Long.Parse(txtGlobalFirstCode.Text)
        loteInicial = txtLote.Value

        If valor = 0 Then
            ' si es 0 hay que calcular el nº de chip inicial para lote.
            If txtCaja.Value = CajaInicial Then
                tempChipInicial = (chipInicial + (jerxLote * (loteInicial - 1)))
            Else

                chipInicial = (chipInicial + ((jerxLote * lotxCaja) * (txtCaja.Value - CajaInicial)))
                tempChipInicial = (chipInicial + (jerxLote * (loteInicial - 1)))
            End If
        Else
            ' si es 1 hay que calcular el nº de chip inicial para caja.
            tempChipInicial = (chipInicial + ((jerxLote * lotxCaja) * (txtCaja.Value - CajaInicial)))
        End If
        tempFirtCode = tempChipInicial.ToString
    End Sub
    Public Function ultimaCaja(semana As Integer, año As Integer) As Integer
        Dim numCajaini As Integer
        Dim uspcaja As New DataSet1TableAdapters.HistImpEtiqDMTableAdapter
        Dim cajaFinal As DataTable
        cajaFinal = uspcaja.GetData()
        If cajaFinal.Rows.Count > 0 AndAlso semana = cajaFinal.Rows(0)(3) AndAlso año = cajaFinal.Rows(0)(4) Then
            'Si la impresion de etiquetas es en la misma semana y año que las ultimas impresas calculamos la caja inicial de impresión.
            numCajaini = (cajaFinal.Rows(0)(0) + cajaFinal.Rows(0)(1))
        Else
            'Si es en diferente semana que las ultimas impresas ponemos el contador a 1
            numCajaini = 1

        End If
        Return numCajaini
    End Function
    Public Function pedidoImpreso(npedido As Integer) As String
        Dim resultado As Integer
        Dim encontrado As String = ""
        Dim uspbuscar As New DataSet1TableAdapters.uspBuscarPedidoHistoricoTableAdapter
        Dim pedidoEncontrado As DataTable
        pedidoEncontrado = uspbuscar.GetData(npedido, resultado)
        If pedidoEncontrado.Rows.Count > 0 Then
            encontrado = pedidoEncontrado.Rows(0)(11)
        End If
        Return encontrado
    End Function
    Public Sub modificarPedidoImpreso(npedido As Integer)
        Dim resultado As Integer
        Dim uspbuscar As New DataSet1TableAdapters.uspBuscarPedidoHistoricoTableAdapter
        Dim pedidoEncontrado As DataTable
        pedidoEncontrado = uspbuscar.GetData(npedido, resultado)
        reImprimir = True
        etiManual = True
        txtSemana.Value = pedidoEncontrado.Rows(0)("Nsemana")
        txtAño.Value = pedidoEncontrado.Rows(0)("Naño")
        txtCaja.Value = pedidoEncontrado.Rows(0)("CajaInicial")
        CajaInicial = txtCaja.Value
        TotalCajas = (CajaInicial + numCajas.Value)
        ListBox1.SelectedIndex = ListBox1.FindString(Trim(pedidoEncontrado.Rows(0)("TipoEtiqLote")))
        ListBox2.SelectedIndex = ListBox2.FindString(Trim(pedidoEncontrado.Rows(0)("TipoEtiqCaja")))
        cmImpManual_Click(Nothing, Nothing)


        If resultado = 1 Then MsgBox("Error al extraer datos en la BBDD", MsgBoxStyle.Information)
    End Sub

#End Region
    Private Sub cmdImprimirLotes_Click(sender As Object, e As EventArgs) Handles cmdImprimirLotes.Click
        'ImprimirLotes()
        Dim i As Integer
        loteinicial = Integer.Parse(txtLote.Text)
        auxcaja = Integer.Parse(txtCaja.Value)

        Select Case ListBox1.SelectedItem

            Case "Std", "tipo12"
                fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                For i = txtLote.Value To (numLotes.Value + i) - 1
                    numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                    printLoteStandart()
                    loteinicial += 1
                    If loteinicial = (lotxCaja + 1) Then
                        loteinicial = 1
                        auxcaja += 1
                    End If
                Next
            Case "tipo5"
                fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                For i = txtLote.Value To (numLotes.Value + i) - 1
                    numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                    printLoteTipo5()
                    loteinicial += 1
                    If loteinicial = (lotxCaja + 1) Then
                        loteinicial = 1
                        auxcaja += 1
                    End If
                Next
            Case "tipo15"
                fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                For i = txtLote.Value To (numLotes.Value + txtLote.Value) - 1
                    numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                    printLoteTipo15()
                    loteinicial += 1
                    If loteinicial = (lotxCaja + 1) Then
                        loteinicial = 1
                        auxcaja += 1
                    End If
                Next
            Case "report35"
                fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                'si la etiqueta es manual y numero de lote no es 1 se tiene que calcula el chip inicial de lote
                If etiManual = True AndAlso txtLote.Value <> 1 Then calcularChipInicial(0)

                For i = txtLote.Value To (numLotes.Value + txtLote.Value) - 1
                    numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                    printReport35()
                    loteinicial += 1
                    If loteinicial = (lotxCaja + 1) Then
                        loteinicial = 1
                        auxcaja += 1
                    End If
                Next
            Case "report37"
                'si la etiqueta es manual y numero de lote no es 1 se tiene que calcula el chip inicial de lote
                If etiManual = True Then calcularChipInicial(0)
                For i = txtLote.Value To (numLotes.Value + txtLote.Value) - 1
                    numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                    printReport37()
                    loteinicial += 1
                    If loteinicial = (lotxCaja + 1) Then
                        loteinicial = 1
                        auxcaja += 1
                    End If
                Next
            Case "report39", "report43"
                fecCaducidad = "EXP.:  20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                For i = txtLote.Value To (numLotes.Value + txtLote.Value) - 1
                    numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                    printReport39()
                    loteinicial += 1
                    If loteinicial = (lotxCaja + 1) Then
                        loteinicial = 1
                        auxcaja += 1
                    End If
                Next

        End Select
        If histGrabado = False Then
            grabarHistorico("insertar", 0)
            histGrabado = True
        Else
            grabarHistorico("actualizar", 2)
        End If
        cerrarBartender()
    End Sub

    Private Sub cmdImprimirCajas_Click(sender As Object, e As EventArgs) Handles cmdImprimirCajas.Click
        'ImprimirCajas()
        Dim auxCaja As Integer = Integer.Parse(txtCaja.Value)
        Select Case ListBox2.SelectedItem
            Case "Std"
                For i As Integer = 0 To (numCajas.Value) - 1
                    numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                    printCajaStandart()
                    auxCaja += 1
                Next
            Case "tipo3"
                fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                For i As Integer = 0 To (numCajas.Value) - 1
                    numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                    printCajaTipo3()
                    auxCaja += 1
                Next
            Case "tipo5"
                For i As Integer = 0 To (numCajas.Value) - 1
                    numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                    printCajaTipo5()
                    auxCaja += 1
                Next
            Case "tipo11"
                fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                For i As Integer = 0 To (numCajas.Value) - 1
                    numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                    printCajaTipo11()
                    auxCaja += 1
                Next
            Case "tipo15"
                MsgBox("Este cliente Se imprimen 2 Etiquetas de Diferentes Tamaños. Las Primeras de 60x47mm y Las segundas de 80x80mm_
                        _ Asegurese de tener las etiquetas correctas", MsgBoxStyle.Information)
                For i As Integer = 0 To (numCajas.Value) - 1
                    numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                    printCajaTipo15()
                    auxCaja += 1
                Next
                mostrarEtiqCaja("tipo12")
                MsgBox("!Se van ha Imprimir las Etiquetas de 80x80mm¡, Cambie el tamaño de las etiquetas", MsgBoxStyle.Information)

                fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                For i As Integer = 0 To (numCajas.Value) - 1
                    numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                    printCajaTipo12()
                    auxCaja += 1
                Next
            Case "report36", "report42", "report45"
                Dim numLoteTmp As Integer
                fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                'si la etiqueta es manual y numero de caja no es 1 se tiene que calcula el chip inicial de caja
                If etiManual = True Then calcularChipInicial(1)
                'asignamos la cantidad de lotes a una variable temporal.
                numLoteTmp = numLotes.Value
                For i As Integer = 0 To (numCajas.Value) - 1
                    numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                    If (jerxLote * lotxCaja) > (jerxLote * ((numLoteTmp + txtLote.Value) - 1)) Then lotxCaja = ((numLoteTmp + txtLote.Value) - 1)
                    'restamos a la variable de lotes temporales la cantidad de lotes por caja. Esto se hace por si la ultima caja hay menos lotes que los reglamentarios.
                    numLoteTmp = numLoteTmp - lotxCaja
                    If ListBox2.SelectedItem = "report36" Then
                        'imprimimos report36
                        printCajaReport36(0)
                    ElseIf ListBox2.SelectedItem = "report42" Then
                        'imprimimos report42 es igual que el 36, pero cambia el texto de descripcion.
                        printCajaReport36(1)
                    Else
                        'imprimimos report45 es igual que el 36, pero cambia el texto de descripción.
                        printCajaReport36(2)
                    End If
                    auxCaja += 1
                Next
                lotxCaja = numLotesPorCaja.Value
            Case "report38"
                Dim numLoteTmp As Integer
                'si la etiqueta es manual y numero de caja no es 1 se tiene que calcula el chip inicial de caja
                If etiManual = True AndAlso txtCaja.Value <> 1 Then calcularChipInicial(1)
                'asignamos la cantidad de lotes a una variable temporal.
                numLoteTmp = numLotes.Value
                For i As Integer = 0 To (numCajas.Value) - 1
                    numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                    If (jerxLote * lotxCaja) > (jerxLote * ((numLoteTmp + txtLote.Value) - 1)) Then lotxCaja = ((numLoteTmp + txtLote.Value) - 1)
                    'restamos a la variable de lotes temporales la cantidad de lotes por caja. Esto se hace por si la ultima caja hay menos lotes que los reglamentarios salga bien la numeración.
                    numLoteTmp = numLoteTmp - lotxCaja
                    printCajaReport38()
                    auxCaja += 1
                Next
                lotxCaja = numLotesPorCaja.Value
            Case "report40", "report44"
                fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                For i As Integer = 0 To (numCajas.Value) - 1
                    numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                    printCajaReport40()
                    auxCaja += 1
                Next
        End Select
        If histGrabado = False Then
            grabarHistorico("insertar", 1)
            histGrabado = True
        Else
            grabarHistorico("actualizar", 2)
        End If
        cerrarBartender()
    End Sub
    Private Sub cmdImprimirPalets_Click(sender As Object, e As EventArgs) Handles cmdImprimirPalets.Click
        Dim nbox, nboxfin As Integer
        Dim numcajaini, numcajafin As String
        fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
        nbox = Integer.Parse(txtCaja.Value)

        For i = 1 To numPalets.Value
            numcajaini = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & nbox.ToString("0000")
            nboxfin = (nbox + numCajasPorPalet.Value) - 1
            numcajafin = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & nboxfin.ToString("0000")
            printEtiPalet(txtnumsap.Text, i, numcajaini, numcajafin)
            nbox = nboxfin
        Next
    End Sub


#Region "Carga de Variables"
    Private Sub txtSiglaPais_TextChanged(sender As Object, e As EventArgs) Handles txtSiglaPais.TextChanged
        sigla = txtSiglaPais.Text
    End Sub

    Private Sub txtNombProducto_TextChanged(sender As Object, e As EventArgs) Handles txtNombProducto.TextChanged
        descArticulo = txtNombProducto.Text
    End Sub

    Private Sub txtRefProducto_TextChanged(sender As Object, e As EventArgs) Handles txtRefProducto.TextChanged
        refArticulo = txtRefProducto.Text
    End Sub

    Private Sub txtGlobalFirstCode_TextChanged(sender As Object, e As EventArgs) Handles txtGlobalFirstCode.TextChanged
        firstCode = txtGlobalFirstCode.Text
        tempFirtCode = txtGlobalFirstCode.Text
    End Sub

    Private Sub txtLastCodigoChip_TextChanged(sender As Object, e As EventArgs) Handles txtLastCodigoChip.TextChanged
        lastCode = txtLastCodigoChip.Text
        tempLastCode = txtLastCodigoChip.Text
    End Sub

    Private Sub txtNumPedido_TextChanged(sender As Object, e As EventArgs) Handles txtNumPedido.TextChanged
        numPedido = txtNumPedido.Text
        dt1 = uspRangoChip.GetData(txtNumPedido.Text)
        rellenarCamposRango()
    End Sub

    Private Sub numLotes_ValueChanged(sender As Object, e As EventArgs) Handles numLotes.ValueChanged
        'cantLote = numLotes.Value
    End Sub

    Private Sub numCajas_ValueChanged(sender As Object, e As EventArgs) Handles numCajas.ValueChanged
        'cantCaja = numCajas.Value
    End Sub

    Private Sub numLotesPorCaja_ValueChanged(sender As Object, e As EventArgs) Handles numLotesPorCaja.ValueChanged
        ' lotxCaja = numLotesPorCaja.Value
    End Sub



    Private Sub numUnidadesPorlote_ValueChanged(sender As Object, e As EventArgs) Handles numUnidadesPorlote.ValueChanged
        ' jerxLote = numUnidadesPorlote.Value
    End Sub
#End Region


    Private Sub txtCodPedido_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCodPedido.KeyDown
        'si pulsamos la tecla enter o introducimos un salto de linea ejecutamos el codigo.
        If e.KeyCode = Keys.Enter Or e.KeyValue = Asc(13) Then

            cargarPedido()

        End If
    End Sub


    Private Sub txtLote_ValueChanged(sender As Object, e As EventArgs) Handles txtLote.ValueChanged
        If numLotesPorCaja.Value <> 0 Then
            If txtLote.Value > numLotesPorCaja.Value Then
                MsgBox("El Numero de Lote Inicial no puede ser Mayor Que los lotes por Caja", MsgBoxStyle.Information)
                txtLote.Value = numLotesPorCaja.Value
            Else
                Dim st As New System.Diagnostics.StackTrace()
                Dim nombreMetodo As String
                nombreMetodo = st.GetFrame(3).GetMethod().Name
                Select Case nombreMetodo
                    Case "UpButton"
                        numLotes.Value -= 1
                    Case "DownButton"
                        numLotes.Value += 1
                    Case "ParseEditText"
                        Dim lotes As Integer
                        lotes = (totalLotes - (txtCaja.Value * numLotesPorCaja.Value)) + ((numLotesPorCaja.Value) - txtLote.Value) + 1
                        numLotes.Value = lotes
                End Select
            End If

        End If
    End Sub

    Private Sub txtSemana_ValueChanged(sender As Object, e As EventArgs) Handles txtSemana.ValueChanged
        CajaInicial = ultimaCaja(txtSemana.Value, txtAño.Value)
        txtCaja.Value = CajaInicial
    End Sub

    Private Sub txtCaja_ValueChanged(sender As Object, e As EventArgs) Handles txtCaja.ValueChanged

        Dim st As New System.Diagnostics.StackTrace()
        Dim nombreMetodo As String
        nombreMetodo = st.GetFrame(3).GetMethod().Name
        Try
            Select Case nombreMetodo
                Case "UpButton"
                    'si pulsamos la flecha UP de la caja numerica para sumar 
                    numLotes.Value = (numLotes.Value - numLotesPorCaja.Value)
                    numCajas.Value -= 1
                    If (txtCaja.Value + cantCaja) < TotalCajas Then 'Or txtCaja.Value > (CajaInicial + numCajas.Value)
                        MsgBox("El Numero de Caja Inicial No puede ser Mayor al Total de Cajas a Imprimir", MsgBoxStyle.Information)
                        numLotes.Value = (numLotes.Value + numLotesPorCaja.Value)
                        numCajas.Value += 1
                        txtCaja.Value -= 1
                    End If
                Case "DownButton"
                    'si pulsamos la flecha Down de la caja numerica para restar
                    numLotes.Value = (numLotes.Value + numLotesPorCaja.Value)
                    numCajas.Value += 1
                    If txtCaja.Value < CajaInicial Then
                        MsgBox("El numero de Caja No puede ser Inferior al Numero de Caja Inicial", MsgBoxStyle.Information)
                        numLotes.Value = (numLotes.Value - numLotesPorCaja.Value)
                        numCajas.Value -= 1
                        txtCaja.Value += 1
                    End If
                Case "ParseEditText"
                    'si ponemos el numero  a mano.
                    If txtCaja.Value > TotalCajas Then
                        MsgBox("El Numero de Caja Inicial No puede ser Mayor al Total de Cajas a Imprimir", MsgBoxStyle.Information)
                        txtCaja.Value = CajaInicial
                    Else
                        Dim ncajImp As Integer
                        ncajImp = (TotalCajas - txtCaja.Value)
                        numCajas.Value = ncajImp
                        numLotes.Value = (totalLotes - (numLotesPorCaja.Value * (txtCaja.Value - 1)))
                    End If

            End Select
        Catch ex As Exception

            If ex.HResult = -2146233086 Then
                MsgBox("Tiene Que Imprimir Como Minimo 1 Caja y 1 Lote del pedio Actual")
                txtCaja.Value -= 1

            End If

        End Try

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        mostrarEtiqLote(ListBox1.SelectedItem)
    End Sub

    Private Sub cmImpManual_Click(sender As Object, e As EventArgs) Handles cmdImpManual.Click
        etiManual = True
        txtSemana.Enabled = True
        txtCaja.Enabled = True
        txtLote.Enabled = True
        numCajasPorPalet.Value = True
        numLotes.Enabled = True
        numCajas.Enabled = True
        numPalets.Enabled = True
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        mostrarEtiqCaja(ListBox2.SelectedItem)
    End Sub

End Class
