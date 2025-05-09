Public Class Form1

    Public uspDatosPedidos As New DataSet1TableAdapters.conPedidosRefSapDesartTableAdapter
    Public uspRangoChip As New DataSet1TableAdapters.OrdenesTableAdapter
    Public dt, dt1 As DataTable
    Public etiManual, histGrabado, reImprimir, añadirDosPorCiento As Boolean
    'Public IDhistorico, CajaInicial, TotalCajas, totalLotes As Integer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Cargar_Conexiones_SQL()
        If compBloqueo = False Then comprobarBloqueo()
        etiManual = False
        histGrabado = False
        reImprimir = False
        txtAño.Value = Now.ToString("yy")
        'txtAño.Minimum = Now.ToString("yy") - 1
        'txtAño.Maximum = Now.ToString("yy") + 1

        numAñoCad.Value = Now.ToString("yy") + 3
        'numAñoCad.Minimum = Now.ToString("yy") - 1 + 3
        'numAñoCad.Maximum = Now.ToString("yy") + 1 + 3

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
        txtPaletImp.Text = 0
        txtCajasImp.Text = 0
        txtlotesimp.Text = 0
        fecFabricacion = "20" & txtAño.Value & "-" & numMesCad.Value.ToString("00")
        ListBox1.SelectedIndex = -1
        ListBox2.SelectedIndex = -1
        'Bloqueamos los botones de imprimir hasta que no se introduzca un pedido.
        cmdImprimirCajas.Enabled = False
        cmdImprimirLotes.Enabled = False
        cmdImpManual.Enabled = False
        cmdImprimirPalets.Enabled = False
        'Cargamos la impresora por defecto
        Dim ImpDefecto As New System.Drawing.Printing.PrinterSettings
        cListImp.Text = INIRead(My.Application.Info.DirectoryPath & "\settings.ini", My.Computer.Name, "Impresora", ImpDefecto.PrinterName)
        'deshabilitamos la hibernación del equipo
        suspension.deshibernacion()
    End Sub

    Private Sub Cargar_Conexiones_SQL()

        My.Settings.Item("PartesMicrochipConnectionString") = INIRead(My.Application.Info.DirectoryPath & "\settings.ini", "Conexion", "PartesMicrochip", "")
        My.Settings.Item("MicrochipsConnectionString") = INIRead(My.Application.Info.DirectoryPath & "\settings.ini", "Conexion", "Microchips", "")
        My.Settings.Item("OrdenesConnectionString") = INIRead(My.Application.Info.DirectoryPath & "\settings.ini", "Conexion", "Ordenes", "")
        My.Settings.Item("grab_genericaSQLConnectionString") = INIRead(My.Application.Info.DirectoryPath & "\settings.ini", "Conexion", "grab_genericaSQL", "")
        My.Settings.Item("PartesMicrochipConnectionStringSRV3") = INIRead(My.Application.Info.DirectoryPath & "\settings.ini", "Conexion", "PartesMicrochipSRV3", "")
        My.Settings.Item("PID_QueueConnectionString") = INIRead(My.Application.Info.DirectoryPath & "\settings.ini", "Conexion", "PID_Queue", "")
        My.Settings.Item("grab_genericaSQLConnectionString1") = INIRead(My.Application.Info.DirectoryPath & "\settings.ini", "Conexion", "grab_genericaSQLSRV3", "")


    End Sub

#Region "Funciones y Procedimientos "

    Function sacarSiglaPais(valor As String) As String
        'funcion que saca las siglas del pais donde va el microchip.
        Dim uspsigla As New DataSet1TableAdapters.CodigosPaisesTableAdapter
        Dim dt3 As DataTable
        Dim dato As String
        dt3 = uspsigla.GetData(valor)
        dato = dt3.Rows(0)("A 3")
        Return dato
    End Function
    Private Sub activarBotones()
        'Si la etiqueta es Ninguna deshabilitamos el boton de imprimir etiqueta.
        cmdImprimirLotes.Enabled = True
        cmdImprimirCajas.Enabled = True
        cmdImpManual.Enabled = True
        cmdImprimirPalets.Enabled = True
        btn_MuestraLotes.Enabled = True
        btn_Muestra_Cajas.Enabled = True
        btn_Muestra_Palets.Enabled = True
    End Sub
    Private Function cargarPedido() As Boolean
        'procedimiento de busqueda de Pedido en la BBDD.
        Dim respuesta As Integer
        Dim etiquetasImpresas As String
        Dim retorno As Boolean = False
        respuesta = 0
        If etiManual = True Then
            txtSemana.Enabled = False
            txtCaja.Enabled = False
            txtLote.Enabled = False
            txtpalet.Enabled = False
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
        If etiquetasImpresas <> "" Then respuesta = MsgBox("Pedido impreso el " & etiquetasImpresas & ". ¿Quieres reimprimirlo?", MsgBoxStyle.YesNo)
        If etiquetasImpresas = "" Or etiquetasImpresas <> "" AndAlso respuesta = 6 Then
            'si el pedido es nuevo o Ya impreso y queremos modificarlo
            dt = uspDatosPedidos.GetData(txtCodPedido.Text)
            rellenarCamposPedido()
            reImprimir = False
            If dt.Rows.Count > 0 Then
                If respuesta = 0 Then
                    'Para pedido nuevo
                    'calculamos la caja inicial
                    CajaInicial = ultimaCaja(txtSemana.Value, txtAño.Value)
                    txtCaja.Value = CajaInicial
                    TotalCajas = (CajaInicial + numCajas.Value)
                    ListBox1.SelectedIndex = ListBox1.FindString(txtBDEtiLot.Text.Trim)

                    ListBox2.SelectedIndex = ListBox2.FindString(txtBDEtiCaja.Text.Trim)
                Else
                    'Reimprimir 
                    modificarPedidoImpreso(txtCodPedido.Text)
                End If

                retorno = True
            Else
                MessageBox.Show("El pedido no existe. Compruébelo", "ATENCIÓN", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                limpiar()
                retorno = False
            End If
        Else
            txtCodPedido.Text = " "
            limpiar()
            retorno = False
        End If
        Return retorno
    End Function
    Private Sub limpiar()
        Form1_Load(Nothing, Nothing)
    End Sub
    Function grabarHistorico(tipoGrab As String) As Boolean
        'Procedimiento para grabar datos
        Dim uspInsertarHisto As New DataSet1TableAdapters.qryInsertHistorico
        Dim equipo As String
        Dim respuesta As Boolean = False

        equipo = SystemInformation.ComputerName
        Dim grabCorrecta As Integer
        If tipoGrab = "insertar" Then
            If ListBox2.SelectedItem Is Nothing Then
                uspInsertarHisto.uspInsertarHistoricoEtiquetasDM(txtCodPedido.Text, txtSemana.Value, txtAño.Value, empezarXCaja, numCajas.Value, numLotes.Value, numPalets.Value, txtlotesimp.Text, txtCajasImp.Text, txtPaletImp.Text,
                                                         ListBox1.SelectedItem, "", Now, equipo, usuarioactual.getID, reImprimir, grabCorrecta, IDhistorico)
            Else
                uspInsertarHisto.uspInsertarHistoricoEtiquetasDM(txtCodPedido.Text, txtSemana.Value, txtAño.Value, empezarXCaja, numCajas.Value, numLotes.Value, numPalets.Value, txtlotesimp.Text, txtCajasImp.Text, txtPaletImp.Text,
                                                         ListBox1.SelectedItem, ListBox2.SelectedItem, Now, equipo, usuarioactual.getID, reImprimir, grabCorrecta, IDhistorico)
            End If
        ElseIf tipoGrab = "actualizar" Then
            uspInsertarHisto.uspActualizarHistoricoEtiquetas(IDhistorico, txtlotesimp.Text, txtCajasImp.Text, txtPaletImp.Text, Now, grabCorrecta)
        End If
        If grabCorrecta = 1 Then
            MsgBox("Se ha producido un error al grabar en la BB.DD.", MsgBoxStyle.Information)
            respuesta = False
        Else
            respuesta = True
        End If
        Return respuesta
    End Function


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
    Sub mostrarEtiqPalet(valor As String)
        Dim imgEtiqueta As String()
        imgEtiqueta = vistaPreviaEtiqueta(3, valor)
        picEtiquetaPalet.ImageLocation = imgEtiqueta(0)
        picEtiquetaPalet.Width = Integer.Parse(imgEtiqueta(1))
        picEtiquetaPalet.Height = Integer.Parse(imgEtiqueta(2))
    End Sub
    Public Sub rellenarCamposPedido()
        'procedimiento de rellenar los campos del pedido.

        If dt.Rows.Count > 0 Then
            If Not dt.Rows(0)("cantidad") Is Nothing Then txtCantidad.Text = dt.Rows(0)("cantidad")
            If Not dt.Rows(0)("tipoProducto") Is Nothing Then tipoProducto = dt.Rows(0)("tipoProducto")
            If Not dt.Rows(0)("Pedido") Is Nothing Then txtNumPedido.Text = dt.Rows(0)("Pedido").ToString
            If Not dt.Rows(0)("Cliente") Is Nothing Then txtNombreCliente.Text = dt.Rows(0)("Cliente")
            If Not dt.Rows(0)("jeringasPorExpositor") Then numUnidadesPorlote.Value = dt.Rows(0)("jeringasPorExpositor")
            If Not dt.Rows(0)("jeringasPorCaja") Then
                Dim valor As Integer
                valor = (dt.Rows(0)("jeringasPorCaja") / dt.Rows(0)("jeringasPorExpositor"))
                If valor <= 0 Then valor = 1
                numLotesPorCaja.Value = valor
            End If
            If Not dt.Rows(0)("refProducto") Is Nothing Then txtRefProducto.Text = dt.Rows(0)("refProducto")
            If Not dt.Rows(0)("Artículo") Is Nothing Then txtNombProducto.Text = dt.Rows(0)("Artículo")
            If Not IsDBNull(dt.Rows(0)("pedidoSap")) Then txtnumsap.Text = dt.Rows(0)("pedidoSap")
            If Not dt.Rows(0)("notasEtiquetasExpositor") Is Nothing Then txtBDEtiLot.Text = dt.Rows(0)("notasEtiquetasExpositor")
            If Not dt.Rows(0)("notasTipoEtiquetaCaja") Is Nothing Then txtBDEtiCaja.Text = dt.Rows(0)("notasTipoEtiquetaCaja")
            If Not IsDBNull(dt.Rows(0)("RefCliente")) Then txtRefCliente.Text = dt.Rows(0)("RefCliente")
            If Not IsDBNull(dt.Rows(0)("nave")) Then txtLugarFab.Text = dt.Rows(0)("nave")
            If tipoProducto = 2 And txtBDEtiCaja.Text = "" Then
                txtBDEtiCaja.Text = "Std"
            End If
            'Comprobamos si el pedido es mayor de cierta cantidad de jeringas para poder enviar todo el rango.
            Dim qtyRedondeo As Integer
            If Not (txtRefProducto.Text.Trim = "996 0000-AR2" Or txtRefProducto.Text.Trim = "996 0925-JPN" Or txtRefProducto.Text.Trim = "992 4001-JPN") Then
                qtyRedondeo = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "REDONDEO", "cantidad")
                If Integer.Parse(txtCantidad.Text) >= qtyRedondeo Then
                    ' verificamos que la cantidad del pedido y la cantidad del rango, si el rango es mayor esa es la cantidad del pedido que debemos usar.
                    comprobarSobrantes()
                    If Integer.Parse(cantRango) > Integer.Parse(txtCantidad.Text) Then
                        If Not reImprimir Then MsgBox("Pedido Con Sobrantes. Se cambiará la Cantidad del Pedido al Rango: " & cantRango)
                        txtCantidad.Text = cantRango
                    End If
                End If
            End If
            'llamamos al procedimiento para calcular la cantidad de cajas, lotes y palets a imprimir.
            calculoCantidadLCP()
            If IsDBNull(dt.Rows(0)("fechaInicioReal")) Then
                fecInicioReal = ""
            Else
                fecInicioReal = dt.Rows(0)("fechaInicioReal").ToString
            End If
        Else
            txtNumPedido.Text = ""
            txtNombreCliente.Text = ""
            numUnidadesPorlote.Value = 10
            numLotesPorCaja.Value = 5
            txtRefProducto.Text = ""
            txtNombProducto.Text = ""
            numLotes.Value = 1
            totalLotes = numLotes.Value
            numCajas.Value = 1
            maxcajas = numCajas.Value
            numPalets.Value = 1
            palets = numPalets.Value
        End If
    End Sub

    Public Sub comprobarSobrantes()
        añadirDosPorCiento = False
        Dim workid As New DataSet1TableAdapters.uspGetIdSobrantesTableAdapter
        Dim sobra As Integer = workid.GetData(txtCodPedido.Text).Count
        Dim añadir As Integer = 2 * Integer.Parse(txtCantidad.Text) / 100

        If sobra > 0 Then
            ImprimirSobrantes.ShowDialog()
            If añadirDosPorCiento = True Then
                añadirDosPorCiento = False
                cantRango = Integer.Parse(txtCantidad.Text) + añadir
            End If
        End If
    End Sub

    Public Sub calculoCantidadLCP()
        'procedimiento para calcular la cantidad de cajas,lotes y palets a imprimir.
        Dim numchips, nCajasTemp As Integer

        numchips = txtCantidad.Text.Trim
        'numLotes.Value = Math.Round(numchips / dt.Rows(0)("jeringasPorExpositor"))
        'con esta funcion siempre redondea al entero superior.
        numLotes.Value = Math.Ceiling(numchips / dt.Rows(0)("jeringasPorExpositor"))
        'variable donde guardamos la cantidad de lotes que compone el pedido. No se modifica.
        totalLotes = numLotes.Value
        If dt.Rows(0)("jeringasPorCaja") = 0 Then
            'Si la cantidad de jeringas por caja es 0, dividimos por la cantidad del expositor. Para pedidos agranel.
            'nCajasTemp = Math.Round((numchips / dt.Rows(0)("jeringasPorExpositor")) + 0.1)
            nCajasTemp = Math.Ceiling((numchips / dt.Rows(0)("jeringasPorExpositor")))
        Else
            'nCajasTemp = Math.Round((numchips / dt.Rows(0)("jeringasPorCaja")) + 0.1)
            nCajasTemp = Math.Ceiling((numchips / dt.Rows(0)("jeringasPorCaja")))
        End If

        If nCajasTemp < 1 Then nCajasTemp = 1 'si es menor de 1 lo redondeamos a la cantidad minima.
        'variable donde guardamos la cantidad de cajas que compone el pedido. No se modifica.
        numCajas.Value = nCajasTemp
        maxcajas = numCajas.Value
        'calculamos el numero de Palets a imprimir
        'variable donde guardamos la cantidad de palets que compone el pedido. No se modifica.
        palets = Math.Ceiling(numCajas.Value / numCajasPorPalet.Value)
        If palets < 1 Then palets = 1 ' si es menor de uno lo redondeamos por que siempre se tiene que imprimir 1.
        numPalets.Value = palets

    End Sub

    Public Sub rellenarCamposRango()
        'procedimiento de rellenar los campos de Rango de Chip.

        If dt1.Rows.Count > 0 Then
            txtGlobalFirstCode.Text = dt1.Rows(0)("Inicio").ToString
            If txtGlobalFirstCode.Text.Length < 15 Then txtGlobalFirstCode.Text = rellenarCero(txtGlobalFirstCode.Text, txtGlobalFirstCode.Text.Length)
            txtLastCodigoChip.Text = dt1.Rows(0)("Final").ToString
            If txtLastCodigoChip.Text.Length < 15 Then txtLastCodigoChip.Text = rellenarCero(txtLastCodigoChip.Text, txtLastCodigoChip.Text.Length)
            cantRango = (dt1.Rows(0)("Final") - dt1.Rows(0)("Inicio")) + 1

            txtSiglaPais.Text = sacarSiglaPais(Strings.Left(txtGlobalFirstCode.Text, 3))
        Else
            'Introducimos el rango manualmente.
            txtGlobalFirstCode.Text = InputBox("Introduzca Inicio Rango si el pedido es ordenado, si no ponga 0")
            txtLastCodigoChip.Text = InputBox("Introduzca Fin Rango si el pedido es ordenado, si no ponga 0")

            cantRango = (txtLastCodigoChip.Text - txtGlobalFirstCode.Text) + 1
            ' numLotes.Value = 1
            'totalLotes = numLotes.Value
            'numCajas.Value = 1
            'maxcajas = numCajas.Value
            'numPalets.Value = 1
            'palets = numPalets.Value
            If txtGlobalFirstCode.Text.Trim = 0 Then
                txtSiglaPais.Text = ""
            Else
                txtSiglaPais.Text = sacarSiglaPais(Strings.Left(txtGlobalFirstCode.Text, 3))
            End If

        End If
    End Sub
    Private Function rellenarCero(rango As String, valor As Integer) As String
        Dim newRango As String
        newRango = rango
        For i = valor To 15
            newRango = "0" & newRango
            If newRango.Length = 15 Then Exit For
        Next
        Return newRango
    End Function
    Public Sub calcularChipInicial(valor As Byte)
        Dim chipInicial, tempChipInicial As Long
        Dim loteInicial As Integer
        'convertimos el numero de chip inicial, caja inicial y lote inicial a numero para poder calcular.
        chipInicial = Long.Parse(txtGlobalFirstCode.Text)
        loteInicial = empezarXlote

        If valor = 0 Then
            ' si es 0 hay que calcular el nº de chip inicial para lote.
            If empezarXCaja = CajaInicial Then
                tempChipInicial = (chipInicial + (jerxLote * (loteInicial - 1)))
            Else

                chipInicial = (chipInicial + ((jerxLote * lotxCaja) * (empezarXCaja - CajaInicial)))
                tempChipInicial = (chipInicial + (jerxLote * (loteInicial - 1)))
            End If
            'asignamos a la variable temporal de inicio de lote
            tempFirstCodeLot = tempChipInicial.ToString
        Else
            ' si es 1 hay que calcular el nº de chip inicial para caja.
            tempChipInicial = (chipInicial + ((jerxLote * lotxCaja) * (empezarXCaja - CajaInicial)))
            'asignamos a la variable temporal de inicio de caja
            tempFirtCode = tempChipInicial.ToString
        End If

    End Sub
    Public Function ultimaCaja(semana As Integer, año As Integer) As Integer
        Dim numCajaini As Integer
        Dim uspcaja As New DataSet1TableAdapters.HistImpEtiqDMTableAdapter
        Dim cajaFinal As DataTable
        'Buscamos la caja final y pasamos la semana y el año, para evitar años anteriores.
        cajaFinal = uspcaja.GetData(txtSemana.Value, txtAño.Value)
        If cajaFinal.Rows.Count > 0 AndAlso semana = cajaFinal.Rows(0)(3) AndAlso año = cajaFinal.Rows(0)(4) Then
            Dim ci, cf As Integer
            'Si la impresion de etiquetas es en la misma semana y año que las ultimas impresas calculamos la caja inicial de impresión.
            ci = cajaFinal.Rows(0)(0)
            cf = cajaFinal.Rows(0)(1)
            numCajaini = (ci + cf)
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
            reImprimir = True
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
        ListBox1.SelectedIndex = ListBox1.FindString(pedidoEncontrado.Rows(0)("TipoEtiqLote").trim)
        ListBox1.Enabled = False
        ListBox2.SelectedIndex = ListBox2.FindString(pedidoEncontrado.Rows(0)("TipoEtiqCaja").trim)
        ListBox2.Enabled = False
        cmImpManual_Click(Nothing, Nothing)


        If resultado = 1 Then MsgBox("Error al extraer datos en la BBDD", MsgBoxStyle.Information)
    End Sub

#End Region

    Private Sub Imprimir_Lotes(ByVal num_lotes As Integer, ByVal muestra As Boolean)

        Dim i, l, contador, respuesta As Integer
        Dim grabBBDDok As Boolean
        cmdImprimirLotes.Enabled = False
        If muestra Then
            grabBBDDok = True
        Else
            If histGrabado = False Then
                grabBBDDok = grabarHistorico("insertar")
                histGrabado = True
            Else
                grabBBDDok = grabarHistorico("actualizar")
            End If
        End If
        If grabBBDDok = True Then

            l = empezarXlote
            loteinicial = Integer.Parse(empezarXlote) '(txtLote.Text) se cambia el valor de la caja de texto por la variable
            auxcaja = Integer.Parse(empezarXCaja) '(txtCaja.Value) se cambia el valor de la caja de texto por la variable.
            'cada vez que vayamos a imprimir unas etiquetas inicializamos las variables.
            FirstCodeLot = txtGlobalFirstCode.Text.Trim
            tempFirstCodeLot = txtGlobalFirstCode.Text.Trim

            Select Case ListBox1.SelectedItem

                Case "Std", "tipo12"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    contador = 1
                    'abrimos la etiqueta bartender
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos objeto Etiqueta para el BarTender
                    label = objbt.Formats.Open(ruta & "DML_standart.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteStandart()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'contador += 1
                        'If contador = 18 Then
                        '    respuesta = MsgBox("Corta las etiquetas y cuadra la impresora para que salgan correctas", vbOK)
                        '    contador = 1
                        'End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    'cerramos para liberar espacio
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "tipo27"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    contador = 1
                    'abrimos la etiqueta bartender
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos objeto Etiqueta para el BarTender
                    label = objbt.Formats.Open(ruta & "DML1027.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteTipo27()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'contador += 1
                        'If contador = 18 Then
                        '    respuesta = MsgBox("Corta las etiquetas y cuadra la impresora para que salgan correctas", vbOK)
                        '    contador = 1
                        'End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    'cerramos para liberar espacio
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "Tipo 26"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    contador = 1
                    'abrimos la etiqueta bartender
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos objeto Etiqueta para el BarTender
                    label = objbt.Formats.Open(ruta & "DML1026.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteStandart()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'contador += 1
                        'If contador = 18 Then
                        '    respuesta = MsgBox("Corta las etiquetas y cuadra la impresora para que salgan correctas", vbOK)
                        '    contador = 1
                        'End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    'cerramos para liberar espacio
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "Tipo48"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    contador = 1
                    'abrimos la etiqueta bartender
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos objeto Etiqueta para el BarTender
                    label = objbt.Formats.Open(ruta & "DML1026.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteStandart()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    MsgBox("Se va a imprimir el otro tipo de pegatina, cambie el rollo de etiquetas")
                    label = objbt.Formats.Open(ruta & "DML_CCA.btw")
                    fecCaducidad = "" & numMesCad.Value.ToString("00") & "/20" & numAñoCad.Value.ToString("00")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printTipo26B()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next

                    'cerramos para liberar espacio
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "Tipo 26B"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    contador = 1
                    'abrimos la etiqueta bartender
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos objeto Etiqueta para el BarTender
                    label = objbt.Formats.Open(ruta & "DML1026.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteStandart()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    MsgBox("Se va a imprimir el otro tipo de pegatina, cambie el rollo de etiquetas")
                    label = objbt.Formats.Open(ruta & "DML1026B.btw")
                    fecCaducidad = "" & numMesCad.Value.ToString("00") & "/20" & numAñoCad.Value.ToString("00")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printTipo26B()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next

                    'cerramos para liberar espacio
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "tipo3"
                    fecCaducidad = "01/" & numMesCad.Value.ToString("00") & "/" & "20" & numAñoCad.Value
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos objeto Etiqueta para el BarTender
                    label = objbt.Formats.Open(ruta & "DML1003.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteTipo3()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If

                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)

                Case "tipo5"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    contador = 1
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos objeto Etiqueta para el BarTender
                    label = objbt.Formats.Open(ruta & "DML1005.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteTipo5()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'contador += 1
                        'If contador = 12 Then
                        '    respuesta = MsgBox("Corta las etiquetas y cuadra la impresora para que salgan correctas", vbOK)
                        '    contador = 1
                        'End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)

                Case "tipo7"
                    fecCaducidad = "20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos el objeto etiqueta
                    label = objbt.Formats.Open(ruta & "DML1007.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteTipo7()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If

                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "tipo9"
                    fecCaducidad = "20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos el objeto etiqueta
                    label = objbt.Formats.Open(ruta & "DML1009.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteTipo9()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If

                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "tipo13", "tipo14", "tipo22", "tipo17"
                    fecCaducidad = "EXP.:  20" & (numAñoCad.Value) & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos el objeto impresora
                    If ListBox1.SelectedItem = "tipo13" Then
                        label = objbt.Formats.Open(ruta & "DML1013.btw")
                    ElseIf ListBox1.SelectedItem = "tipo14" Then
                        label = objbt.Formats.Open(ruta & "DML1014.btw")
                    ElseIf ListBox1.SelectedItem = "tipo22" Then
                        label = objbt.Formats.Open(ruta & "DML1022.btw")
                    ElseIf ListBox1.SelectedItem = "tipo17" Then
                        label = objbt.Formats.Open(ruta & "DML1017.btw")
                    End If

                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + empezarXlote) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteTipo13()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "tipo15"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos el objeto etiqueta
                    label = objbt.Formats.Open(ruta & "DML1015.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + empezarXlote) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteTipo15()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "report35"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    'si la etiqueta es manual y numero de lote no es 1 se tiene que calcula el chip inicial de lote
                    If etiManual = True Then calcularChipInicial(0) 'AndAlso empezarXlote <> 1
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos el objeto etiqueta
                    label = objbt.Formats.Open(ruta & "DML1035.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + empezarXlote) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printReport35()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "report37"
                    'si la etiqueta es manual y numero de lote no es 1 se tiene que calcula el chip inicial de lote
                    If etiManual = True Then calcularChipInicial(0)
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    'Creamos el objeto etiqueta
                    label = objbt.Formats.Open(ruta & "DML1037.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + empezarXlote) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printReport37()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "report39", "report43", "Tipo28"
                    fecCaducidad = "EXP.:  20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")

                    Dim nombreEtiqueta As String

                    'Si el pedido es de Kyoritsu cambiamos el formato de la fecha de caducidad y ponemos el numero de esterilización.
                    If txtRefProducto.Text.Trim = "996 0925-JPN" Or txtRefProducto.Text.Trim = "992 4001-JPN" Then
                        Dim expRegular, expAño As String
                        Dim cadValida As Boolean = False
                        expAño = "(" & txtAño.Value & "|" & (txtAño.Value - 1) & ")"
                        expRegular = "^(" & expAño & "(0[1-9]|1[012])([0-2][0-9]|3[01]))/(\d\d[A-D])$"
                        fecCaducidad = "EXP   20" & numAñoCad.Value & "." & numMesCad.Value.ToString("00")
                        'fecCaducidad = "" & numMesCad.Value.ToString("00") & "/20" & numAñoCad.Value.ToString("00")
                        'Do While cadValida = False
                        '    numPedido = txtJaponBatch.Text
                        '    'numPedido = InputBox("Este Pedido se tiene que Introducir el certificado de Esterilización", "Certificado Esterilizacion").ToUpper
                        '    If System.Text.RegularExpressions.Regex.IsMatch(numPedido, expRegular) = True Then
                        '        cadValida = True

                        '    Else
                        '        MessageBox.Show("Formato de Certificado Incorrecto, Por Favor Compruebelo", "Certificado Esterilización", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        '        cmdImprimirLotes.Enabled = True
                        '        Exit Sub
                        '    End If
                        'Loop
                        numPedido = txtJaponBatch.Text
                        If System.Text.RegularExpressions.Regex.IsMatch(numPedido, expRegular) = False Then
                            MessageBox.Show("Formato de Certificado Incorrecto, Por Favor Compruebelo", "Certificado Esterilización", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Sub
                        End If
                    Else
                        numPedido = txtNumPedido.Text
                    End If

                    If txtRefProducto.Text.Trim = "992 4001-JPN" Then
                        label = objbt.Formats.Open(ruta & "DML1028.btw")
                    Else
                        label = objbt.Formats.Open(ruta & "DML1039.btw")
                    End If

                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + empezarXlote) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")

                        If txtRefProducto.Text.Trim = "992 4001-JPN" Then
                            printTipo28()
                        Else
                            printReport39()
                        End If

                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.05)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "Tipo44", "Tipo45"
                    fecCaducidad = "EXP.:  20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")

                    Dim nombreEtiqueta As String

                    'Si el pedido es de Kyoritsu cambiamos el formato de la fecha de caducidad y ponemos el numero de esterilización.
                    'If txtRefProducto.Text.Trim = "996 0925-JPN" Or txtRefProducto.Text.Trim = "992 4001-JPN" Then
                    '    Dim expRegular, expAño As String
                    '    Dim cadValida As Boolean = False
                    '    expAño = "(" & txtAño.Value & "|" & (txtAño.Value - 1) & ")"
                    '    expRegular = "^(" & expAño & "(0[1-9]|1[012])([0-2][0-9]|3[01]))/(\d\d[A-D])$"
                    '    fecCaducidad = "EXP   20" & numAñoCad.Value & "." & numMesCad.Value.ToString("00")
                    '    'fecCaducidad = "" & numMesCad.Value.ToString("00") & "/20" & numAñoCad.Value.ToString("00")
                    '    'Do While cadValida = False
                    '    '    numPedido = txtJaponBatch.Text
                    '    '    'numPedido = InputBox("Este Pedido se tiene que Introducir el certificado de Esterilización", "Certificado Esterilizacion").ToUpper
                    '    '    If System.Text.RegularExpressions.Regex.IsMatch(numPedido, expRegular) = True Then
                    '    '        cadValida = True

                    '    '    Else
                    '    '        MessageBox.Show("Formato de Certificado Incorrecto, Por Favor Compruebelo", "Certificado Esterilización", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    '    '        cmdImprimirLotes.Enabled = True
                    '    '        Exit Sub
                    '    '    End If
                    '    'Loop
                    '    numPedido = txtJaponBatch.Text
                    '    If System.Text.RegularExpressions.Regex.IsMatch(numPedido, expRegular) = False Then
                    '        MessageBox.Show("Formato de Certificado Incorrecto, Por Favor Compruebelo", "Certificado Esterilización", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    '        Exit Sub
                    '    End If
                    'End If

                    'Creamos el objeto etiqueta
                    If ListBox1.SelectedItem = "Tipo45" Then
                        label = objbt.Formats.Open(ruta & "DML1045.btw")
                    Else
                        label = objbt.Formats.Open(ruta & "DML1044.btw")
                    End If

                    fecCaducidad = "EXP   20" & numAñoCad.Value & "." & numMesCad.Value.ToString("00")
                    fecCaducidad = "" & numMesCad.Value.ToString("00") & "/20" & numAñoCad.Value.ToString("00")

                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + empezarXlote) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")

                        printLote44()

                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.05)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "TVAL", "tipo16", "tipo18", "tipo25", "Tipo 25 (DML1025)", "Tipo 42"
                    'ETIQUETAS ESPECIFICAS DE TVA
                    Dim modeloEti As String
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    If ListBox1.SelectedItem = "tipo18" Then
                        modeloEti = "DML1018.btw"
                    ElseIf ListBox1.SelectedItem = "tipo16" Then
                        modeloEti = "DML1016.btw"
                    ElseIf ListBox1.SelectedItem = "Tipo 42" Then
                        modeloEti = "DML1042.btw"
                    Else
                        modeloEti = "DML1025.btw"
                    End If
                    'Creamos el objeto etiqueta
                    label = objbt.Formats.Open(ruta & modeloEti)
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + l) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteTVA()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "TVALO", "FL1100", "FL1103", "tipo19", "tipo24", "TIPO 41", "Tipo 46", "Tipo 47"
                    'Etiquetas especificas TVA Lote Ordenado.
                    If ListBox1.SelectedItem = "Tipo 46" Then
                        fecCaducidad = "20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    Else
                        fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    End If
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    If etiManual = True Then calcularChipInicial(0) 'AndAlso empezarXlote <> 1
                    ' ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    If ListBox1.SelectedItem = "TVALO" Then
                        label = objbt.Formats.Open(ruta & "FL1099_DM Bag From-To.btw") '("C:\Users\Ezequiel\Documents\etiquetas_datamars\FL1099_DM Bag From-To.btw")
                    Else
                        If ListBox1.SelectedItem = "FL1100" Then label = objbt.Formats.Open(ruta & "FL1100.btw")
                        If ListBox1.SelectedItem = "FL1103" Then label = objbt.Formats.Open(ruta & "FL1103_DM_ft.btw")
                        If ListBox1.SelectedItem = "tipo19" Then label = objbt.Formats.Open(ruta & "DML1019.btw")
                        If ListBox1.SelectedItem = "tipo24" Then label = objbt.Formats.Open(ruta & "DML1024.btw")
                        If ListBox1.SelectedItem = "TIPO 41" Then label = objbt.Formats.Open(ruta & "DML1041.btw")
                        If ListBox1.SelectedItem = "Tipo 46" Then label = objbt.Formats.Open(ruta & "DML1046.btw")
                        If ListBox1.SelectedItem = "Tipo 47" Then label = objbt.Formats.Open(ruta & "DML1047.btw")
                    End If
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + empezarXlote) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteTVAO()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "Tipo 43"
                    'Etiquetas especificas TVA Lote Ordenado.
                    'fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    fecCaducidad = numMesCad.Value.ToString("00") & "/" & "20" & numAñoCad.Value
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")

                    If etiManual = True Then calcularChipInicial(0) 'AndAlso empezarXlote <> 1
                    ' ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    label = objbt.Formats.Open(ruta & "DML1043.btw")

                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + empezarXlote) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        printLoteTipo43()
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "tipo32", "tipo31", "tipo29", "tipo30", "tipo33"
                    'NUEVAS ETIQUETAS ITALIA QIR-5510
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "LOTE", "ruta_lote")
                    If etiManual = True Then calcularChipInicial(0) 'AndAlso empezarXlote <> 1
                    If ListBox1.SelectedItem = "tipo32" Then label = objbt.Formats.Open(ruta & "DML1032_992 0000-IT2.btw")
                    If ListBox1.SelectedItem = "tipo31" Then label = objbt.Formats.Open(ruta & "DML1031_992 1000-IT2.btw")
                    If ListBox1.SelectedItem = "tipo29" Then label = objbt.Formats.Open(ruta & "DML1029_992 0000-IT3.btw")
                    If ListBox2.SelectedItem = "tipo30" Then label = objbt.Formats.Open(ruta & "DML1030_992 0000-ITA.btw")
                    If ListBox2.SelectedItem = "tipo33" Then label = objbt.Formats.Open(ruta & "DML1033.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text

                    For i = empezarXlote To (num_lotes + empezarXlote) - 1
                        numLote = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxcaja.ToString("0000") & "/" & loteinicial.ToString("00")
                        If ListBox2.SelectedItem = "tipo33" Then
                            printLoteTVAO(False)
                        Else
                            printLoteTVAO()
                        End If
                        loteinicial += 1
                        If loteinicial = (lotxCaja + 1) Then
                            loteinicial = 1
                            auxcaja += 1
                        End If
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
            End Select

            txtlotesimp.Text = 1
        Else
            MsgBox("No se han impreso las Etiquetas por un Error en la BBDD, Avise a Informatica.")
            histGrabado = False
        End If
        cmdImprimirLotes.Enabled = True

    End Sub

    Private Sub Imprimir_Cajas(ByVal num_lotes As Integer, ByVal num_cajas As Integer, ByVal muestra As Boolean)

        Dim grabBBDDok As Boolean
        Dim auxCaja As Integer = Integer.Parse(empezarXCaja) '(txtCaja.Value) se cambia el valor de la caja de texto por la variable
        firstCode = txtGlobalFirstCode.Text.Trim
        tempFirtCode = txtGlobalFirstCode.Text.Trim
        'Bloqueamos el boton para que no se pueda volver a imprimir.
        cmdImprimirCajas.Enabled = False
        If muestra Then
            grabBBDDok = True
        Else
            If histGrabado = False Then
                grabBBDDok = grabarHistorico("insertar")
                histGrabado = True
            Else
                grabBBDDok = grabarHistorico("actualizar")
            End If
        End If

        If grabBBDDok = True Then
            Select Case ListBox2.SelectedItem
                Case "Std"
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC_standart.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        printCajaStandart()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "tipo3"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC10003.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        printCajaTipo3()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "tipo4", "tipo7", "tipo6"
                    Dim respuesta As Integer
                    Dim modeloeti As String
                    If ListBox2.SelectedItem = "tipo6" Then
                        'segun el tipo de etiqueta abrimos un fichero u otro del Bartender
                        modeloeti = "DMC10006.btw"
                    Else
                        If ListBox2.SelectedItem = "tipo7" Then
                            modeloeti = "DMC10007.btw"
                        Else
                            modeloeti = "DMC10004.btw"
                        End If
                    End If
                    ' si reimprimimos
                    If reImprimir = True And etiManual = True Then
                        respuesta = MsgBox("Tiene 2 tipos de Etiquetas. Estas son las Etiquetas 80mmX80mm. ¿Necesitas Imprimirlas?", vbYesNo, "REIMPRESION DE ETIQUETAS")
                    Else
                        respuesta = 6

                    End If
                    If respuesta = 6 Then
                        MsgBox("Tiene 2 tipos de Etiquetas. Estas son las Etiquetas 80mmX80mm, Cambie las etiquetas en la impresora Referencia BOM:700 0500-011", MsgBoxStyle.Information)
                        fecCaducidad = "20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                        ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                        label = objbt.Formats.Open(ruta & modeloeti)
                        etImpConf = label.PrintSetup
                        etImpConf.Printer = cListImp.Text
                        'Calculamos la fecha de embolsado con la fecha de fabricación.
                        If fecPedFab.Contains("/") Then
                            fecsealing = Replace(fecPedFab, "/", "-")
                        Else
                            fecsealing = Strings.Right(fecPedFab, 2) & "-" & Mid(fecPedFab, 5, 2) & "-" & Strings.Left(fecPedFab, 4)
                        End If

                        For i As Integer = 0 To (num_cajas) - 1
                            numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "9" & auxCaja.ToString("0000")
                            printCajaTipo4()
                            auxCaja += 1
                            'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                            espera(0.2)
                        Next
                        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    End If
                    If reImprimir = True And etiManual = True Then
                        respuesta = MsgBox("Se van a imprimir las segundas Etiquetas de este Pedido de 60mmX47mm.¿Necesita Reimprimirlas?", vbYesNo, "REIMPRESION DE ETIQUETAS")
                    Else
                        respuesta = 6
                    End If
                    If respuesta = 6 Then

                        If ListBox2.SelectedItem = "tipo6" Then
                            mostrarEtiqCaja("Std")
                            MsgBox("!Se van a Imprimir las Etiquetas de 60x47mm¡, Cambie las etiquetas en la impresora Referencia BOM:600 0200-050", MsgBoxStyle.Information)
                            ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                            label = objbt.Formats.Open(ruta & "DMC_standart.btw")
                            etImpConf = label.PrintSetup
                            etImpConf.Printer = cListImp.Text
                            'ponemos el contador de cajas al valor de cajas iniciales, para que no sume. si no que imprima desde el valor inicial.
                            auxCaja = Integer.Parse(empezarXCaja) '(txtCaja.Value)
                            For i As Integer = 0 To (num_cajas) - 1
                                numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                                printCajaStandart()
                                auxCaja += 1
                                'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                                espera(0.2)
                            Next
                        Else

                            'se tiene que realizar el cambio de las etiquetas
                            mostrarEtiqCaja("tipo1")
                            MsgBox("!Se van a Imprimir las Etiquetas de 60x47mm¡, Cambie las etiquetas en la impresora Referencia BOM:600 0200-050", MsgBoxStyle.Information)
                            ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                            label = objbt.Formats.Open(ruta & "DMC10001.btw")
                            etImpConf = label.PrintSetup
                            etImpConf.Printer = cListImp.Text
                            fecCaducidad = "EXP.: 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                            'ponemos el contador de cajas al valor de cajas iniciales, para que no sume. si no que imprima desde el valor inicial.
                            auxCaja = Integer.Parse(empezarXCaja) '(txtCaja.Value)
                            For i As Integer = 0 To (num_cajas) - 1
                                numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                                printCajaTipo1()
                                auxCaja += 1
                                'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                                espera(0.2)
                            Next
                        End If

                        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    End If
                Case "tipo5"
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC10005.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        printCajaTipo5()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)

                Case "tipo9"
                    Dim respuesta As Integer
                    'Se imprimen Primero las etiquetas del Tipo 9
                    If reImprimir = True And etiManual = True Then
                        respuesta = MsgBox("Tiene 2 tipos de Etiquetas. Estas son las Etiquetas 40mmX50mm. ¿Necesitas Imprimirlas?", vbYesNo, "REIMPRESION DE ETIQUETAS")
                    Else
                        respuesta = 6

                    End If
                    If respuesta = 6 Then
                        MsgBox("Tiene 2 tipos de Etiquetas. Estas son las Etiquetas 40mmX50mm, Cambie las etiquetas en la impresora Referencia BOM:600 0200-053", MsgBoxStyle.Information)
                        fecCaducidad = "20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                        ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                        label = objbt.Formats.Open(ruta & "DMC10009.btw")
                        etImpConf = label.PrintSetup
                        etImpConf.Printer = cListImp.Text
                        For i As Integer = 0 To (num_cajas) - 1
                            numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                            printCajaTipo9()
                            auxCaja += 1
                            'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                            espera(0.2)
                        Next
                        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    End If

                    If reImprimir = True And etiManual = True Then
                        respuesta = MsgBox("Se van a imprimir las segundas Etiquetas de este Pedido de 60mmX47mm.¿Necesita Reimprimirlas?", vbYesNo, "REIMPRESION DE ETIQUETAS")
                    Else
                        respuesta = 6
                    End If
                    If respuesta = 6 Then
                        'se tiene que realizar el cambio de las etiquetas
                        mostrarEtiqCaja("Std")
                        MsgBox("!Se van a Imprimir las Etiquetas de 60x47mm¡, Cambie las etiquetas en la impresora Referencia BOM:600 0200-050", MsgBoxStyle.Information)
                        ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                        label = objbt.Formats.Open(ruta & "DMC_standart.btw")
                        etImpConf = label.PrintSetup
                        etImpConf.Printer = cListImp.Text
                        'ponemos el contador de cajas al valor de cajas iniciales, para que no sume. si no que imprima desde el valor inicial.
                        auxCaja = Integer.Parse(empezarXCaja) '(txtCaja.Value)
                        For i As Integer = 0 To (num_cajas) - 1
                            numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                            printCajaStandart()
                            auxCaja += 1
                            'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                            espera(0.2)
                        Next
                        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    End If
                Case "tipo10"
                    Dim respuesta As Integer
                    'Se imprimen Primero las etiquetas del Tipo 9
                    If reImprimir = True And etiManual = True Then
                        respuesta = MsgBox("Tiene 2 tipos de Etiquetas. Estas son las Etiquetas 40mmX50mm. ¿Necesitas Imprimirlas?", vbYesNo, "REIMPRESIÓN DE ETIQUETAS")
                    Else
                        respuesta = 6

                    End If
                    If respuesta = 6 Then
                        MsgBox("Tiene 2 tipos de Etiquetas. Estas son las Etiquetas 40mmX50mm, Cambie las etiquetas en la impresora Referencia: 600 0200-053", MsgBoxStyle.Information)
                        fecCaducidad = "20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                        ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                        label = objbt.Formats.Open(ruta & "DMC10010.btw")
                        etImpConf = label.PrintSetup
                        etImpConf.Printer = cListImp.Text
                        'refCliente = "5010605008023"
                        For i As Integer = 0 To (num_cajas) - 1
                            numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                            printCajaTipo10()
                            auxCaja += 1
                            'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                            espera(0.2)
                        Next
                        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    End If

                    If reImprimir = True And etiManual = True Then
                        respuesta = MsgBox("Se van a imprimir las segundas Etiquetas de este Pedido de 60mmX47mm.¿Necesita Reimprimirlas?", vbYesNo, "REIMPRESION DE ETIQUETAS")
                    Else
                        respuesta = 6
                    End If
                    If respuesta = 6 Then
                        'se tiene que realizar el cambio de las etiquetas
                        mostrarEtiqCaja("Std")
                        MsgBox("!Se van a Imprimir las Etiquetas de 60x47mm¡, Cambie las etiquetas en la impresora Referencia :600 0200-050", MsgBoxStyle.Information)
                        ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                        label = objbt.Formats.Open(ruta & "DMC_standart.btw")
                        etImpConf = label.PrintSetup
                        etImpConf.Printer = cListImp.Text
                        'ponemos el contador de cajas al valor de cajas iniciales, para que no sume. si no que imprima desde el valor inicial.
                        auxCaja = Integer.Parse(empezarXCaja) '(txtCaja.Value)
                        For i As Integer = 0 To (num_cajas) - 1
                            numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                            printCajaStandart()
                            auxCaja += 1
                            'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                            espera(0.2)
                        Next
                        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    End If
                Case "tipo11"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC10011.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        printCajaTipo11()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "tipo12", "tipo27"
                    Dim respuesta As Integer
                    'Se imprimen Primero las etiquetas del Tipo 9
                    If reImprimir = True And etiManual = True Then
                        respuesta = MsgBox("Tiene 2 tipos de Etiquetas. Estas son las Etiquetas 80mmX80mm. ¿Necesitas Imprimirlas?", vbYesNo, "REIMPRESION DE ETIQUETAS")
                    Else
                        respuesta = 6

                    End If
                    If respuesta = 6 Then
                        MsgBox("Tiene 2 tipos de Etiquetas. Estas son las Etiquetas 80mmX80mm, Cambie las etiquetas en la impresora Referencia BOM:600 0200-053", MsgBoxStyle.Information)
                        fecCaducidad = "20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                        ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                        label = objbt.Formats.Open(ruta & "DMC10012.btw")
                        etImpConf = label.PrintSetup
                        etImpConf.Printer = cListImp.Text
                        For i As Integer = 0 To (num_cajas) - 1
                            numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                            printCajaTipo12()
                            auxCaja += 1
                            'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                            espera(0.2)
                        Next
                        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    End If

                    If reImprimir = True And etiManual = True Then
                        respuesta = MsgBox("Se van a imprimir las segundas Etiquetas de este Pedido de 60mmX47mm.¿Necesita Reimprimirlas?", vbYesNo, "REIMPRESION DE ETIQUETAS")
                    Else
                        respuesta = 6
                    End If
                    If respuesta = 6 Then
                        'se tiene que realizar el cambio de las etiquetas

                        MsgBox("!Se van a Imprimir las Etiquetas de 60x47mm¡, Cambie las etiquetas en la impresora Referencia BOM:600 0200-050", MsgBoxStyle.Information)

                        If ListBox2.SelectedItem = "tipo12" Then
                            mostrarEtiqCaja("Std")
                            label = objbt.Formats.Open(ruta & "DMC_standart.btw")
                        ElseIf ListBox2.SelectedItem = "tipo27" Then
                            mostrarEtiqCaja("tipo27")
                            label = objbt.Formats.Open(ruta & "DMC10027.btw")
                        End If

                        ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                        label = objbt.Formats.Open(ruta & "DMC_standart.btw")
                        etImpConf = label.PrintSetup
                        etImpConf.Printer = cListImp.Text
                        'ponemos el contador de cajas al valor de cajas iniciales, para que no sume. si no que imprima desde el valor inicial.
                        auxCaja = Integer.Parse(empezarXCaja) '(txtCaja.Value)
                        For i As Integer = 0 To (num_cajas) - 1
                            numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                            If ListBox2.SelectedItem = "tipo12" Then
                                mostrarEtiqCaja("Std")
                                label = objbt.Formats.Open(ruta & "DMC_standart.btw")
                                printCajaStandart()
                            ElseIf ListBox2.SelectedItem = "tipo27" Then
                                mostrarEtiqCaja("tipo27")
                                label = objbt.Formats.Open(ruta & "DMC10027.btw")
                                printCajaTipo27()
                            End If
                            auxCaja += 1
                            'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                            espera(0.2)
                        Next
                        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    End If
                Case "tipo13"
                    fecCaducidad = "EXP. 20" & (numAñoCad.Value) & "-" & numMesCad.Value.ToString("00") 'quitar el 2 es temporal
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC10013.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        printCajaTipo13()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "tipo14", "tipo17", "tipo22", "tipo23"
                    fecCaducidad = "EXP. 20" & (numAñoCad.Value) & "-" & numMesCad.Value.ToString("00") 'quitar el 2 es temporal
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    If ListBox2.SelectedItem = "tipo14" Then
                        label = objbt.Formats.Open(ruta & "DMC10014.btw")
                    ElseIf ListBox2.SelectedItem = "tipo17" Then
                        label = objbt.Formats.Open(ruta & "DMC10017.btw")
                    ElseIf ListBox2.SelectedItem = "tipo22" Then
                        label = objbt.Formats.Open(ruta & "DMC10022.btw")
                    ElseIf ListBox2.SelectedItem = "tipo23" Then
                        label = objbt.Formats.Open(ruta & "DMC10023.btw")
                    End If
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        printcajaTipo14()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "Tipo48"
                    Dim respuesta As Integer
                    'Se imprimen Primero las etiquetas del Tipo 9
                    If reImprimir = True And etiManual = True Then
                        respuesta = MsgBox("Tiene 2 tipos de Etiquetas. Estas son las Etiquetas 80mmX80mm. ¿Necesitas Imprimirlas?", vbYesNo, "REIMPRESION DE ETIQUETAS")
                    Else
                        respuesta = 6

                    End If
                    If respuesta = 6 Then
                        MsgBox("Tiene 2 tipos de Etiquetas. Estas son las Etiquetas 80mmX80mm, Cambie las etiquetas en la impresora Referencia BOM:600 0200-053", MsgBoxStyle.Information)
                        fecCaducidad = "EXP. 20" & (numAñoCad.Value) & "-" & numMesCad.Value.ToString("00") 'quitar el 2 es temporal
                        ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                        label = objbt.Formats.Open(ruta & "DMC10022.btw")
                        etImpConf = label.PrintSetup
                        etImpConf.Printer = cListImp.Text
                        For i As Integer = 0 To (num_cajas) - 1
                            numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                            printCajaTipo14()
                            auxCaja += 1
                            'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                            espera(0.2)
                        Next
                        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    End If

                    If reImprimir = True And etiManual = True Then
                        respuesta = MsgBox("Se van a imprimir las segundas Etiquetas de este Pedido de 60mmX47mm.¿Necesita Reimprimirlas?", vbYesNo, "REIMPRESION DE ETIQUETAS")
                    Else
                        respuesta = 6
                    End If
                    If respuesta = 6 Then
                        'se tiene que realizar el cambio de las etiquetas
                        MsgBox("!Se van a Imprimir las Etiquetas de 60x47mm¡, Cambie las etiquetas en la impresora Referencia BOM:600 0200-050", MsgBoxStyle.Information)
                        mostrarEtiqCaja("Std")
                        ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                        label = objbt.Formats.Open(ruta & "DMC_standart.btw")
                        etImpConf = label.PrintSetup
                        etImpConf.Printer = cListImp.Text
                        'ponemos el contador de cajas al valor de cajas iniciales, para que no sume. si no que imprima desde el valor inicial.
                        auxCaja = Integer.Parse(empezarXCaja) '(txtCaja.Value)
                        For i As Integer = 0 To (num_cajas) - 1
                            numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                            mostrarEtiqCaja("Std")
                            label = objbt.Formats.Open(ruta & "DMC_standart.btw")
                            printCajaStandart()
                            auxCaja += 1
                            'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                            espera(0.2)
                        Next
                        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    End If

                Case "tipo15"
                    MsgBox("Este cliente Se imprimen 2 Etiquetas de Diferentes Tamaños. Las Primeras de 60x47mm y Las segundas de 80x80mm_
                        _ Asegurese de tener las etiquetas correctas", MsgBoxStyle.Information)
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC10015.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        printCajaTipo15()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                    mostrarEtiqCaja("tipo12")
                    MsgBox("!Se van ha Imprimir las Etiquetas de 80x80mm¡, Cambie el tamaño de las etiquetas", MsgBoxStyle.Information)

                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC10012.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    'Inicializamos el valor de la variable de caja para que se impriman las mismas cajas.
                    auxCaja = Integer.Parse(empezarXCaja) '(txtCaja.Value)
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        printCajaTipo12()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "report36", "report42", "report45"
                    Dim numLoteTmp As Integer
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    'si la etiqueta es manual y numero de caja no es 1 se tiene que calcula el chip inicial de caja
                    If etiManual = True Then calcularChipInicial(1)
                    'asignamos la cantidad de lotes a una variable temporal.
                    numLoteTmp = num_lotes
                    'lotxCaja = numLotesPorCaja.Value
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC10036.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        If (jerxLote * lotxCaja) > (jerxLote * ((numLoteTmp + empezarXlote) - 1)) Then lotxCaja = ((numLoteTmp + empezarXlote) - 1)
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
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "report38"
                    Dim numLoteTmp As Integer
                    Dim contador As Integer

                    'si la etiqueta es manual y numero de caja no es 1 se tiene que calcula el chip inicial de caja
                    If etiManual = True Then calcularChipInicial(1) 'AndAlso empezarXCaja <> 1
                    'asignamos la cantidad de lotes a una variable temporal.
                    numLoteTmp = num_lotes
                    If fecPedFab.Contains("/") Then
                        fecLotImp = Strings.Right(fecPedFab, 4) & "-" & Mid(fecPedFab, 4, 2) & "-" & Strings.Left(fecPedFab, 2)
                    Else
                        fecLotImp = Strings.Left(fecPedFab, 4) & "-" & Mid(fecPedFab, 5, 2) & "-" & Strings.Right(fecPedFab, 2)
                    End If

                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC10038.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    'Para sacar el numero de caja actual de la cantidad total restamos la caja inicial por la variable empezarXcaja.
                    'Inicializamos el contador fuera del bucle
                    contador = (empezarXCaja - txtCaja.Value) + 1
                    For i As Integer = 0 To (num_cajas) - 1

                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        If (jerxLote * lotxCaja) > (jerxLote * ((numLoteTmp + empezarXlote) - 1)) Then lotxCaja = ((numLoteTmp + empezarXlote) - 1)
                        'restamos a la variable de lotes temporales la cantidad de lotes por caja. Esto se hace por si la ultima caja hay menos lotes que los reglamentarios salga bien la numeración.
                        numLoteTmp = numLoteTmp - lotxCaja
                        printCajaReport38(contador)
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                        contador += 1
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                'lotxCaja = numLotesPorCaja.Value
                Case "report40", "report44", "Tipo28"
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    If txtRefProducto.Text.Trim = "992 4001-JPN" Then
                        label = objbt.Formats.Open(ruta & "DMC10028.btw")
                    Else
                        label = objbt.Formats.Open(ruta & "DMC10040.btw")
                    End If
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    'Si el pedido es de Kyoritsu cambiamos el formato de la fecha de caducidad y ponemos el numero de esterilización.
                    If txtRefProducto.Text.Trim = "996 0925-JPN" Or txtRefProducto.Text.Trim = "992 4001-JPN" Then
                        Dim expRegular, expAño As String
                        Dim cadValida As Boolean = False
                        expAño = "(" & txtAño.Value & "|" & (txtAño.Value - 1) & ")"
                        expRegular = "^(" & expAño & "(0[1-9]|1[012])([0-2][0-9]|3[01]))/(\d\d[A-D])$"
                        fecCaducidad = "EXP   20" & numAñoCad.Value & "." & numMesCad.Value.ToString("00")
                        'Do While cadValida = False
                        '    'numPedido = InputBox("Este Pedido se tiene que Introducir el certificado de Esterilización", "Certificado Esterilizacion").ToUpper
                        '    numPedido = txtJaponBatch.Text
                        '    If System.Text.RegularExpressions.Regex.IsMatch(numPedido, expRegular) = True Then
                        '        cadValida = True
                        '    Else
                        '        MessageBox.Show("Formato de Certificado Incorrecto, Por Favor Compruebelo", "Certificado Esterilización", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        '    End If
                        'Loop

                        numPedido = txtJaponBatch.Text
                        If System.Text.RegularExpressions.Regex.IsMatch(numPedido, expRegular) = False Then
                            MessageBox.Show("Formato de Certificado Incorrecto, Por Favor Compruebelo", "Certificado Esterilización", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Sub
                        End If
                    Else
                        numPedido = txtNumPedido.Text
                    End If
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        If txtRefProducto.Text.Trim = "992 4001-JPN" Then
                            Dim cantidad As Integer = numLotesPorCaja.Value * numUnidadesPorlote.Value
                            printCajaTipo28(cantidad)
                        Else
                            printCajaReport40()
                        End If
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)


                Case "Tipo44", "Tipo45"

                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    If ListBox2.SelectedItem = "Tipo45" Then
                        label = objbt.Formats.Open(ruta & "DMC10045.btw")
                        fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    Else
                        label = objbt.Formats.Open(ruta & "DMC10044.btw")
                        fecCaducidad = numMesCad.Value.ToString("00") & "/20" & numAñoCad.Value
                    End If

                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    'Si el pedido es de Kyoritsu cambiamos el formato de la fecha de caducidad y ponemos el numero de esterilización.


                    'fecCaducidad = numMesCad.Value.ToString("00") & "/20" & numAñoCad.Value

                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        If ListBox2.SelectedItem = "Tipo45" Then
                            printCaja45()
                        Else
                            printCaja44()
                        End If
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)

                Case "Tipoar"
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC100AR2.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    Dim contador As Integer = 1
                    Dim textoEti As String
                    Select Case refArticulo
                        Case "996 0000-AR2"
                            textoEti = "STANDARD IN BOX OF 10 STUD"
                        Case "996 0000-981"
                            textoEti = "BOX OF TEN"

                    End Select
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        printCajaAr(contador, txtnumsap.Text.Trim, textoEti, num_cajas)
                        contador += 1
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "TVAC", "tipo16", "tipo17", "tipo18", "tipo25", "Tipo 25 (DMC10025)", "Tipo 42"
                    Dim modeloeti As String
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    If ListBox2.SelectedItem = "tipo17" Then
                        'segun el tipo de etiqueta abrimos un fichero u otro del Bartender
                        modeloeti = "DMC10017.btw"
                    ElseIf ListBox2.SelectedItem = "tipo18" Then
                        modeloeti = "DMC10018.btw"
                    ElseIf ListBox2.SelectedItem = "tipo16" Then
                        modeloeti = "DMC10016.btw"
                    ElseIf ListBox2.SelectedItem = "Tipo 42" Then
                        modeloeti = "DMC10042.btw"
                    Else
                        modeloeti = "DMC10025.btw"
                    End If
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & modeloeti)
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    'label = objbt.Formats.Open("C:\etiquetas\FC10088_CajaDM.btw")
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        printCajaTVA()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
                Case "TVACO", "FC10092", "FC10093", "tipo19", "tipo21", "tipo24", "TIPO 41"
                    Dim numLoteTmp As Integer
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    'ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    If ListBox2.SelectedItem = "TVACO" Then
                        label = objbt.Formats.Open(ruta & "FC10091_DM Box From-To.btw") '("C:\Users\Ezequiel\Documents\etiquetas_datamars\FC10091_DM Box From-To.btw")
                    Else
                        If ListBox2.SelectedItem = "FC10092" Then label = objbt.Formats.Open(ruta & "FC10092.btw")
                        If ListBox2.SelectedItem = "FC10093" Then label = objbt.Formats.Open(ruta & "FC10093_DM_ft.btw")
                        If ListBox2.SelectedItem = "tipo19" Then label = objbt.Formats.Open(ruta & "DMC10019.btw")
                        If ListBox2.SelectedItem = "tipo21" Then label = objbt.Formats.Open(ruta & "DMC10021.btw")
                        If ListBox2.SelectedItem = "tipo24" Then label = objbt.Formats.Open(ruta & "DMC10024.btw")
                        If ListBox2.SelectedItem = "TIPO 41" Then label = objbt.Formats.Open(ruta & "DMC10041.btw")
                    End If
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    'si la etiqueta es manual y numero de caja no es 1 se tiene que calcula el chip inicial de caja
                    If etiManual = True Then calcularChipInicial(1)
                    'asignamos la cantidad de lotes a una variable temporal.
                    numLoteTmp = num_lotes
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        If (jerxLote * lotxCaja) > (jerxLote * ((numLoteTmp + empezarXlote) - 1)) Then lotxCaja = ((numLoteTmp + empezarXlote) - 1)
                        'restamos a la variable de lotes temporales la cantidad de lotes por caja. Esto se hace por si la ultima caja hay menos lotes que los reglamentarios.
                        numLoteTmp = numLoteTmp - lotxCaja
                        printCajaTVAO()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)


                Case "Tipo 43"
                    Dim numLoteTmp As Integer
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    fecCaducidad = numMesCad.Value.ToString("00") & "/20" & numAñoCad.Value
                    'ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    label = objbt.Formats.Open(ruta & "DMC10043.btw")

                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    'si la etiqueta es manual y numero de caja no es 1 se tiene que calcula el chip inicial de caja
                    If etiManual = True Then calcularChipInicial(1)
                    'asignamos la cantidad de lotes a una variable temporal.
                    numLoteTmp = num_lotes
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        If (jerxLote * lotxCaja) > (jerxLote * ((numLoteTmp + empezarXlote) - 1)) Then lotxCaja = ((numLoteTmp + empezarXlote) - 1)
                        'restamos a la variable de lotes temporales la cantidad de lotes por caja. Esto se hace por si la ultima caja hay menos lotes que los reglamentarios.
                        numLoteTmp = numLoteTmp - lotxCaja
                        printCajaTipo43()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)

                    'NUEVAS ETIQUETAS ITALIA QIR-5510
                Case "tipo32", "tipo31", "tipo29", "tipo30", "tipo33"
                    Dim numLoteTmp As Integer
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    If ListBox2.SelectedItem = "tipo32" Then label = objbt.Formats.Open(ruta & "DMC10032_992 0000-IT2.btw")
                    If ListBox2.SelectedItem = "tipo31" Then label = objbt.Formats.Open(ruta & "DMC10031_992 1000-IT2.btw")
                    If ListBox2.SelectedItem = "tipo29" Then label = objbt.Formats.Open(ruta & "DMC10029_992 0000-IT3.btw")
                    If ListBox2.SelectedItem = "tipo30" Then label = objbt.Formats.Open(ruta & "DMC10030_992 0000-ITA.btw")
                    If ListBox2.SelectedItem = "tipo33" Then label = objbt.Formats.Open(ruta & "DMC10033.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    'si la etiqueta es manual y numero de caja no es 1 se tiene que calcula el chip inicial de caja
                    If etiManual = True Then calcularChipInicial(1)
                    'asignamos la cantidad de lotes a una variable temporal.
                    numLoteTmp = num_lotes
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        If (jerxLote * lotxCaja) > (jerxLote * ((numLoteTmp + empezarXlote) - 1)) Then lotxCaja = ((numLoteTmp + empezarXlote) - 1)
                        'restamos a la variable de lotes temporales la cantidad de lotes por caja. Esto se hace por si la ultima caja hay menos lotes que los reglamentarios.
                        numLoteTmp = numLoteTmp - lotxCaja
                        If ListBox2.SelectedItem = "tipo33" Then
                            printCajaTVAO(False)
                        Else
                            printCajaTVAO()
                        End If
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)


                Case "tipo20"
                    Dim numLoteTmp As Integer
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    label = objbt.Formats.Open(ruta & "DMC10020.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    'si la etiqueta es manual y numero de caja no es 1 se tiene que calcula el chip inicial de caja
                    If etiManual = True Then calcularChipInicial(1)
                    'asignamos la cantidad de lotes a una variable temporal.
                    numLoteTmp = num_lotes
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        If (jerxLote * lotxCaja) > (jerxLote * ((numLoteTmp + empezarXlote) - 1)) Then lotxCaja = ((numLoteTmp + empezarXlote) - 1)
                        'restamos a la variable de lotes temporales la cantidad de lotes por caja. Esto se hace por si la ultima caja hay menos lotes que los reglamentarios.
                        numLoteTmp = numLoteTmp - lotxCaja
                        printCaja20()
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)

                Case "Tipo 46", "Tipo 47"
                    Dim numLoteTmp As Integer
                    fecCaducidad = "20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
                    'si la etiqueta es manual y numero de caja no es 1 se tiene que calcula el chip inicial de caja
                    If etiManual = True Then calcularChipInicial(1)
                    'asignamos la cantidad de lotes a una variable temporal.
                    numLoteTmp = num_lotes
                    'lotxCaja = numLotesPorCaja.Value
                    ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "CAJA", "ruta_caja")
                    If ListBox2.SelectedItem = "Tipo 46" Then label = objbt.Formats.Open(ruta & "DMC10046.btw")
                    If ListBox2.SelectedItem = "Tipo 47" Then label = objbt.Formats.Open(ruta & "DMC10047.btw")
                    etImpConf = label.PrintSetup
                    etImpConf.Printer = cListImp.Text
                    For i As Integer = 0 To (num_cajas) - 1
                        numCaja = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & auxCaja.ToString("0000")
                        If (jerxLote * lotxCaja) > (jerxLote * ((numLoteTmp + empezarXlote) - 1)) Then lotxCaja = ((numLoteTmp + empezarXlote) - 1)
                        'restamos a la variable de lotes temporales la cantidad de lotes por caja. Esto se hace por si la ultima caja hay menos lotes que los reglamentarios.
                        numLoteTmp = numLoteTmp - lotxCaja
                        printCajaReport36(3)
                        auxCaja += 1
                        'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                        espera(0.2)
                    Next
                    label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
            End Select
            txtCajasImp.Text = 1
        Else
            MsgBox("No se han impreso las Etiquetas por un Error en la BBDD, Avise a Informatica.")
            histGrabado = False
        End If

        'Desbloqueamos el boton
        cmdImprimirCajas.Enabled = True

    End Sub

    Private Sub Imprimir_Palets_Nuevo(ByVal num_palets As Integer, ByVal muestra As Boolean)

        Dim grabBBDDok As Boolean
        If muestra Then
            grabBBDDok = True
        Else
            If histGrabado = False Then
                grabBBDDok = grabarHistorico("insertar")
                histGrabado = True
            Else
                grabBBDDok = grabarHistorico("actualizar")
            End If
        End If
        If grabBBDDok Then
            ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "PALET", "ruta_palet")
            label = objbt.Formats.Open(ruta & "PEGATINAS_PALLETS.btw")
            etImpConf = label.PrintSetup
            etImpConf.Printer = cListImp.Text
            For i = 1 To num_palets
                printEtiPalet_Nueva(txtNumPedido.Text, i, txtNombreCliente.Text)
            Next
            label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
        Else
            MsgBox("No se han impreso las Etiquetas por un Error en la BBDD, Avise a Informatica.")
            histGrabado = False
        End If

    End Sub

    Private Sub Imprimir_Palets(ByVal num_cajas As Integer, ByVal muestra As Boolean)

        Dim nbox, nboxfin, nboxtemp, paletinicial, i As Integer
        Dim aux_cajas As Integer = 0
        Dim cajas_palet As Integer = 1
        Dim numcajaini, numcajafin, caja_actual As String
        Dim grabBBDDok As Boolean
        Dim qr As Boolean = False
        Dim auxFin, auxInicio As Integer

        If muestra Then
            grabBBDDok = True
        Else
            If histGrabado = False Then
                grabBBDDok = grabarHistorico("insertar")
                histGrabado = True
            Else
                grabBBDDok = grabarHistorico("actualizar")
            End If
        End If

        If grabBBDDok = True Then
            fecCaducidad = "EXP. 20" & numAñoCad.Value & "-" & numMesCad.Value.ToString("00")
            paletinicial = Integer.Parse(txtpalet.Value)
            nbox = Integer.Parse(empezarXCaja) '(txtCaja.Value)
            nboxtemp = num_cajas
            ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "PALET", "ruta_palet")

            If txtNombreCliente.Text = "TRU TEST BRASIL" Or txtNombreCliente.Text = "DATAMARS BRASIL TECNOLOGIA AGROPECUARIA LTDA" Then
                label = objbt.Formats.Open(ruta & "etiqueta_palet_brasil.btw")
                qr = True
            Else
                label = objbt.Formats.Open(ruta & "etiqueta_palet.btw")
            End If

            For i = 1 To numPalets.Value
                Dim qrTmp As String = ""
                numcajaini = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & nbox.ToString("0000")
                If nboxtemp < numCajasPorPalet.Value Then
                    nboxfin = nbox + (nboxtemp - 1)
                Else
                    nboxfin = (nbox + numCajasPorPalet.Value) - 1
                    nboxtemp -= numCajasPorPalet.Value
                End If
                numcajafin = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & nboxfin.ToString("0000")

                If qr Then
                    auxFin = Integer.Parse(numcajafin.Substring(numcajafin.Length - 5))
                    auxInicio = Integer.Parse(numcajaini.Substring(numcajaini.Length - 5))
                    For j = auxInicio To auxFin
                        If cajas_palet < 51 Then
                            qrTmp = qrTmp & Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_" & auxInicio + aux_cajas & "-981" & vbCrLf
                            aux_cajas += 1
                            cajas_palet += 1
                        Else
                            cajas_palet = 1
                        End If

                    Next
                    cajas_palet = 1
                    printEtiPaletBrasil(txtnumsap.Text, i, numcajaini, numcajafin, qrTmp)
                Else

                    printEtiPalet(txtnumsap.Text, paletinicial, numcajaini, numcajafin)
                    paletinicial += 1
                    nbox = nboxfin + 1
                    'introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                    espera(0.2)
                End If
                'numcajafin = Integer.Parse(txtSemana.Text).ToString("00") & txtAño.Text & "_9" & nboxfin.ToString("0000")
                'printEtiPalet(txtnumsap.Text, paletinicial, numcajaini, numcajafin)
                'paletinicial += 1
                'nbox = nboxfin + 1
                ''introducimos una espera para que la impresora reciba la etiqueta y no se alternen.
                'espera(0.2)
            Next
            qr = False
            label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
            txtPaletImp.Text = 1

            'al terminar ponemos los valores a cero
            nbox = 0
            nboxfin = 0
            nboxtemp = 0
        Else
            MsgBox("No se han impreso las Etiquetas por un Error en la BBDD, Avise a Informatica.")
            histGrabado = False
        End If

    End Sub

    Private Sub cmdImprimirLotes_Click(sender As Object, e As EventArgs) Handles cmdImprimirLotes.Click

        Imprimir_Lotes(numLotes.Value, False)

    End Sub

    Private Sub cmdImprimirCajas_Click(sender As Object, e As EventArgs) Handles cmdImprimirCajas.Click

        Imprimir_Cajas(numLotes.Value, numCajas.Value, False)

    End Sub
    Private Sub cmdImprimirPalets_Click(sender As Object, e As EventArgs) Handles cmdImprimirPalets.Click

        'Imprimir_Palets(numCajas.Value, False)
        Imprimir_Palets_Nuevo(numPalets.Value, False)

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
        'asignamos el valor del chip inicial a las variables de Caja 
        firstCode = txtGlobalFirstCode.Text
        tempFirtCode = txtGlobalFirstCode.Text
        'asignamos el valor del chip inicial a las variables de lote
        FirstCodeLot = txtGlobalFirstCode.Text
        tempFirstCodeLot = txtGlobalFirstCode.Text
    End Sub

    Private Sub txtLastCodigoChip_TextChanged(sender As Object, e As EventArgs) Handles txtLastCodigoChip.TextChanged
        'asignamos el valor del chip final a las variables de Caja
        lastCode = txtLastCodigoChip.Text
        tempLastCode = txtLastCodigoChip.Text
        'asignamos el valor del chip final a las variables de Lote
        LastCodeLot = txtLastCodigoChip.Text
        tempLastCodeLot = txtLastCodigoChip.Text
    End Sub

    Private Sub txtNumPedido_TextChanged(sender As Object, e As EventArgs) Handles txtNumPedido.TextChanged
        'numPedido = txtNumPedido.Text
        dt1 = uspRangoChip.GetData(Integer.Parse(txtCodPedido.Text.Trim))
        rellenarCamposRango()
    End Sub

    Private Sub numUnidadesPorlote_ValueChanged(sender As Object, e As EventArgs) Handles numUnidadesPorlote.ValueChanged
        jerxLote = numUnidadesPorlote.Value
    End Sub
    Private Sub numLotesPorCaja_ValueChanged(sender As Object, e As EventArgs) Handles numLotesPorCaja.ValueChanged
        lotxCaja = numLotesPorCaja.Value
    End Sub
#End Region


    Private Sub txtCodPedido_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCodPedido.KeyDown
        'si pulsamos la tecla enter o introducimos un salto de linea ejecutamos el codigo.
        If e.KeyCode = Keys.Enter Or e.KeyValue = Asc(13) Then
            'Realizamos consulta con el WorkOrder para sacar la fecha de produccion y de caducidad
            Dim pedidoExiste As Boolean = False
            Dim consulta As New DataSet1TableAdapters.qryInsertHistorico
            'reseteamos los maximos y minimos.
            resetDsdHst()
            pedidoExiste = cargarPedido()
            If pedidoExiste = True Then
                cargaFechasPedido(txtCodPedido.Text.Trim)
            End If

            'Try
            '    numAñoCad.Value = consulta.GetAñosCaducidad(txtRefProducto.Text.Trim)
            'Catch ex As Exception

            'End Try

            If txtRefProducto.Text.Trim = "996 0925-JPN" Or txtRefProducto.Text.Trim = "992 4001-JPN" Then
                labelJapon.Visible = True
                txtJaponBatch.Visible = True
            Else
                labelJapon.Visible = False
                txtJaponBatch.Visible = False
            End If

            MessageBox.Show("Antes de Imprimir las etiquetas Compruebe todos los datos de la izquierda son correctos." & vbCrLf &
                        "Si no lo son Avise a Informatica", "COMPRUEBE LOS DATOS A IMPRIMIR", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub numLotes_ValueChanged(sender As Object, e As EventArgs) Handles numLotes.ValueChanged
        cantLote = numLotes.Value
    End Sub

    Private Sub numCajas_ValueChanged(sender As Object, e As EventArgs) Handles numCajas.ValueChanged
        cantCaja = numCajas.Value
    End Sub

    Private Sub txtLote_ValueChanged(sender As Object, e As EventArgs) Handles txtLote.ValueChanged
        empezarXlote = txtLote.Value
        'If numLotesPorCaja.Value <> 0 Then
        '    If txtLote.Value > numLotesPorCaja.Value Then
        '        MsgBox("El Numero de Lote Inicial no puede ser Mayor Que los lotes por Caja", MsgBoxStyle.Information)
        '        txtLote.Value = numLotesPorCaja.Value
        '    Else
        '        Dim st As New System.Diagnostics.StackTrace()
        '        Dim nombreMetodo As String
        '        nombreMetodo = st.GetFrame(3).GetMethod().Name
        '        Select Case nombreMetodo
        '            Case "UpButton"
        '                numLotes.Value -= 1
        '            Case "DownButton"
        '                numLotes.Value += 1
        '            Case "ParseEditText"
        '                Dim lotes As Integer
        '                lotes = (totalLotes - (txtCaja.Value * numLotesPorCaja.Value)) + ((numLotesPorCaja.Value) - txtLote.Value) + 1
        '                numLotes.Value = lotes
        '        End Select
        '    End If

        'End If
    End Sub

    Private Sub txtSemana_ValueChanged(sender As Object, e As EventArgs) Handles txtSemana.ValueChanged
        CajaInicial = ultimaCaja(txtSemana.Value, txtAño.Value)
        txtCaja.Value = CajaInicial
    End Sub

    Private Sub txtCaja_ValueChanged(sender As Object, e As EventArgs) Handles txtCaja.ValueChanged
        empezarXCaja = txtCaja.Value
        'Dim st As New System.Diagnostics.StackTrace()
        'Dim nombreMetodo As String
        'nombreMetodo = st.GetFrame(3).GetMethod().Name
        'Try
        '    Select Case nombreMetodo
        '        Case "UpButton"
        '            'si pulsamos la flecha UP de la caja numerica para sumar 
        '            numLotes.Value = (numLotes.Value - numLotesPorCaja.Value)
        '            numCajas.Value -= 1
        '            If (txtCaja.Value + cantCaja) < TotalCajas Then 'Or txtCaja.Value > (CajaInicial + numCajas.Value)
        '                MsgBox("El Numero de Caja Inicial No puede ser Mayor al Total de Cajas a Imprimir", MsgBoxStyle.Information)
        '                numLotes.Value = (numLotes.Value + numLotesPorCaja.Value)
        '                numCajas.Value += 1
        '                txtCaja.Value -= 1
        '            End If
        '        Case "DownButton"
        '            'si pulsamos la flecha Down de la caja numerica para restar
        '            numLotes.Value = (numLotes.Value + numLotesPorCaja.Value)
        '            numCajas.Value += 1
        '            If txtCaja.Value < CajaInicial Then
        '                MsgBox("El numero de Caja No puede ser Inferior al Numero de Caja Inicial", MsgBoxStyle.Information)
        '                numLotes.Value = (numLotes.Value - numLotesPorCaja.Value)
        '                numCajas.Value -= 1
        '                txtCaja.Value += 1
        '            End If
        '        Case "ParseEditText"
        '            'si ponemos el numero  a mano.
        '            If txtCaja.Value < CajaInicial Then

        '                MsgBox("El Numero de Caja no puede ser Menor al Numero de Caja Inicial", MsgBoxStyle.Information)
        '                txtCaja.Value = CajaInicial
        '            Else
        '                If txtCaja.Value > (TotalCajas - 1) Then
        '                    MsgBox("El Numero de Caja Inicial No puede ser Mayor al Total de Cajas a Imprimir", MsgBoxStyle.Information)
        '                    txtCaja.Value = CajaInicial
        '                Else
        '                    Dim ncajImp As Integer
        '                    ncajImp = (TotalCajas - txtCaja.Value)
        '                    numCajas.Value = ncajImp
        '                    numLotes.Value = (totalLotes - (numLotesPorCaja.Value * (txtCaja.Value - 1)))
        '                End If
        '            End If

        '    End Select
        'Catch ex As Exception

        '    If ex.HResult = -2146233086 Then
        '        MsgBox("Tiene Que Imprimir Como Minimo 1 Caja y 1 Lote del pedio Actual")
        '        txtCaja.Value -= 1

        '    End If

        'End Try

    End Sub

    Private Sub txtpalet_ValueChanged(sender As Object, e As EventArgs) Handles txtpalet.ValueChanged
        Dim st As New System.Diagnostics.StackTrace()
        Dim nombreMetodo As String
        nombreMetodo = st.GetFrame(3).GetMethod().Name
        Select Case nombreMetodo
            Case "UpButton"
                If txtpalet.Value > palets Then
                    MessageBox.Show("El Palet Inicial no puede Ser Mayor a la cantidad de Palets a Imprimir", "ATENCION", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtpalet.Value -= 1
                Else
                    'restamos al numero de cajas totales la cantidad de cajas que lleva un Palet
                    numCajas.Value = (numCajas.Value - numCajasPorPalet.Value)
                    'ponemos el numero de caja inical la cantidad de cajas por palet 
                    txtCaja.Value = (txtCaja.Value + numCajasPorPalet.Value)
                    'restamos al numero de lotes por imprimir la cantidad de lotes que lleva un Palet
                    numLotes.Value = (numLotes.Value - (numLotesPorCaja.Value * numCajasPorPalet.Value))
                    numPalets.Value -= 1
                End If
            Case "DownButton"
                'sumamos a las cajas totales la cantidad de cajas que lleva un palet
                numCajas.Value = (numCajas.Value + numCajasPorPalet.Value)
                txtCaja.Value = (txtCaja.Value - numCajasPorPalet.Value)
                numLotes.Value = (numLotes.Value + (numLotesPorCaja.Value * numCajasPorPalet.Value))
                numPalets.Value += 1
            Case "ParseEditText"
                If txtpalet.Value > palets Then
                    MessageBox.Show("El Palet Inicial no puede Ser Mayor a la cantidad de Palets a Imprimir", "ATENCION", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtpalet.Value = palets
                Else
                    Dim npaletimp As Integer
                    npaletimp = palets - txtpalet.Value
                    txtpalet.Value = npaletimp
                    numCajas.Value = numCajas.Value - (numCajasPorPalet.Value * (txtpalet.Value - 1))
                    txtCaja.Value = txtCaja.Value - (numCajasPorPalet.Value * (txtpalet.Value - 1))
                    numLotes.Value = numLotes.Value + (numLotesPorCaja.Value * (numCajasPorPalet.Value * (txtpalet.Value - 1)))
                End If
        End Select
    End Sub

    Private Sub txtnumsap_TextChanged(sender As Object, e As EventArgs) Handles txtnumsap.TextChanged
        numPedido = txtnumsap.Text.Trim
    End Sub

    Private Sub txtRefCliente_TextChanged(sender As Object, e As EventArgs) Handles txtRefCliente.TextChanged
        refCliente = txtRefCliente.Text.Trim
    End Sub
#Region "Impresoras"
    Private Sub cListImp_DropDown(sender As Object, e As EventArgs) Handles cListImp.DropDown
        'Buscamos las impresoras que tenemos instaladas
        Dim pkListPrinters As String
        For Each pkListPrinters In Drawing.Printing.PrinterSettings.InstalledPrinters
            cListImp.Items.Add(pkListPrinters)
        Next pkListPrinters

    End Sub
    Private Sub cListImp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cListImp.SelectedIndexChanged
        cListImp.Text = sender.Text
        'Guardamos en el fichero INI la impresora seleccionada
        INIWrite(My.Application.Info.DirectoryPath & "\settings.ini", My.Computer.Name, "Impresora", cListImp.Text)
    End Sub
#End Region
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        mostrarEtiqLote(ListBox1.SelectedItem)
        'Si la etiqueta es Ninguna deshabilitamos el boton de imprimir etiqueta.
        If ListBox1.SelectedItem = "Ninguna" Then cmdImprimirLotes.Enabled = False
        mostrarEtiqPalet("palet")
    End Sub

    Private Sub rImpCajdsd_ValueChanged(sender As Object, e As EventArgs) Handles rImpCajdsd.ValueChanged
        Dim st As New System.Diagnostics.StackTrace()
        Dim nombreMetodo As String
        nombreMetodo = st.GetFrame(3).GetMethod().Name
        Try
            Select Case nombreMetodo
                Case "UpButton"
                    'si pulsamos la flecha UP de la caja numerica para sumar 
                    numLotes.Value = (numLotes.Value - numLotesPorCaja.Value)
                    numCajas.Value -= 1
                    'If (rImpCajdsd.Value + cantCaja) < TotalCajas Then 'Or txtCaja.Value > (CajaInicial + numCajas.Value)
                    '    MsgBox("El Numero de Caja Inicial No puede ser Mayor al Total de Cajas a Imprimir", MsgBoxStyle.Information)
                    '    numLotes.Value = (numLotes.Value + numLotesPorCaja.Value)
                    '    numCajas.Value += 1
                    '    rImpCajdsd.Value -= 1
                    'End If
                Case "DownButton"
                    'si pulsamos la flecha Down de la caja numerica para restar
                    numLotes.Value = (numLotes.Value + numLotesPorCaja.Value)
                    numCajas.Value += 1
                    'If rImpCajdsd.Value < CajaInicial Then
                    '    MsgBox("El numero de Caja No puede ser Inferior al Numero de Caja Inicial", MsgBoxStyle.Information)
                    '    numLotes.Value = (numLotes.Value - numLotesPorCaja.Value)
                    '    numCajas.Value -= 1
                    '    rImpCajdsd.Value += 1
                    'End If
                Case "ParseEditText"
                    'si ponemos el numero  a mano.
                    'If rImpCajdsd.Value < CajaInicial Then

                    '    MsgBox("El Numero de Caja no puede ser Menor al Numero de Caja Inicial", MsgBoxStyle.Information)
                    '    rImpCajdsd.Value = CajaInicial
                    'Else
                    '    If rImpCajdsd.Value > (TotalCajas - 1) Then
                    '        MsgBox("El Numero de Caja Inicial No puede ser Mayor al Total de Cajas a Imprimir", MsgBoxStyle.Information)
                    '        rImpCajdsd.Value = CajaInicial
                    '    Else
                    Dim ncajImp As Integer
                    ncajImp = (rImpCajHst.Value - rImpCajdsd.Value) + 1

                    numCajas.Value = ncajImp
                    numLotes.Value = (numLotesPorCaja.Value * ncajImp)
                    '    End If
                    'End If

            End Select
        Catch ex As Exception

            If ex.HResult = -2146233086 Then
                MsgBox("Revise la Caja Inicial")
                rImpCajdsd.Value = rImpCajdsd.Minimum

            End If

        End Try
        'damos valor a la variable por la cual se debe imprimir la caja inicial.
        empezarXCaja = rImpCajdsd.Value
        'Si la caja desde y la caja hasta son la misma activamos la casilla de lote hasta
        If rImpCajdsd.Value = rImpCajHst.Value Then
            rImpLotHst.Enabled = True
        Else
            rImpLotHst.Enabled = False
        End If
    End Sub

    Private Sub rImpCajHst_ValueChanged(sender As Object, e As EventArgs) Handles rImpCajHst.ValueChanged
        Dim st As New System.Diagnostics.StackTrace()
        Dim nombreMetodo As String
        nombreMetodo = st.GetFrame(3).GetMethod().Name
        Try
            Select Case nombreMetodo
                Case "UpButton"
                    'si pulsamos la flecha UP de la caja numerica para sumar 
                    numLotes.Value = (numLotes.Value + numLotesPorCaja.Value)
                    numCajas.Value += 1
                    'If (rImpCajHst.Value + cantCaja) > TotalCajas - 1 Then 'Or txtCaja.Value > (CajaInicial + numCajas.Value)
                    '    MsgBox("El Numero de Caja Final No puede ser Mayor al Total de Cajas a Imprimir", MsgBoxStyle.Information)
                    '    numLotes.Value = (numLotes.Value - numLotesPorCaja.Value)
                    '    numCajas.Value -= 1
                    '    rImpCajHst.Value -= 1
                    'End If
                Case "DownButton"
                    'si pulsamos la flecha Down de la caja numerica para restar
                    numLotes.Value = (numLotes.Value - numLotesPorCaja.Value)
                    numCajas.Value -= 1
                    'If rImpCajHst.Value < CajaInicial Then
                    '    MsgBox("El numero de Caja No puede ser Inferior al Numero de Caja Inicial", MsgBoxStyle.Information)
                    '    numLotes.Value = (numLotes.Value + numLotesPorCaja.Value)
                    '    numCajas.Value += 1
                    '    rImpCajHst.Value -= 1
                    'End If
                Case "ParseEditText"
                    'si ponemos el numero  a mano.
                    'If rImpCajHst.Value < CajaInicial Then

                    '    MsgBox("El Numero de Caja no puede ser Menor al Numero de Caja Inicial", MsgBoxStyle.Information)
                    '    rImpCajHst.Value = CajaInicial
                    'Else
                    '    If rImpCajHst.Value > (TotalCajas - 1) Then
                    '        MsgBox("El Numero de Caja Inicial No puede ser Mayor al Total de Cajas a Imprimir", MsgBoxStyle.Information)
                    '        rImpCajHst.Value = CajaInicial
                    '    Else
                    Dim ncajImp As Integer
                    ncajImp = (rImpCajHst.Value - rImpCajdsd.Value) + 1
                    'If ncajImp = 0 Then ncajImp = 1
                    numCajas.Value = ncajImp
                    numLotes.Value = (numLotesPorCaja.Value * ncajImp)
                    'End If
                    'End If

            End Select
        Catch ex As Exception

            If ex.HResult = -2146233086 Then
                MsgBox("Revise las cajas Finales")
                rImpCajHst.Value = rImpCajHst.Maximum

            End If

        End Try
        'Si la caja desde y la caja hasta son la misma activamos la casilla de lote hasta
        If rImpCajdsd.Value = rImpCajHst.Value Then
            rImpLotHst.Enabled = True
        Else
            rImpLotHst.Enabled = False
        End If
    End Sub

    Private Sub rImpLotDsd_ValueChanged(sender As Object, e As EventArgs) Handles rImpLotDsd.ValueChanged
        'If rImpLotDsd.Value > numLotesPorCaja.Value Then
        '    MsgBox("El Numero de Lote Inicial no puede ser Mayor Que los lotes por Caja", MsgBoxStyle.Information)
        '    rImpLotDsd.Value = numLotesPorCaja.Value
        'Else
        Dim st As New System.Diagnostics.StackTrace()
        Dim nombreMetodo As String
        nombreMetodo = st.GetFrame(3).GetMethod().Name
        Select Case nombreMetodo
            Case "UpButton"
                numLotes.Value -= 1

            Case "DownButton"
                numLotes.Value += 1

        End Select
        empezarXlote = rImpLotDsd.Value
        If rImpLotHst.Enabled Then rImpLotHst.Minimum = rImpLotDsd.Value
        'End If
    End Sub

    Private Sub rImpLotHst_ValueChanged(sender As Object, e As EventArgs) Handles rImpLotHst.ValueChanged
        Dim st As New System.Diagnostics.StackTrace()
        Dim nombreMetodo As String
        nombreMetodo = st.GetFrame(3).GetMethod().Name
        Select Case nombreMetodo
            Case "UpButton"
                numLotes.Value += 1
            Case "DownButton"
                numLotes.Value -= 1

        End Select
        rImpLotDsd.Maximum = rImpLotHst.Value
    End Sub

    Private Sub txtCodPedido_TextChanged(sender As Object, e As EventArgs) Handles txtCodPedido.TextChanged

    End Sub

    Private Sub btn_MuestraLotes_Click(sender As Object, e As EventArgs) Handles btn_MuestraLotes.Click

        Imprimir_Lotes(1, True)

    End Sub

    Private Sub btn_Muestra_Cajas_Click(sender As Object, e As EventArgs) Handles btn_Muestra_Cajas.Click

        Imprimir_Cajas(1, 1, True)

    End Sub

    Private Sub btn_Muestra_Palets_Click(sender As Object, e As EventArgs) Handles btn_Muestra_Palets.Click

        'Imprimir_Palets(1, True)
        Imprimir_Palets_Nuevo(1, True)

    End Sub

    Private Sub cmImpManual_Click(sender As Object, e As EventArgs) Handles cmdImpManual.Click
        'si vamos a modificar las etiquetas y es la primera vez que sacamos este pedido
        'Creamos la linea 0
        'If txtCaja.Value = 1 AndAlso reImprimir = False Then txtSemana.Enabled = True
        If reImprimir = False Then
            grabarHistorico("insertar")
            reImprimir = True
        End If
        etiManual = True
        ' txtCaja.Enabled = True
        ' txtLote.Enabled = True
        'txtpalet.Enabled = True
        'numCajasPorPalet.Enabled = True
        'numLotes.Enabled = True
        'numCajas.Enabled = True
        'numPalets.Enabled = True
        ListBox1.Enabled = True
        ListBox2.Enabled = True
        grpReImp.Visible = True
        If grpReImp.Visible Then
            'si la El cuadro de Reimpresion esta visible.
            rImpCajdsd.Value = empezarXCaja
            'ponemos los maximos y minimos
            rImpCajdsd.Minimum = empezarXCaja
            rImpCajdsd.Maximum = empezarXCaja + (cantCaja - 1)
            rImpCajHst.Value = empezarXCaja + (cantCaja - 1)
            'ponemos los maximos y minimos
            rImpCajHst.Minimum = empezarXCaja
            rImpCajHst.Maximum = empezarXCaja + (cantCaja - 1)
            rImpLotDsd.Value = 1
            rImpLotDsd.Maximum = numLotesPorCaja.Value
            rImpLotHst.Value = numLotesPorCaja.Value
            rImpLotHst.Maximum = numLotesPorCaja.Value
        End If
    End Sub



    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        mostrarEtiqCaja(ListBox2.SelectedItem)
    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        'Al cerrar el formulario activamos la hibernación.
        suspension.actHibernacion()
        'Al cerrar el programa actualizamos la base de datos de bloqueo.
        liberarBloqueo()
    End Sub
    Private Sub resetDsdHst()
        'Procedimiento que resetea los valores maximos y minimos de las cajas de lote, cuando introduzcamos un pedido nuevo.
        rImpCajdsd.Minimum = 1
        rImpCajdsd.Maximum = 9999
        rImpCajHst.Minimum = 1
        rImpCajHst.Maximum = 9999
        rImpLotDsd.Minimum = 1
        rImpLotDsd.Maximum = 999
        rImpLotHst.Minimum = 1
        rImpLotHst.Maximum = 999

    End Sub
End Class
