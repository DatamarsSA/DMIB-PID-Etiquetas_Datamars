Module Globales
    Public usuarioactual As Usuario
    Public listausuarios As New List(Of Usuario)
    Public uspDatosPedidos As New DataSet1TableAdapters.conPedidosRefSapDesartTableAdapter
    ' Public uspRangoChip As New DataSet1TableAdapters.OrdenesGrabacionTableAdapter
    'Public dt, dt1 As DataTable
    Public tipoProducto As Integer ' variable para tipo de producto bulk o jeringa.
    Public etiManual, histGrabado, reImprimir As Boolean
    Public cantRango As Integer
    'variables publicas para rellenar las etiquetas
    Public numLote, numCaja, fecCaducidad, refArticulo, descArticulo, sigla, numPedido, fecFabricacion, refCliente, fecPedFab, fecLotImp, fecInicioReal, fecsealing As String
    Public cantCaja, cantLote, loteinicial, auxcaja, lotxCaja, jerxLote, qtAñosCaducidad As Integer
    Public firstCode, lastCode, tempFirtCode, tempLastCode As String
    Public FirstCodeLot, LastCodeLot, tempFirstCodeLot, tempLastCodeLot As String
    Public IDhistorico, CajaInicial, TotalCajas, totalLotes, maxcajas, palets, empezarXCaja, empezarXlote As Integer

    'variables para las etiquetas
    Public label As BarTender.Format
    Public etImpConf As BarTender.PrintSetup 'variable para la configuración de la impresora
    Public ruta As String
    Public objbt As New BarTender.Application
    'declaramos una variable de la clase hibernar para activar o desactivar la hibernación de los equipos.
    Public suspension As New hibernar
    Public compBloqueo As Boolean = False

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
    Public Sub espera(ByVal tiempo As Double)
        Static tiempo_final As DateTime
        Static tiempo_actual As DateTime
        tiempo_actual = Now
        tiempo_final = tiempo_actual.AddSeconds(tiempo)
        Do
            tiempo_actual = Now
            Application.DoEvents()
            If tiempo_actual > tiempo_final Then
                Exit Do
            End If
        Loop
    End Sub
    Public Function compTipoEti(valor As String) As Boolean
        Dim resultado As Boolean
        Select Case valor
            Case "tipo3", "tipo7", "tipo9", "tipo13", "tipo10", "report48", "report49"
                resultado = False
            Case "tipo4", "tipo6", "tipo7", "tipo9", "tipo10", "tipo13", "tipo14"
                resultado = False
            Case Else
                resultado = True
        End Select
        Return resultado
    End Function
    Public Sub cargaFechasPedido(idpedido As Integer)
        'Procedimiento para sacar las fechas de produccion y caducidad del pedido.
        Dim BusFecPrCa As New DataSet1TableAdapters.EMB_ChipsTableAdapter
        Dim BusFecPrCa2 As New DataSet1TableAdapters.EMB_Chips1TableAdapter
        Dim fecPedido As DataTable
        Dim fecPedCad, añocad, mescad, semana, expRegular As String
        'Dim fechaTmp As Date
        Dim fecvalida As Boolean = False
        If fecInicioReal = "" Then
            fecPedido = BusFecPrCa2.GetData(idpedido)

            'si el pedido es de la calle K no hace falta buscar la fecha caducidad del Chip lo calculamos.    
            ' fecPedido = BusFecPrCa.GetData(idpedido)

        End If

        'If Not fecPedido.Rows.Count > 0 Then
        '    'Si no hay datos en la grabacion de Chips. Miramos en la Tabla arreglos.
        '    Dim BusARRFecPrCa As New DataSet1TableAdapters.EMB_ArreglosTableAdapter
        '    fecPedido = BusARRFecPrCa.GetData(idpedido)
        'End If
        'volvemos a comprobar si hay datos.
        If Not String.IsNullOrEmpty(fecInicioReal) Then
            'Si tenemos fecha de inicio real sacamos la fecha de caducidad del articulo.
            Dim año, mes As String
            año = Mid(fecInicioReal, 7, 4)
            mes = Mid(fecInicioReal, 4, 2)
            'Sacamos año de fabricacion antes de sumarle la caducidad.
            fecFabricacion = año & "/" & mes
            If Form1.txtRefProducto.Text.StartsWith("992") Or Form1.txtRefProducto.Text.StartsWith("995") Then
                ''If Form1.txtRefProducto.Text = "992 0000-MKD" Then
                ''    año = (Integer.Parse(año) + 3).ToString("0000")
                ''Else
                ''    año = (Integer.Parse(año) + 5).ToString("0000")
                ''End If

                Dim consulta As New DataSet1TableAdapters.qryInsertHistorico
                Dim añosCad As Integer = consulta.GetAñosCaducidad(Form1.txtRefProducto.Text.Trim)

                If añosCad = 0 Then
                    año = (Integer.Parse(año) + 5).ToString("0000")
                Else
                    año = (Integer.Parse(año) + añosCad).ToString("0000")
                End If

            ElseIf Form1.txtrefproducto.Text.StartsWith("996") Then
                Dim consulta As New DataSet1TableAdapters.qryInsertHistorico
                año = año + consulta.GetAñosCaducidad(Form1.txtRefProducto.Text.Trim)
            Else
                año = (Integer.Parse(año) + 5).ToString("0000")
            End If
            'damos el valor a la variable con la fecha de caducidad calculada segun el producto.
            ''If Form1.txtRefProducto.Text = "992 0000-FOA" Then
            ''    año = (Integer.Parse(Mid(fecInicioReal, 7, 4)) + 4).ToString("0000")
            ''End If

            'año = consulta.GetAñosCaducidad(txtRefProducto.Text.Trim)

            fecPedCad = año & mes
                fecPedFab = Left(fecInicioReal, 10)
                'ElseIf fecPedido.Rows.Count > 0 Then
                '    'Si hay datos cogemos las fechas.
                '    fecPedFab = Left(fecPedido.Rows(0)(0), 8) ' cogemos los primeros 8 caracteres
                '    fecPedCad = fecPedido.Rows(0)(1)
                '    If fecPedCad = 0 Then
                '        'Si no tiene fecha de caducidad ponemos la fecha de grabación del Pedido. Esto es para los pedidos Bulk.
                '        fecPedCad = fecPedido.Rows(0)("Fecha")
                '        fecPedCad = Left(fecPedCad, 6)
                '    End If


            Else
                expRegular = "^((20)\d\d)?((((0[13578])|(1[02]))?(([0-2][0-9])|(3[01])))|(((0[469])|(11))?(([0-2][0-9])|(30)))|(02?[0-2][0-9]))$"
            'Si no hay datos es que no estan en chip
            'Do While fecvalida = False
            '    ' fecPedFab = InputBox("Este Pedido no tiene fecha Fabricación introduzca una. Formato 'YYYYMMDD'")
            '    If System.Text.RegularExpressions.Regex.IsMatch(fecPedFab, expRegular) = True Then
            '        fecvalida = True
            '    Else
            '        MsgBox("Fecha Introducida Incorrecta, Por Favor introduzca una Correcta")
            '    End If

            'Loop
            'Ponemos a False la variable para entrar en otro bucle y cambiamos la expresión Regular con los meses.
            fecvalida = False
            expRegular = "^((20)\d\d)(0?[1-9]|1[012])$"
            Do While fecvalida = False
                fecPedCad = InputBox("Este Pedido no tiene Fecha Caducidad, Introduzca una. Formato 'YYYYMM'")
                If System.Text.RegularExpressions.Regex.IsMatch(fecPedCad, expRegular) = True Then
                    fecvalida = True
                Else
                    MsgBox("Fecha Introducida Incorrecta, Por Favor introduzca una Correcta")
                End If
            Loop
            'If refArticulo.Trim.StartsWith("992") OrElse refArticulo.Trim.StartsWith("995") Then
            '    fecPedCad = Now.Year + 5 & Now.Month
            'Else
            '    fecPedCad = Now.Year + 3 & Now.Month
            'End If
            'si hay siglas nos lo saltamos.
            If Form1.txtSiglaPais.Text.Trim = "" Then
                Form1.txtSiglaPais.Enabled = True
                Form1.txtSiglaPais.Text = InputBox("Introduzca Siglas Pais o Codigo Fabricante. Ejemplo: 'ESP o 981'")
            End If
        End If
        'damos formato a la fecha para calcular la semana
        '    fechaTmp = Right(fecPedFab, 2) & "/" & Mid(fecPedFab, 5, 2) & "/" & Left(fecPedFab, 4)
        'Form1.txtAño.Value = Integer.Parse(Mid(fecPedFab, 3, 2))
        'Dim cal As System.Globalization.Calendar = System.Globalization.DateTimeFormatInfo.CurrentInfo.Calendar
        'semana = cal.GetWeekOfYear(fechaTmp, Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday).ToString
        'Form1.txtSemana.Value = semana

        añocad = Mid(fecPedCad, 3, 2)
        mescad = Right(fecPedCad, 2)

        Form1.numAñoCad.Value = Integer.Parse(añocad)
        Form1.numMesCad.Value = Integer.Parse(mescad)

    End Sub
    Public Sub comprobarBloqueo()
        Dim busBlock As New DataSet1TableAdapters.bloqueoEtiDMTableAdapter
        Dim equipo As String
        Dim bloqueado As Boolean
        Dim reg As DataTable
        reg = busBlock.GetData(1)
        equipo = reg(0)(1)
        bloqueado = reg(0)(0)
        If Not bloqueado And equipo.Trim = "SE" Or equipo.Trim = "" Then
            ' si el equipo no esta bloqueado y no tiene equipo lo bloqueamos.
            Dim updBloqueo As New DataSet1TableAdapters.qryInsertHistorico
            Dim respuesta As Integer
            bloqueado = True
            equipo = My.Computer.Name
            respuesta = updBloqueo.uspActualizarBloqueo(bloqueado, equipo, 1)
            compBloqueo = True
        ElseIf bloqueado And Not equipo = My.Computer.Name Then

            MessageBox.Show("Actualmente esta Bloqueado por " & equipo.TrimEnd & ". Intentelo pasado unos minutos." & vbCrLf & "El programa se cerrara.", "PROGRAMA BLOQUEADO", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Form1.Close()

        End If
    End Sub
    Public Sub liberarBloqueo()
        Dim busBlock As New DataSet1TableAdapters.bloqueoEtiDMTableAdapter
        Dim equipo As String
        Dim bloqueado As Boolean
        Dim reg As DataTable
        reg = busBlock.GetData(1)
        equipo = reg(0)(1)
        bloqueado = reg(0)(0)
        If bloqueado And equipo.Trim = My.Computer.Name Then
            Dim updBloqueo As New DataSet1TableAdapters.qryInsertHistorico
            Dim respuesta As Integer
            ' si esta bloqueado y es el equipo que ha bloqueado el fichero lo liberamos.
            bloqueado = False
            equipo = "SE"
            respuesta = updBloqueo.uspActualizarBloqueo(bloqueado, equipo, 1)
            compBloqueo = False
        End If
    End Sub


End Module
