Module etiquetas




    Public Function vistaPreviaEtiqueta(clasetiq As Integer, tipoetiq As String) As String()
        Dim viewEtiqueta As String()
        Dim ruta, altoEtiq, anchoEtiq As String
        'valores por defecto del tamaño del contenedor de la imagen. Estan en el fichero Settings.ini
        altoEtiq = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "IMAGENES", "alto_img")
        anchoEtiq = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "IMAGENES", "ancho_img")
        ruta = iniaccess.INI_Read(My.Application.Info.DirectoryPath & "\settings.ini", "IMAGENES", "ruta_img")
        ReDim viewEtiqueta(2)
        If clasetiq = 0 Then
            'etiqueta lote
            Select Case tipoetiq
                Case "Std", "tipo12"
                    viewEtiqueta(0) = ruta & "DML_standart.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo3"
                    viewEtiqueta(0) = ruta & "DML1003.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "Tipo 26"
                    viewEtiqueta(0) = ruta & "DML1026.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo5"
                    viewEtiqueta(0) = ruta & "DML1005.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo7"
                    viewEtiqueta(0) = ruta & "DML1007.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo9"
                    viewEtiqueta(0) = ruta & "DML1009.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo13"
                    viewEtiqueta(0) = ruta & "DML1013.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo14"
                    viewEtiqueta(0) = ruta & "DML1014.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo15"
                    viewEtiqueta(0) = ruta & "DML1015.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "report35"
                    viewEtiqueta(0) = ruta & "DML1035.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "report37"
                    viewEtiqueta(0) = ruta & "DML1037.jpg"
                    viewEtiqueta(1) = altoEtiq - 5
                    viewEtiqueta(2) = anchoEtiq + 5
                Case "report39", "report43"
                    viewEtiqueta(0) = ruta & "DML1039.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo25", "Tipo 25 (DML1025)"
                    viewEtiqueta(0) = ruta & "DML1025.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "TIPO 41"
                    viewEtiqueta(0) = ruta & "DML1041.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "Tipo 42"
                    viewEtiqueta(0) = ruta & "DML1042.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "Tipo 43"
                    viewEtiqueta(0) = ruta & "DML1043.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo32"
                    viewEtiqueta(0) = ruta & "DML1032.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo31"
                    viewEtiqueta(0) = ruta & "DML1031.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo29"
                    viewEtiqueta(0) = ruta & "DML1029.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo30"
                    viewEtiqueta(0) = ruta & "DML1030.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo33"
                    viewEtiqueta(0) = ruta & "DML10033.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "Tipo44"
                    viewEtiqueta(0) = ruta & "DML1044.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "Tipo 46"
                    viewEtiqueta(0) = ruta & "DML1046.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case Else
                    viewEtiqueta(0) = ruta & "etiqueta_gen.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
            End Select
        ElseIf clasetiq = 1 Then
            'etiqueta caja
            Select Case tipoetiq
                Case "Std"
                    viewEtiqueta(0) = ruta & "DMC_standart.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo1"
                    viewEtiqueta(0) = ruta & "DMC10001.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo3"
                    viewEtiqueta(0) = ruta & "DMC10003.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo4"
                    viewEtiqueta(0) = ruta & "DMC10004.jpg"
                    viewEtiqueta(1) = anchoEtiq + 3
                    viewEtiqueta(2) = anchoEtiq + 3
                Case "tipo5"
                    viewEtiqueta(0) = ruta & "DMC10005.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo6"
                    viewEtiqueta(0) = ruta & "DMC10006.jpg"
                    viewEtiqueta(1) = anchoEtiq + 3
                    viewEtiqueta(2) = anchoEtiq + 3
                Case "tipo7"
                    viewEtiqueta(0) = ruta & "DMC10004.jpg"
                    viewEtiqueta(1) = anchoEtiq + 3
                    viewEtiqueta(2) = anchoEtiq + 3
                Case "tipo9"
                    viewEtiqueta(0) = ruta & "DMC10009.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo11"
                    viewEtiqueta(0) = ruta & "DMC10011.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo12"
                    viewEtiqueta(0) = ruta & "DMC10012.jpg"
                    viewEtiqueta(1) = anchoEtiq + 3
                    viewEtiqueta(2) = anchoEtiq + 3
                Case "tipo13"
                    viewEtiqueta(0) = ruta & "DMC10013.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo14"
                    viewEtiqueta(0) = ruta & "DMC10014.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo15"
                    viewEtiqueta(0) = ruta & "DMC10015.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "report36", "report42", "report45"
                    viewEtiqueta(0) = ruta & "DMC10036.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "report38"
                    viewEtiqueta(0) = ruta & "DMC10038.jpg"
                    viewEtiqueta(1) = altoEtiq - 5
                    viewEtiqueta(2) = anchoEtiq + 5
                Case "tipo25", "Tipo 25 (DMC10025)"
                    viewEtiqueta(0) = ruta & "DMC10025.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "report40", "report44"
                    viewEtiqueta(0) = ruta & "DMC10040.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "Tipoar"
                    viewEtiqueta(0) = ruta & "DMC100AR2.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo32"
                    viewEtiqueta(0) = ruta & "DMC10032.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo31"
                    viewEtiqueta(0) = ruta & "DMC10031.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo29"
                    viewEtiqueta(0) = ruta & "DMC10029.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo30"
                    viewEtiqueta(0) = ruta & "DMC10030.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo33"
                    viewEtiqueta(0) = ruta & "DMC10033.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "TIPO 41"
                    viewEtiqueta(0) = ruta & "DMC10041.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "Tipo 42"
                    viewEtiqueta(0) = ruta & "DMC10042.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "Tipo 43"
                    viewEtiqueta(0) = ruta & "DMC10043.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "Tipo44"
                    viewEtiqueta(0) = ruta & "DMC10044.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "Tipo 46"
                    viewEtiqueta(0) = ruta & "DMC10046.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case Else
                    viewEtiqueta(0) = ruta & "etiqueta_gen.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq

            End Select
        Else
            viewEtiqueta(0) = ruta & "label_palet.jpg"
            viewEtiqueta(1) = anchoEtiq
            viewEtiqueta(2) = altoEtiq
        End If
        Return viewEtiqueta
    End Function

#Region "Etiquetas Lotes"

    Public Sub printLoteTipo27()
        'Impresión de etiqueta standart de lote tamaño etiqueta 60mm x 47mm
        'asignamos el tipo de etiqueta

        Dim lastCodeChip As Long

        lastCodeChip = Long.Parse(tempFirstCodeLot) + (jerxLote)


        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        label.SetNamedSubStringValue("From", tempFirstCodeLot)
        label.SetNamedSubStringValue("To", lastCodeChip - 1)
        label.SetNamedSubStringValue("siglapais", sigla)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)

        tempFirstCodeLot = lastCodeChip

    End Sub
    Public Sub printLoteStandart()
        'Impresión de etiqueta standart de lote tamaño etiqueta 60mm x 47mm
        'asignamos el tipo de etiqueta
        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)


    End Sub
    Public Sub printLoteTipo3()
        'impresion de etiquetas Tipo3 de lote tamaño etiqueta 40mm x 50mm
        label.SetNamedSubStringValue("barcode", numLote)
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)
        'objbt.Visible = True
    End Sub
    Public Sub printLoteTipo5()
        'impresion de etiqueta Tipo5 de lote tamaño etiqueta 60mm x 47mm

        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)


    End Sub
    Public Sub printLoteTipo7()
        'impresion de etiquetas Tipo7 de lote tamaño etiqueta 40mm x 50mm
        Dim barcode, nlotelim As String
        'quitamos los simbolos raros
        nlotelim = Replace(numLote, "_", "")
        nlotelim = Replace(nlotelim, "/", "")
        'creamos el codigo de barras especifico de la etiqueta
        barcode = "(10)" & nlotelim & "(17)" & Replace(fecCaducidad, "-", "")
        label.SetNamedSubStringValue("barcode", barcode)
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)
        'objbt.Visible = True
    End Sub
    Public Sub printLoteTipo9()
        'impresion etiquetas de Lote tamaño etiqueta 55mm x 40mm
        label.SetNamedSubStringValue("barcode", numLote)
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)
        'objbt.Visible = True
    End Sub
    Public Sub printLoteTipo13()
        'Impresion de etiqueta Tipo13 de lote de tamaño 55mm x 40mm
        label.SetNamedSubStringValue("barcode", numLote)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("qtbox", jerxLote)
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
    End Sub
    Public Sub printLoteTipo15()
        'impresion de etiqueta Tipo15 de lote tamaño etiqueta 60mm x 47mm

        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)


    End Sub
    Public Sub printReport35()
        'impresion de etiqueta report35 de lote tamaño 55mm x 40mm
        Dim lastCodeChip As Long

        lastCodeChip = Long.Parse(tempFirstCodeLot) + (jerxLote)
        tempLastCodeLot = lastCodeChip.ToString("000000000000000")
        label.SetNamedSubStringValue("barcode", numLote)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("firstcode", tempFirstCodeLot)
        label.SetNamedSubStringValue("lastcode", (lastCodeChip - 1))
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'asignamos a la variable tempFirstCode el valor calculado de tempLastCode, para que sea el codigo del primer chip si tenemos que imprimir
        'mas de un lote.
        tempFirstCodeLot = tempLastCodeLot


    End Sub
    Public Sub printReport37()
        'impresión de etiqueta report37 de lote tamaño 75mm x 110mm
        Dim codechip As Long

        codechip = Long.Parse(tempFirstCodeLot)
        tempLastCodeLot = tempFirstCodeLot
        label.SetNamedSubStringValue("barcodeLot", numLote)
        label.SetNamedSubStringValue("numlote", Left(numLote, 10))
        label.SetNamedSubStringValue("numsmallbox", "0" & Right(numLote, 2))
        label.SetNamedSubStringValue("sigpais", Left(sigla, 2))
        label.SetNamedSubStringValue("rangecode", Mid(tempFirstCodeLot, 4, 7))
        'rellenamos los 10 digitos de los chips de las jeringas.
        For i As Integer = 1 To 10
            label.SetNamedSubStringValue("code" & i, Right(tempLastCodeLot, 5))
            label.SetNamedSubStringValue("barcode" & i & "cod", codechip)
            codechip += 1
            tempLastCodeLot = codechip.ToString("000000000000000")
        Next
        tempFirstCodeLot = tempLastCodeLot
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub
    Public Sub printReport39()
        'impresion de etiqueta report39 de lote tamaño 55mm x 40mm

        label.SetNamedSubStringValue("barcode", numLote)
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("numarticulo", refArticulo)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("numorden", numPedido)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub


    Public Sub printLote44()
        'impresion de etiqueta report39 de lote tamaño 55mm x 40mm

        'label.SetNamedSubStringValue("barcode", numLote)
        label.SetNamedSubStringValue("numLote", numLote)
        label.SetNamedSubStringValue("numarticulo", refArticulo)
        label.SetNamedSubStringValue("description", refCliente)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub

    Public Sub printTipo28()
        'impresion de etiqueta report39 de lote tamaño 55mm x 40mm

        'label.SetNamedSubStringValue("barcode", numLote)
        label.SetNamedSubStringValue("numlote", numLote)
        'label.SetNamedSubStringValue("numarticulo", refArticulo)
        'label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("numorden", numPedido)
        label.SetNamedSubStringValue("feccadu", fecCaducidad.Replace("EXP   ", "").Replace(".", "/"))
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub

    Public Sub printLoteTVA()
        label.SetNamedSubStringValue("numLote", numLote)
        label.SetNamedSubStringValue("descripcion", descArticulo.ToUpper)
        label.SetNamedSubStringValue("qtlote", jerxLote)
        label.SetNamedSubStringValue("fechaCad", fecCaducidad)
        label.PrintOut(False, False)
    End Sub
    Public Sub printLoteTVAO(Optional ByVal ordenado As Boolean = True)
        Dim lastCodeChip As Long

        lastCodeChip = Long.Parse(tempFirstCodeLot) + (jerxLote)
        tempLastCodeLot = lastCodeChip.ToString("000000000000000")
        label.SetNamedSubStringValue("numLote", numLote)
        label.SetNamedSubStringValue("descripcion", descArticulo)
        label.SetNamedSubStringValue("qtlote", jerxLote)
        label.SetNamedSubStringValue("fechaCad", fecCaducidad)
        If ordenado Then
            label.SetNamedSubStringValue("chipInicio", Decimal.Parse(tempFirstCodeLot).ToString("000000000000000"))
            label.SetNamedSubStringValue("chipFinal", (lastCodeChip - 1).ToString("000000000000000"))
        End If

        ' objbt.Visible = True
        label.PrintOut(False, False)
        'asignamos a la variable tempFirstCode el valor calculado de tempLastCode, para que sea el codigo del primer chip si tenemos que imprimir
        'mas de un lote.
        tempFirstCodeLot = tempLastCodeLot
    End Sub

    Public Sub printLoteTipo43()
        Dim lastCodeChip As Long

        lastCodeChip = Long.Parse(tempFirstCodeLot) + (jerxLote)
        tempLastCodeLot = lastCodeChip.ToString("000000000000000")
        label.SetNamedSubStringValue("numLote", numLote)
        label.SetNamedSubStringValue("fechaCad", fecFabricacion.Split("/")(1) & "/" & fecFabricacion.Split("/")(0))
        label.SetNamedSubStringValue("chipInicio", "A0060000" & Decimal.Parse(tempFirstCodeLot).ToString("000000000000000"))
        label.SetNamedSubStringValue("chipFinal", "A0060000" & (lastCodeChip - 1).ToString("000000000000000"))
        ' objbt.Visible = True
        label.PrintOut(False, False)
        'asignamos a la variable tempFirstCode el valor calculado de tempLastCode, para que sea el codigo del primer chip si tenemos que imprimir
        'mas de un lote.
        tempFirstCodeLot = tempLastCodeLot
    End Sub
#End Region
#Region "Etiquetas Cajas"

    Public Sub printCajaStandart()
        'impresión de etiqueta standart de Caja tamaño 60mm x 47mm

        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub

    Public Sub printCajaTipo27()
        'impresión de etiqueta standart de Caja tamaño 60mm x 47mm

        Dim lastcode As Long

        lastcode = (Long.Parse(tempFirtCode)) + ((jerxLote * lotxCaja))

        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("From", tempFirtCode)
        label.SetNamedSubStringValue("To", lastcode - 1)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)

        tempFirtCode = lastcode
    End Sub

    Public Sub printCajaTipo1()
        'impresión de etiqueta Tipo1 de Caja tamaño 60mm x 47mm

        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub
    Public Sub printCajaTipo3()
        'impresión de etiqueta Tipo3 de Caja tamaño 60mm x 47mm

        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub
    Public Sub printCajaTipo4()


        'impresion de etiqueta Tipo4 de Caja tamaño 80mm x 80mm
        label.SetNamedSubStringValue("refCliente", refCliente)
        label.SetNamedSubStringValue("numCaja", numCaja)
        label.SetNamedSubStringValue("fecCadu", fecCaducidad)
        label.SetNamedSubStringValue("barcodeprod", "(02)7640369(91)" & refCliente & "(37)100")
        label.SetNamedSubStringValue("barcodelote", "(10)" & numCaja.Trim & "(17)" & Replace(fecCaducidad, "-", ""))
        label.SetNamedSubStringValue("fecFabri", fecSealing)
        'objbt.Visible = True
        label.PrintOut()
    End Sub
    Public Sub printCajaTipo5()
        'impresión de etiqueta Tipo5 de Caja tamaño 60mm x 47mm

        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub

    Public Sub printCajaTipo9()
        'impresion de etiqueta Tipo 9 de Caja tamaño 40mm x 50mm
        label.SetNamedSubStringValue("barcode", numCaja)
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)
        'objbt.Visible = True
    End Sub
    Public Sub printCajaTipo10()
        Dim datosMatix As String
        'impresion de etiqueta Tipo 10 de Caja tamaño 40mm x 50mm
        label.SetNamedSubStringValue("descripcion", Left(descArticulo, 4))
        label.SetNamedSubStringValue("sigla", sigla)
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        label.SetNamedSubStringValue("refcliente", refCliente)
        datosMatix = refCliente & "-" & numCaja & "-" & fecCaducidad
        label.SetNamedSubStringValue("datamatrix", datosMatix)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)
    End Sub
    Public Sub printCajaTipo11()
        'impresión de etiqueta Tipo11 de Caja tamaño 60mm x 47mm

        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub
    Public Sub printCajaTipo12()
        'impresion de etiqueta Tipo12 de Caja Tamaño 80x80mm. Modelo DM spagna_0700_0500_011
        Dim material, peso, regZooSan As String

        'If descArticulo = "T-IT 8100" Or descArticulo = "T-IS 8010" Then
        '    material = "vidrio"
        '    peso = "0,114 grs"
        '    regZooSan = "02694-MUZ"
        'ElseIf descArticulo = "T-SL 8010" Then
        material = "polimero"
        peso = "0,044 grs"
        regZooSan = "02995-MUZ"
        'Else
        '    material = ""
        '    peso = ""
        '    regZooSan = ""
        'End If
        label.SetNamedSubStringValue("material", material)
        label.SetNamedSubStringValue("tipoproducto", descArticulo)
        label.SetNamedSubStringValue("peso", peso)
        label.SetNamedSubStringValue("feccaduc", fecCaducidad)
        label.SetNamedSubStringValue("regzoosan", regZooSan)
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("fecfabrica", fecFabricacion)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True

        label.PrintOut(False, False)


    End Sub
    Public Sub printCajaTipo13()
        label.SetNamedSubStringValue("barcode", numCaja)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("qtbox", (jerxLote * lotxCaja))
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
    End Sub
    Public Sub printcajaTipo14()
        label.SetNamedSubStringValue("barcode", numCaja)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("qtbox", (jerxLote * lotxCaja))
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
    End Sub
    Public Sub printCajaTipo15()
        'impresión de etiqueta Tipo15 de Caja tamaño 60mm x 47mm

        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub
    Public Sub printCajaReport36(reporte As Integer)
        'impresión de etiquetas Report36 de Caja tamaño 55mm x 40mm
        Dim lastcode, contBox As Long
        Dim texto As String

        lastcode = (Long.Parse(tempFirtCode) + ((jerxLote * lotxCaja)))
        tempLastCode = lastcode.ToString("000000000000000")
        If reporte = 0 Then
            'texto descripcion report36
            texto = "SYRINGE VALUE IN BOX OF 100"
        ElseIf reporte = 1 Then
            'texto descripcion report42
            texto = "STANDARD IN BOX OF  " & (jerxLote * lotxCaja)
        ElseIf reporte = 2 Then
            'texto descripcion report45
            texto = "NEEDLE KIT IN BOX OF 100"
        Else
            texto = "STANDARD IN BOX OF 100"
        End If
        contBox = ((lastcode - Long.Parse(firstCode)) / (jerxLote * lotxCaja))
        label.SetNamedSubStringValue("barcode", numCaja)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("texto", texto)
        label.SetNamedSubStringValue("firstcode", tempFirtCode)
        label.SetNamedSubStringValue("lastcode", (lastcode - 1))
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("boxnum", contBox.ToString("000"))
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        tempFirtCode = tempLastCode


    End Sub
    Public Sub printCajaReport38(count As Integer)
        'impresión de etiquetas Report38 de Caja tamaño 75mm x 100mm
        Dim lastcode As Long

        lastcode = (Long.Parse(tempFirtCode)) + ((jerxLote * lotxCaja))
        tempLastCode = lastcode.ToString("000000000000000")
        label.SetNamedSubStringValue("cajaactual", count.ToString("000")) 'Right(numCaja, 3))
        label.SetNamedSubStringValue("cajatotal", maxcajas.ToString("000")) 'Se cambia la variable de cajas a imprimir por la cantidad de cajas Maximas.
        label.SetNamedSubStringValue("cantpiezas", (jerxLote * lotxCaja))
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("barcodenumlote", numCaja)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("fecactual", fecLotImp)
        label.SetNamedSubStringValue("sigpais", Left(sigla, 2))
        label.SetNamedSubStringValue("firstcode", Right(tempFirtCode, 12))
        label.SetNamedSubStringValue("barcodefirstcode", tempFirtCode)
        'restamos una unidad al ultimo codigo para poner el numero de chip real. Ej.: 501 al 520 son 20 numeros, pero suma 19.
        label.SetNamedSubStringValue("lastcode", Right((lastcode - 1), 12))
        label.SetNamedSubStringValue("barcodelastcode", (lastcode - 1))
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        tempFirtCode = tempLastCode


    End Sub
    Public Sub printCajaReport40()
        'impresion de etiquetas Report40 de Caja tamañon 55mm x 40mm

        label.SetNamedSubStringValue("barcode", numCaja)
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("numarticulo", refArticulo)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("numorden", numPedido)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)

    End Sub

    Public Sub printCajaTipo28(cantidad As Integer)
        'impresion de etiquetas Report40 de Caja tamañon 55mm x 40mm

        label.SetNamedSubStringValue("numCaja", numCaja)
        'label.SetNamedSubStringValue("numlote", numCaja)
        'label.SetNamedSubStringValue("numarticulo", refArticulo)
        label.SetNamedSubStringValue("cantidad", cantidad)
        label.SetNamedSubStringValue("numLote", numPedido)
        '        label.SetNamedSubStringValue("fechaCad", fecCaducidad.Replace("EXP", "").Trim)

        'Dim fechaCad As String = fecCaducidad.Replace("EXP", "").Trim.Split("/")(1) & "/" & fecCaducidad.Replace("EXP", "").Trim.Split("/")(0)
        Dim fechaCad As String = fecCaducidad.Replace("EXP", "").Trim.Split(".")(0) & "/" & fecCaducidad.Replace("EXP", "").Trim.Split(".")(1)

        label.SetNamedSubStringValue("fechaCad", fechaCad)

        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)

    End Sub

    Public Sub printEtiPalet(pedsap As String, npalet As Integer, cajaini As String, cajafin As String)

        label.SetNamedSubStringValue("numsap", pedsap)
        label.SetNamedSubStringValue("referencia", refArticulo)
        label.SetNamedSubStringValue("numpalet", npalet)
        label.SetNamedSubStringValue("feccaduc", fecCaducidad)
        label.SetNamedSubStringValue("numcajaIni", cajaini)
        label.SetNamedSubStringValue("numCajaFin", cajafin)
        label.PrintOut(False, False)
        'objbt.Visible = True


    End Sub

    Public Sub printEtiPaletBrasil(pedsap As String, npalet As Integer, cajaini As String, cajafin As String, qrTmp As String)

        label.SetNamedSubStringValue("numsap", pedsap)
        label.SetNamedSubStringValue("referencia", refArticulo)
        label.SetNamedSubStringValue("numpalet", npalet)
        label.SetNamedSubStringValue("feccaduc", fecCaducidad)
        label.SetNamedSubStringValue("Lista_Palets", qrTmp)
        label.PrintOut(False, False)
        'objbt.Visible = True


    End Sub
    Public Sub printCajaTVA()
        label.SetNamedSubStringValue("numCaja", numCaja)
        label.SetNamedSubStringValue("descripcion", descArticulo.ToUpper)
        label.SetNamedSubStringValue("qtcaja", (jerxLote * lotxCaja))
        label.SetNamedSubStringValue("fechaCad", fecCaducidad)
        label.PrintOut(False, False)
    End Sub
    Public Sub printCajaTVAO(Optional ByVal ordenado As Boolean = True)
        'impresión de etiquetas Report36 de Caja tamaño 55mm x 40mm
        Dim lastcode As Long

        lastcode = (Long.Parse(tempFirtCode) + ((jerxLote * lotxCaja)))
        tempLastCode = lastcode.ToString("000000000000000")

        label.SetNamedSubStringValue("numCaja", numCaja)
        label.SetNamedSubStringValue("descripcion", descArticulo)
        label.SetNamedSubStringValue("qtcaja", (jerxLote * lotxCaja))
        label.SetNamedSubStringValue("fechaCad", fecCaducidad)
        If ordenado Then
            label.SetNamedSubStringValue("chipInicio", Decimal.Parse(tempFirtCode).ToString("000000000000000"))
            label.SetNamedSubStringValue("chipFinal", (lastcode - 1).ToString("000000000000000"))
        End If
        'objbt.Visible = True
        label.PrintOut(False, False)
        'asignamos a la variable tempFirstCode el valor calculado de tempLastCode, para que sea el codigo del primer chip si tenemos que imprimir
        'mas de un lote.
        tempFirtCode = tempLastCode

    End Sub


    Public Sub printCajaAr(contador As Integer, pedsap As String, texto As String, cajafin As Integer)

        label.SetNamedSubStringValue("saporder", pedsap)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("textosap", texto)
        label.SetNamedSubStringValue("referencia", refArticulo)
        label.SetNamedSubStringValue("numcaja", contador)
        label.SetNamedSubStringValue("numcajfin", cajafin)
        label.PrintOut(False, False)
        'objbt.Visible = True

    End Sub
    Public Sub printCaja20()
        'impresión de etiquetas Report36 de Caja tamaño 55mm x 40mm
        Dim lastcode As Long
        Dim prefijo As String = "10060000"

        lastcode = (Long.Parse(tempFirtCode) + ((jerxLote * lotxCaja)))
        tempLastCode = lastcode.ToString("000000000000000")

        label.SetNamedSubStringValue("numCaja", numCaja)
        label.SetNamedSubStringValue("descripcion", descArticulo)
        label.SetNamedSubStringValue("qtcaja", (jerxLote * lotxCaja))
        label.SetNamedSubStringValue("fechaCad", fecCaducidad)
        label.SetNamedSubStringValue("chipInicio", prefijo & tempFirtCode)
        label.SetNamedSubStringValue("chipFinal", prefijo & (lastcode - 1))
        'objbt.Visible = True
        label.PrintOut(False, False)
        'asignamos a la variable tempFirstCode el valor calculado de tempLastCode, para que sea el codigo del primer chip si tenemos que imprimir
        'mas de un lote.
        tempFirtCode = tempLastCode

    End Sub

    Public Sub printCajaTipo43()
        Dim lastCode As Long

        lastCode = (Long.Parse(tempFirtCode) + ((Form1.numUnidadesPorlote.Value * Form1.numLotesPorCaja.Value)))
        tempLastCode = lastCode.ToString("000000000000000")



        label.SetNamedSubStringValue("numCaja", numCaja)
        label.SetNamedSubStringValue("fechaCad", fecFabricacion.Split("/")(1) & "/" & fecFabricacion.Split("/")(0))
        label.SetNamedSubStringValue("chipInicio", "A0060000" & Decimal.Parse(tempFirtCode).ToString("000000000000000"))
        label.SetNamedSubStringValue("chipFinal", "A0060000" & (lastCode - 1).ToString("000000000000000"))
        ' objbt.Visible = True
        label.PrintOut(False, False)
        'asignamos a la variable tempFirstCode el valor calculado de tempLastCode, para que sea el codigo del primer chip si tenemos que imprimir
        'mas de un lote.
        tempFirtCode = tempLastCode
    End Sub


    Public Sub printCaja44()
        'impresion de etiqueta report39 de lote tamaño 55mm x 40mm

        label.SetNamedSubStringValue("numCaja", numCaja)
        label.SetNamedSubStringValue("numarticulo", refArticulo)
        label.SetNamedSubStringValue("description", refCliente)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)


    End Sub
#End Region

End Module
