Module etiquetas
    'variables publicas para rellenar las etiquetas
    Public numLote, numCaja, fecCaducidad, refArticulo, descArticulo, sigla, numPedido, fecFabricacion As String
    Public cantCaja, cantLote, loteinicial, auxcaja, lotxCaja, jerxLote As Integer
    Public firstCode, lastCode, tempFirtCode, tempLastCode As String
    Dim objbt As New BarTender.Application

    Public Sub cerrarBartender()
        objbt.Quit(BarTender.BtSaveOptions.btDoNotSaveChanges)

    End Sub
    Public Function vistaPreviaEtiqueta(clasetiq As Integer, tipoetiq As String) As String()
        Dim viewEtiqueta As String()
        Dim ruta, altoEtiq, anchoEtiq As String
        'valores por defecto del tamaño del contenedor de la imagen.
        altoEtiq = "175"
        anchoEtiq = "200"
        ruta = "C:\Users\Ezequiel\Documents\etiquetas_datamars\imagenes\"
        ReDim viewEtiqueta(2)
        If clasetiq = 0 Then
            'etiqueta lote
            Select Case tipoetiq
                Case "Std", "tipo12"
                    viewEtiqueta(0) = ruta & "standart_LB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo5"
                    viewEtiqueta(0) = ruta & "tipo5_LB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo15"
                    viewEtiqueta(0) = ruta & "tipo15_LB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "report35"
                    viewEtiqueta(0) = ruta & "label35_LB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "report37"
                    viewEtiqueta(0) = ruta & "label37_LB.jpg"
                    viewEtiqueta(1) = altoEtiq - 5
                    viewEtiqueta(2) = anchoEtiq + 5
                Case "report39"
                    viewEtiqueta(0) = ruta & "label39_LB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case Else
                    viewEtiqueta(0) = ruta & "etiqueta_gen.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
            End Select
        Else
            'etiqueta caja
            Select Case tipoetiq
                Case "Std"
                    viewEtiqueta(0) = ruta & "standart_BB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo3"
                    viewEtiqueta(0) = ruta & "tipo3_BB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo5"
                    viewEtiqueta(0) = ruta & "tipo5_BB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo11"
                    viewEtiqueta(0) = ruta & "tipo11_BB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "tipo12"
                    viewEtiqueta(0) = ruta & "spagna_0700.jpg"
                    viewEtiqueta(1) = anchoEtiq + 3
                    viewEtiqueta(2) = anchoEtiq + 3

                Case "tipo15"
                    viewEtiqueta(0) = ruta & "tipo15_BB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "report36", "report42", "report45"
                    viewEtiqueta(0) = ruta & "label36_BB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case "report38"
                    viewEtiqueta(0) = ruta & "label38_BB.jpg"
                    viewEtiqueta(1) = altoEtiq - 5
                    viewEtiqueta(2) = anchoEtiq + 5
                Case "report40", "report44"
                    viewEtiqueta(0) = ruta & "label40_BB.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
                Case Else
                    viewEtiqueta(0) = ruta & "etiqueta_gen.jpg"
                    viewEtiqueta(1) = anchoEtiq
                    viewEtiqueta(2) = altoEtiq
            End Select
        End If
        Return viewEtiqueta
    End Function
#Region "Etiquetas Lotes"
    Public Sub printLoteStandart()
        'Impresión de etiqueta standart de lote tamaño etiqueta 60mm x 47mm
        'asignamos el tipo de etiqueta
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\lote\standart_lb.btw")
        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)

    End Sub
    Public Sub printLoteTipo5()
        'impresion de etiqueta Tipo5 de lote tamaño etiqueta 60mm x 47mm
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\lote\tipo5_lb.btw")
        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printLoteTipo15()
        'impresion de etiqueta Tipo5 de lote tamaño etiqueta 60mm x 47mm
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\lote\tipo15_lb.btw")
        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        'imprimimos la etiqueta y ponemos a false que nos muestre el estado de la impresora y el panel de propiedades
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printReport35()
        'impresion de etiqueta report35 de lote tamaño 55mm x 40mm
        Dim lastCodeChip As Long
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\lote\label35_lb.btw")
        lastCodeChip = Long.Parse(tempFirtCode) + (jerxLote)
        tempLastCode = lastCodeChip.ToString("000000000000000")
        label.SetNamedSubStringValue("barcode", numLote)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("firstcode", tempFirtCode)
        label.SetNamedSubStringValue("lastcode", (lastCodeChip - 1))
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'asignamos a la variable tempFirstCode el valor calculado de tempLastCode, para que sea el codigo del primer chip si tenemos que imprimir
        'mas de un lote.
        tempFirtCode = tempLastCode
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printReport37()
        'impresión de etiqueta report37 de lote tamaño 75mm x 110mm
        Dim codechip As Long
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\lote\label37_lb.btw")
        codechip = Long.Parse(tempFirtCode)
        tempLastCode = tempFirtCode
        label.SetNamedSubStringValue("numlote", Left(numLote, 10))
        label.SetNamedSubStringValue("numsmallbox", "0" & Right(numLote, 2))
        label.SetNamedSubStringValue("sigpais", Left(sigla, 2))
        label.SetNamedSubStringValue("rangecode", Mid(tempFirtCode, 4, 7))
        'rellenamos los 10 digitos de los chips de las jeringas.
        For i As Integer = 1 To 10
            label.SetNamedSubStringValue("code" & i, Right(tempLastCode, 5))
            label.SetNamedSubStringValue("barcode" & i & "cod", codechip)
            codechip += 1
            tempLastCode = codechip.ToString("000000000000000")
        Next
        tempFirtCode = tempLastCode
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printReport39()
        'impresion de etiqueta report39 de lote tamaño 55mm x 40mm
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\lote\label39_lb.btw")
        label.SetNamedSubStringValue("barcode", numLote)
        label.SetNamedSubStringValue("numlote", numLote)
        label.SetNamedSubStringValue("numarticulo", refArticulo)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("numorden", numPedido)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        objbt.Visible = True
        'label.PrintOut(False, False)
        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
#End Region
#Region "Etiquetas Cajas"
    Public Sub printCajaStandart()
        'impresión de etiqueta standart de Caja tamaño 60mm x 47mm
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\caja\standart_bb.btw")
        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)

    End Sub
    Public Sub printCajaTipo1()
        'impresión de etiqueta Tipo1 de Caja tamaño 60mm x 47mm
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\caja\tipo3_bb.btw")
        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printCajaTipo3()
        'impresión de etiqueta Tipo3 de Caja tamaño 60mm x 47mm
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\caja\tipo3_bb.btw")
        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printCajaTipo5()
        'impresión de etiqueta Tipo5 de Caja tamaño 60mm x 47mm
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\caja\tipo5_bb.btw")
        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printCajaTipo11()
        'impresión de etiqueta Tipo11 de Caja tamaño 60mm x 47mm
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\caja\tipo11_bb.btw")
        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        label.SetNamedSubStringValue("fechacaduc", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printCajaTipo12()
        'impresion de etiqueta Tipo12 de Caja Tamaño 80x80mm
        Dim material, peso, regZooSan As String
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\caja\spagna_0700_0500_011.btw")
        If descArticulo = "T-IT 8100" Or descArticulo = "T-IS 8010" Then
            material = "vidrio"
            peso = "0,114 grs"
            regZooSan = "02694-MUZ"
        ElseIf descArticulo = "T-SL 8010" Then
            material = "polimero"
            peso = "0,044 grs"
            regZooSan = "02995-MUZ"
        Else
            material = ""
            peso = ""
            regZooSan = ""
        End If
        label.SetNamedSubStringValue("material", material)
        label.SetNamedSubStringValue("tipoproducto", descArticulo)
        label.SetNamedSubStringValue("peso", peso)
        label.SetNamedSubStringValue("feccaduc", fecCaducidad)
        label.SetNamedSubStringValue("regzoosan", regZooSan)
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("fecfabrica", fecFabricacion)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        objbt.Visible = True
        'label.PrintOut(False, False)
        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printCajaTipo15()
        'impresión de etiqueta Tipo15 de Caja tamaño 60mm x 47mm
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\caja\tipo15_bb.btw")
        label.SetNamedSubStringValue("sigproducto", Left(descArticulo, 4))
        label.SetNamedSubStringValue("numcaja", numCaja)
        label.SetNamedSubStringValue("siglapais", sigla)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printCajaReport36(reporte As Integer)
        'impresión de etiquetas Report36 de Caja tamaño 55mm x 40mm
        Dim lastcode As Long
        Dim texto As String
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\caja\label36_bb.btw")
        lastcode = (Long.Parse(tempFirtCode) + ((jerxLote * lotxCaja)))
        tempLastCode = lastcode.ToString("000000000000000")
        If reporte = 0 Then
            'texto descripcion report36
            texto = "SYRINGE VALUE IN BOX OF 100"
        ElseIf reporte = 1 Then
            'texto descripcion report42
            texto = "STANDART IN BOX OF  " & (jerxLote * lotxCaja)
        Else
            'texto descripcion report45
            texto = "NEEDLE KIT IN BOX OF 100"
        End If
        label.SetNamedSubStringValue("barcode", numCaja)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("texto", texto)
        label.SetNamedSubStringValue("firstcode", tempFirtCode)
        label.SetNamedSubStringValue("lastcode", (lastcode - 1))
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("boxnum", Right(numCaja, 3))
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        tempFirtCode = tempLastCode
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printCajaReport38()
        'impresión de etiquetas Report38 de Caja tamaño 75mm x 100mm
        Dim lastcode As Long
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\caja\label38_bb.btw")
        lastcode = (Long.Parse(tempFirtCode)) + ((jerxLote * lotxCaja))
        tempLastCode = lastcode.ToString("000000000000000")
        label.SetNamedSubStringValue("cajaactual", Right(numCaja, 3))
        label.SetNamedSubStringValue("cajatotal", cantCaja.ToString("000"))
        label.SetNamedSubStringValue("cantpiezas", (jerxLote * lotxCaja))
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("barcodenumlote", numCaja)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("fecactual", Format(Date.Now, "yyy-MM-dd"))
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
        'label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
    Public Sub printCajaReport40()
        'impresion de etiquetas Report40 de Caja tamañon 55mm x 40mm
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\caja\label40_bb.btw")
        label.SetNamedSubStringValue("barcode", numCaja)
        label.SetNamedSubStringValue("numlote", numCaja)
        label.SetNamedSubStringValue("numarticulo", refArticulo)
        label.SetNamedSubStringValue("description", descArticulo)
        label.SetNamedSubStringValue("numorden", numPedido)
        label.SetNamedSubStringValue("feccadu", fecCaducidad)
        'se abre el programa Bartender y vemos como queda la etiqueta con los datos reales. QUITAR en programa definitivo.
        'objbt.Visible = True
        label.PrintOut(False, False)
        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
#End Region
    Public Sub printEtiPalet(pedsap As String, npalet As Integer, cajaini As String, cajafin As String)
        Dim label As BarTender.Format = objbt.Formats.Open("c:\Users\Ezequiel\Documents\etiquetas_datamars\etiqueta_palet.btw")
        label.SetNamedSubStringValue("numsap", pedsap)
        label.SetNamedSubStringValue("referencia", refArticulo)
        label.SetNamedSubStringValue("numpalet", npalet)
        label.SetNamedSubStringValue("feccaduc", fecCaducidad)
        label.SetNamedSubStringValue("numcajaIni", cajaini)
        label.SetNamedSubStringValue("numCajaFin", cajafin)
        label.PrintOut(False, False)
        label.Close(BarTender.BtSaveOptions.btDoNotSaveChanges)
    End Sub
End Module
