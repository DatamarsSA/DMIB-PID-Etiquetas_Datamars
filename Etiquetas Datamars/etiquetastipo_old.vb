Module etiquetastipo
    '    Public puerto As String = "LPT1"
    '#Region "Etiquetas Cajas Pequeñas"
    '    Public Function PrintAstStandard(sBatch, sSigla, sData)
    '        Dim byGaranzia As Byte
    '        'Dim intOutputFile As Integer
    '        Dim sResetPrint, sLine, sScadenza, producto As String
    '        producto = "T-SL 8010"
    '        byGaranzia = 3
    '        'sScadenza = " EXP." & Right(CStr(Month(Date) + 100), 2) & "/" & CStr(Right(Format(Date, "dd.mm.yyyy"), 4) + byGaranzia)
    '        sScadenza = sData '" EXP." & Left(sData, 4) + byGaranzia & "-" & Right(sData, 2)

    '        'intOutputFile = FreeFile()

    '        sResetPrint = Chr(27) & Chr(64)                         'Inizializza la stampante
    '        sResetPrint = sResetPrint & Chr(27) & Chr(85) & Chr(1)  'Stampa unidirezionale
    '        sResetPrint = sResetPrint & Chr(27) & Chr(120) & Chr(1)  'Letter Quality
    '        sResetPrint = sResetPrint & Chr(27) & Chr(67) & Chr(11)  '12 linee per pagina
    '        sResetPrint = sResetPrint & Chr(27) & Chr(77)            '12 cpi
    '        sResetPrint = sResetPrint & Chr(27) & Chr(74) & Chr(125)   'Avanza di 125/180 di pollice             '

    '        sLine = Left(producto, 4) & " " & sBatch & " - " & sSigla & Chr(13) & Chr(10) & sScadenza
    '        escribe_puerto(sResetPrint & sLine, puerto)
    '        escribe_puerto(Chr(12), puerto)
    '        '    Open "LPT1:" For Output As #intOutputFile
    '        'Print #intOutputFile, sResetPrint & sLine
    '        'Print #intOutputFile, Chr(12)
    '        'Close #intOutputFile
    '        Return True
    '    End Function
    '    Public Function PrintAstTipo3Utility(sBatch, sData)
    '        Dim intVelAvanzEtichetta, intTempTestina,
    '        intAltezzaEtichetta, intDistTraEtichette, intLarghezzaEtichetta,
    '        intPosXEtichetta, intPosYEtichetta,
    '        intOffsetXRiga1, intOffsetYRiga1,
    '        intOffsetXRiga2, intOffsetYRiga2,
    '        intOffsetXRiga3, intOffsetYRiga3,
    '        intOffsetXRiga4, intOffsetYRiga4,
    '        intOffsetXBarcode, intOffsetYBarcode,
    '        intAltezzaBarcode,
    '        intBoldRiga1, intBoldRiga2, intBoldRiga3, intBoldRiga4,
    '        intFontSizeRiga1, intFontSizeRiga2, intFontSizeRiga3, intFontSizeRiga4,
    '        intDistTaglio, intAnniGaranzia As Integer

    '        Dim sTemp, sNumeroLotto,
    '       sTestoBarcode, sNumeroDAgrement,
    '       sNumeroLottoRiga2, sScadenzaRiga3 As String

    '        'Set Form2 = Form_F_UTILITY_STAMPA_ETICH_ASTUCCIO'
    '        If Not Form1.SerialPort1.IsOpen Then
    '            Form1.SerialPort1.Open()
    '        End If
    '        sTemp = ""
    '        sNumeroDAgrement = "Numéro d'agrément 96"
    '        sNumeroLottoRiga2 = "LOT " & sBatch
    '        intAnniGaranzia = 3
    '        'sScadenzaRiga3 = "EXP. 01/" & Right(CStr(Month(Date) + 100), 2) & "/" & CStr(Right(Format(Date, "dd.mm.yyyy"), 4) + intAnniGaranzia)
    '        sScadenzaRiga3 = sData '"EXP. " & Right(sData, 4) + intAnniGaranzia & "-" & Left(sData, 2)

    '        'Inizializzazione stampante
    '        intVelAvanzEtichetta = 10
    '        intTempTestina = -5
    '        intAltezzaEtichetta = 50
    '        intLarghezzaEtichetta = 40
    '        intDistTraEtichette = 54 'Distanza tra l'inizio di una etichetta e quello della successiva, OK
    '        intPosXEtichetta = 1    'OK
    '        intPosYEtichetta = 1    'OK
    '        intOffsetXRiga1 = 3
    '        intOffsetYRiga1 = 41
    '        intOffsetXRiga2 = 19
    '        intOffsetYRiga2 = 8
    '        intOffsetXRiga3 = 28
    '        intOffsetYRiga3 = 16
    '        intOffsetXRiga4 = 35
    '        intOffsetYRiga4 = 8
    '        intOffsetXBarcode = 4   'OK
    '        intOffsetYBarcode = 17
    '        intAltezzaBarcode = 10  'OK

    '        sNumeroLotto = sBatch 'OK
    '        intDistTaglio = 5
    '        intFontSizeRiga1 = 25
    '        intFontSizeRiga2 = 47
    '        intFontSizeRiga3 = 47
    '        intFontSizeRiga4 = 40
    '        intBoldRiga1 = 3
    '        intBoldRiga2 = 3
    '        intBoldRiga3 = 3
    '        intBoldRiga4 = 3

    '        sTestoBarcode = sNumeroLotto  'OK

    '        'Stampa etichetta
    '        'Form2.msComm5.Output = "m m" & Chr(13)                                           'Unità di misura mm
    '        Form1.SerialPort1.Write("m m" & Chr(13))
    '        'Form2.msComm5.Output = "J Inizio_Stampa" & Chr(13)                          'Inizio Job di stampa
    '        Form1.SerialPort1.Write("J Inizio_Stampa" & Chr(13))
    '        sTemp = "H " & intVelAvanzEtichetta                                                          'Velocità avanzamento etichetta
    '        sTemp = sTemp & ", " & intTempTestina                                                     'Temperatura testina
    '        'Form2.msComm5.Output = sTemp & Chr(13)
    '        Form1.SerialPort1.Write(sTemp & Chr(13))
    '        ' Form2.msComm5.Output = "O R, U" & Chr(13)                                   'Orientamento e Rotazione etichetta
    '        Form1.SerialPort1.Write("O R, U" & Chr(13))
    '        sTemp = "S l1; 0.00, 0.00, "                             'Senza marker dietro le etichette; Posizione Left e Top etichetta
    '        sTemp = sTemp & intAltezzaEtichetta & ", "                                           'Altezza etichetta
    '        sTemp = sTemp & intDistTraEtichette & ", "                                           'Distanza tra le etichette
    '        sTemp = sTemp & intLarghezzaEtichetta                                            'Larghezza etichetta
    '        'Form2.msComm5.Output = sTemp & Chr(13)
    '        Form1.SerialPort1.Write(sTemp & Chr(13))
    '        'BARCODE 1
    '        sTemp = "B " & (intOffsetXBarcode + intPosXEtichetta) & ", "                                   'Posizione x BARCODE 1
    '        sTemp = sTemp & (intDistTraEtichette - intPosYEtichetta - 8) & ", "   'Posizione y  BARCODE 1
    '        sTemp = sTemp & "90, e, "                                            'Rotazione, e = EAN128 senza descrizione, BARCODE 1
    '        sTemp = sTemp & intAltezzaBarcode & ", 0.23;" & sTestoBarcode                  'Altezza, narrow, BARCODE 1
    '        Form1.SerialPort1.Write(sTemp & Chr(13))
    '        'Form2.msComm5.Output = sTemp & Chr(13)
    '        'Riga2 LOTTO
    '        sTemp = "T " & (intOffsetXRiga2 + intPosXEtichetta) & ", "                              'Posizione x Riga 2
    '        sTemp = sTemp & (intDistTraEtichette - intPosYEtichetta - intOffsetYRiga2) & ", "    'Posizione y Riga 2
    '        sTemp = sTemp & "90, " & intBoldRiga2 & ", " & (intFontSizeRiga2 / 10) & "; "        'Rotazione 0, Bold, Font Riga 2
    '        sTemp = sTemp & sNumeroLottoRiga2
    '        'Form2.msComm5.Output = sTemp & Chr(13)
    '        Form1.SerialPort1.Write(sTemp & Chr(13))
    '        'Riga3 EXP Date
    '        sTemp = "T " & (intOffsetXRiga3 + intPosXEtichetta) & ", "                              'Posizione x Riga 3
    '        sTemp = sTemp & (intDistTraEtichette - intPosYEtichetta - intOffsetYRiga3) & ", "    'Posizione y Riga 3
    '        sTemp = sTemp & "90, " & intBoldRiga3 & ", " & (intFontSizeRiga3 / 10) & "; "        'Rotazione 0, Bold, Font Riga 3
    '        sTemp = sTemp & sScadenzaRiga3
    '        'Form2.msComm5.Output = sTemp & Chr(13)
    '        'Riga4 Numéro d'agrément 96
    '        sTemp = "T " & (intOffsetXRiga4 + intPosXEtichetta) & ", "                              'Posizione x Riga 4
    '        sTemp = sTemp & (intDistTraEtichette - intPosYEtichetta - intOffsetYRiga4) & ", "    'Posizione y Riga 4
    '        sTemp = sTemp & "90, " & intBoldRiga4 & ", " & (intFontSizeRiga4 / 10) & "; "        'Rotazione 0, Bold, Font Riga 4
    '        sTemp = sTemp & sNumeroDAgrement
    '        'Form2.msComm5.Output = sTemp & Chr(13)
    '        Form1.SerialPort1.Write(sTemp & Chr(13))
    '        'Taglio etichetta
    '        'Form2.msComm5.Output = "C 1, " & intDistTaglio & Chr(13)            'Taglio dopo 1 etichetta e Distanza
    '        Form1.SerialPort1.Write("C 1, " & intDistTaglio & Chr(13))
    '        'Form2.msComm5.Output = "C e" & Chr(13)                                    'Taglio alla fine del Job
    '        Form1.SerialPort1.Write("C e" & Chr(13))
    '        'Stampa 1 etichetta
    '        'Form2.msComm5.Output = "A 1" & Chr(13)
    '        Form1.SerialPort1.Write("A 1" & Chr(13))
    '        Return True
    '    End Function


    '    Public Function PrintAstTipo5(sBatch, sSigla, sData)
    '        Dim byGaranzia As Byte
    '        'Dim intOutputFile As Integer
    '        Dim sResetPrint, sLine, sScadenza, producto As String

    '        byGaranzia = 3
    '        producto = "T-SL 8010"
    '        'sScadenza = " EXP." & Right(CStr(Month(Date) + 100), 2) & "/" & CStr(Right(Format(Date, "dd.mm.yyyy"), 4) + byGaranzia)
    '        sScadenza = sData '" EXP." & Right(sData, 4) + byGaranzia & "-" & Left(sData, 2)

    '        'intOutputFile = FreeFile()

    '        sResetPrint = Chr(27) & Chr(64)                         'Inizializza la stampante
    '        sResetPrint = sResetPrint & Chr(27) & Chr(85) & Chr(1)  'Stampa unidirezionale
    '        sResetPrint = sResetPrint & Chr(27) & Chr(120) & Chr(1) 'Letter Quality
    '        sResetPrint = sResetPrint & Chr(27) & Chr(67) & Chr(11) '11 linee per pagina
    '        sResetPrint = sResetPrint & Chr(27) & Chr(77)           '12 cpi
    '        sResetPrint = sResetPrint & Chr(27) & Chr(74) & Chr(25) 'Avanza di 40/180 di pollice             '

    '        sLine = Left(producto, 4) & " " & sBatch & " - " & sSigla & Chr(13) & Chr(10) & sScadenza
    '        escribe_puerto(sResetPrint & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(35) & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(30) & sLine, puerto)
    '        escribe_puerto(Chr(12), puerto)
    '        'Open "LPT1:" For Output As #intOutputFile
    '        'Print #intOutputFile, sResetPrint & sLine
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(35) & sLine  'Avanza di 40/180 di pollice
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(30) & sLine  'Avanza di 30/180 di pollice
    '        'Print #intOutputFile, Chr(12)
    '        'Close #intOutputFile
    '        Return True
    '    End Function
    '    Public Function PrintAstTipo15(sBatch, sSigla, sData)
    '        Dim byGaranzia As Byte
    '        'Dim intOutputFile As Integer
    '        Dim sResetPrint, sLine, sScadenza, producto As String
    '        producto = "T-SL 8010"
    '        byGaranzia = 3
    '        'sScadenza = " EXP." & Right(CStr(Month(Date) + 100), 2) & "/" & CStr(Right(Format(Date, "dd.mm.yyyy"), 4) + byGaranzia)
    '        sScadenza = sData '" EXP." & Right(sData, 4) + byGaranzia & "-" & Left(sData, 2)

    '        'intOutputFile = FreeFile()

    '        sResetPrint = Chr(27) & Chr(64)                         'Inizializza la stampante
    '        sResetPrint = sResetPrint & Chr(27) & Chr(85) & Chr(1)  'Stampa unidirezionale
    '        sResetPrint = sResetPrint & Chr(27) & Chr(120) & Chr(1) 'Letter Quality
    '        sResetPrint = sResetPrint & Chr(27) & Chr(67) & Chr(11) '11 linee per pagina
    '        sResetPrint = sResetPrint & Chr(27) & Chr(77)           '12 cpi
    '        sResetPrint = sResetPrint & Chr(27) & Chr(74) & Chr(5) 'Avanza di 40/180 di pollice

    '        sLine = Left(producto, 4) & " " & sBatch & " - " & sSigla & Chr(13) & Chr(10) & sScadenza & "     DATAMARS"
    '        escribe_puerto(sResetPrint & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(5), puerto)
    '        escribe_puerto(sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(5), puerto)
    '        escribe_puerto(Chr(12), puerto)

    '        ' Open "LPT1:" For Output As #intOutputFile
    '        'Print #intOutputFile, sResetPrint & sLine
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(5)                'Avanza di 40/180 di pollice
    '        'Print #intOutputFile, sLine
    '        ' Print #intOutputFile, Chr(27) & Chr(74) & Chr(5)                'Avanza di 30/180 di pollice
    '        'Print #intOutputFile, sLine
    '        ' Print #intOutputFile, Chr(12)
    '        'Close #intOutputFile
    '        Return True
    '    End Function
    '#End Region
    '#Region "Etiquetas Cajas Grandes"

    '    Public Function PrintCartStandard(sBatch, sSigla)
    '        Dim intOutputFile As Integer
    '        Dim sResetPrint, sLine, producto As String
    '        producto = "T-SL 8010"
    '        intOutputFile = FreeFile()

    '        sResetPrint = Chr(27) & Chr(64)                         'Inizializza la stampante
    '        sResetPrint = sResetPrint & Chr(27) & Chr(85) & Chr(1)  'Stampa unidirezionale
    '        sResetPrint = sResetPrint & Chr(27) & Chr(120) & Chr(1) 'Letter Quality
    '        sResetPrint = sResetPrint & Chr(27) & Chr(67) & Chr(11) '11 linee per pagina
    '        sResetPrint = sResetPrint & Chr(27) & Chr(77)           '12 cpi
    '        sResetPrint = sResetPrint & Chr(27) & Chr(74) & Chr(40) 'Avanza di 40/180 di pollice

    '        sLine = Left(producto, 4) & " " & sBatch & " - " & sSigla
    '        escribe_puerto(sResetPrint & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(40), puerto)
    '        escribe_puerto(sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(30), puerto)
    '        escribe_puerto(sLine, puerto)
    '        escribe_puerto(Chr(12), puerto)
    '        ' Open "LPT1:" For Output As #intOutputFile
    '        'Print #intOutputFile, sResetPrint & sLine
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(40)                'Avanza di 40/180 di pollice
    '        'Print #intOutputFile, sLine
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(30)                'Avanza di 30/180 di pollice
    '        'Print #intOutputFile, sLine
    '        'Print #intOutputFile, Chr(12)
    '        'Close #intOutputFile
    '        Return True
    '    End Function
    '    Public Function PrintCartTipo1(sBatch, sSigla, sData)
    '        ' Dim intOutputFile As Integer
    '        Dim byGaranzia As Byte
    '        Dim sResetPrint, sLine, sScadenza, producto As String
    '        producto = "T-SL 8010"
    '        byGaranzia = 3
    '        'sScadenza = Right(CStr(Month(Date) + 100), 2) & "/" & CStr(Right(Format(Date, "dd.mm.yyyy"), 4) + byGaranzia)

    '        sScadenza = Right(sData, 7)

    '        'intOutputFile = FreeFile()

    '        sResetPrint = Chr(27) & Chr(64)                                     'Inizializza la stampante
    '        sResetPrint = sResetPrint & Chr(27) & Chr(85) & Chr(1)  'Stampa unidirezionale
    '        sResetPrint = sResetPrint & Chr(27) & Chr(120) & Chr(1)  'Letter Quality
    '        sResetPrint = sResetPrint & Chr(27) & Chr(67) & Chr(11)  '11 linee per pagina
    '        sResetPrint = sResetPrint & Chr(27) & Chr(48)                   'Interlinea 1/8"
    '        sResetPrint = sResetPrint & Chr(15)                                 'Stampa compressa
    '        sResetPrint = sResetPrint & Chr(27) & Chr(74) & Chr(15)   'Avanza di 15/180 di pollice

    '        sLine = "Num‚ro d'agr‚ment 96" & Chr(13) & Chr(10) & Left(producto, 4) & " " & sBatch & " - " & sSigla &
    '                Chr(13) & Chr(10) & "EXP.: " & sScadenza

    '        escribe_puerto(sResetPrint & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(30) & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(20) & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(50) & Chr(12), puerto)
    '        ' Open "LPT1:" For Output As #intOutputFile
    '        'Print #intOutputFile, sResetPrint & sLine
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(30) & sLine                'Avanza di 30/180 di pollice e stampa
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(20) & sLine               'Avanza di 20/180 di pollice e stampa
    '        'Print #intOutputFile, Chr(27) & Chr(50) & Chr(12)
    '        'Close #intOutputFile
    '        Return True
    '    End Function
    '    Public Function PrintCartTipo3(sBatch, sSigla, sData)
    '        Dim intOutputFile As Integer
    '        Dim byGaranzia As Byte
    '        Dim sResetPrint, sLine, sScadenza, producto As String
    '        producto = "T-SL 8010"
    '        byGaranzia = 3
    '        'sScadenza = "EXP.: " & Right(CStr(Month(Date) + 100), 2) & "/" & CStr(Right(Format(Date, "dd.mm.yyyy"), 4) + byGaranzia)
    '        sScadenza = sData '"EXP.:" & Right(sData, 4) + byGaranzia & "-" & Left(sData, 2)

    '        intOutputFile = FreeFile()

    '        sResetPrint = Chr(27) & Chr(64)                                     'Inizializza la stampante
    '        sResetPrint = sResetPrint & Chr(27) & Chr(85) & Chr(1)  'Stampa unidirezionale
    '        sResetPrint = sResetPrint & Chr(27) & Chr(120) & Chr(1)  'Letter Quality
    '        sResetPrint = sResetPrint & Chr(27) & Chr(67) & Chr(11)  '11 linee per pagina
    '        sResetPrint = sResetPrint & Chr(27) & Chr(77)                  '12 cpi
    '        sResetPrint = sResetPrint & Chr(27) & Chr(74) & Chr(25)   'Avanza di 25/180 di pollice

    '        sLine = Left(producto, 4) & " " & sBatch & " - " & sSigla & Chr(13) & Chr(10) & sScadenza

    '        escribe_puerto(sResetPrint & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(35) & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(30) & sLine, puerto)
    '        escribe_puerto(Chr(12), puerto)
    '        ' Open "LPT1:" For Output As #intOutputFile
    '        'Print #intOutputFile, sResetPrint & sLine
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(35) & sLine                'Avanza di 35/180 di pollice e stampa
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(30) & sLine               'Avanza di 30/180 di pollice e stampa
    '        'Print #intOutputFile, Chr(12)
    '        'Close #intOutputFile
    '        Return True
    '    End Function


    '    Public Function PrintCartTipo5(sBatch, sSigla)
    '        Dim intOutputFile As Integer
    '        Dim sResetPrint, sLine, producto As String
    '        producto = "T-SL 8010"
    '        intOutputFile = FreeFile()

    '        sResetPrint = Chr(27) & Chr(64)                             'Inizializza la stampante
    '        sResetPrint = sResetPrint & Chr(27) & Chr(85) & Chr(1)      'Stampa unidirezionale
    '        sResetPrint = sResetPrint & Chr(27) & Chr(120) & Chr(1)     'Letter Quality
    '        sResetPrint = sResetPrint & Chr(27) & Chr(67) & Chr(11)     '11 linee per pagina
    '        sResetPrint = sResetPrint & Chr(27) & Chr(77)               '12 cpi
    '        sResetPrint = sResetPrint & Chr(27) & Chr(74) & Chr(40)     'Avanza di 40/180 di pollice

    '        sLine = Left(producto, 4) & " " & sBatch & " - " & sSigla
    '        escribe_puerto(sResetPrint & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(40), puerto)
    '        escribe_puerto(sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(30), puerto)
    '        escribe_puerto(sLine, puerto)
    '        escribe_puerto(Chr(12), puerto)
    '        'Open "LPT1:" For Output As #intOutputFile
    '        'Print #intOutputFile, sResetPrint & sLine
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(40)            'Avanza di 40/180 di pollice
    '        'Print #intOutputFile, sLine
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(30)            'Avanza di 30/180 di pollice
    '        'Print #intOutputFile, sLine
    '        'Print #intOutputFile, Chr(12)
    '        'Close #intOutputFile
    '        Return True
    '    End Function
    '    Public Function PrintCartTipo11(sBatch, sSigla, sData)
    '        'Dim intOutputFile As Integer
    '        Dim sResetPrint, sLine(0 To 12), sScadenza, producto As String
    '        Dim byGaranzia, cnt As Byte

    '        byGaranzia = 3
    '        sScadenza = sData '"EXP.:" & Right(sData, 4) + byGaranzia & "-" & Left(sData, 2)
    '        producto = "T-SL8010"
    '        'intOutputFile = FreeFile()

    '        sResetPrint = Chr(27) & Chr(64)                                 'Inizializza la stampante
    '        sResetPrint = sResetPrint & Chr(27) & Chr(85) & Chr(1)          'Stampa unidirezionale
    '        sResetPrint = sResetPrint & Chr(27) & Chr(120) & Chr(1)         'Letter Quality
    '        sResetPrint = sResetPrint & Chr(27) & Chr(50)                   'Set line spacing to 1/6
    '        sResetPrint = sResetPrint & Chr(27) & Chr(77)                   '12 cpi
    '        sResetPrint = sResetPrint & Chr(27) & Chr(67) & Chr(12)         'page lenght = line spacing * linee = 1/6 * 12 = 2 inches
    '        sResetPrint = sResetPrint & Chr(27) & Chr(40) & Chr(116) &
    '                      Chr(3) & Chr(0) & Chr(3) & Chr(14) & Chr(0)       'assign char table 3
    '        sResetPrint = sResetPrint & Chr(27) & Chr(116) & Chr(3)         'select char table 3


    '        sLine(1) = "ISO Standart 11784/11785"
    '        sLine(2) = Chr(140) & Chr(105) & Chr(170) & Chr(224) & Chr(174) &
    '                  Chr(231) & Chr(105) & Chr(175) & " " & Chr(167) & " " & Chr(160) &
    '                  Chr(175) & Chr(171) & Chr(105) & Chr(170) & Chr(160) & Chr(226) &
    '                  Chr(174) & Chr(224) & Chr(174) & Chr(172)
    '        sLine(3) = "T-IS 8100"
    '        sLine(5) = Chr(140) & Chr(168) & Chr(170) & Chr(224) & Chr(174) &
    '                  Chr(231) & Chr(168) & Chr(175) & " " & Chr(225) & " " & Chr(160) &
    '                  Chr(175) & Chr(175) & Chr(171) & Chr(168) & Chr(170) & Chr(160) &
    '                  Chr(226) & Chr(174) & Chr(224) & Chr(174) & Chr(172)
    '        sLine(6) = "T-IS 8100"
    '        sLine(8) = Left(producto, 4) & " " & sBatch & " - " & sSigla
    '        sLine(9) = sScadenza

    '        'Open "LPT1:" For Output As #intOutputFile
    '        'Print #intOutputFile, sResetPrint;        'Invia la configurazione (; = non aggiunge CRLF alla fine)
    '        escribe_puerto(sResetPrint & vbCrLf, puerto)
    '        For cnt = 1 To 12
    '            'Print #intOutputFile, sLine(cnt)
    '            If sLine(cnt) <> "" Then
    '                escribe_puerto(sLine(cnt), puerto)
    '            End If
    '        Next
    '        'Print #intOutputFile, Chr(27) & Chr(116) & Chr(0);               'select char table 0
    '        escribe_puerto(Chr(27) & Chr(116) & Chr(0) & vbCrLf, puerto)
    '        'Close #intOutputFile
    '        Return True
    '    End Function
    '    Public Function PrintCartTipo15(sBatch, sSigla)
    '        'Dim intOutputFile As Integer
    '        Dim sResetPrint, sLine, producto As String
    '        producto = "T-SL 8010"
    '        'intOutputFile = FreeFile()

    '        sResetPrint = Chr(27) & Chr(64)                         'Inizializza la stampante
    '        sResetPrint = sResetPrint & Chr(27) & Chr(85) & Chr(1)  'Stampa unidirezionale
    '        sResetPrint = sResetPrint & Chr(27) & Chr(120) & Chr(1) 'Letter Quality
    '        sResetPrint = sResetPrint & Chr(27) & Chr(67) & Chr(11) '11 linee per pagina
    '        sResetPrint = sResetPrint & Chr(27) & Chr(77)           '12 cpi
    '        sResetPrint = sResetPrint & Chr(27) & Chr(74) & Chr(5) 'Avanza di 40/180 di pollice

    '        sLine = Left(producto, 4) & " " & sBatch & " - " & sSigla & Chr(13) & Chr(10) & " DATAMARS"
    '        escribe_puerto(sResetPrint & sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(5), puerto)
    '        escribe_puerto(sLine, puerto)
    '        escribe_puerto(Chr(27) & Chr(74) & Chr(5), puerto)
    '        escribe_puerto(sLine, puerto)
    '        escribe_puerto(Chr(12), puerto)
    '        'Open "LPT1:" For Output As #intOutputFile
    '        'Print #intOutputFile, sResetPrint & sLine
    '        'Print #intOutputFile, Chr(27) & Chr(74) & Chr(5)                'Avanza di 40/180 di pollice
    '        'Print #intOutputFile, sLine
    '        ' Print #intOutputFile, Chr(27) & Chr(74) & Chr(5)                'Avanza di 30/180 di pollice
    '        'Print #intOutputFile, sLine
    '        'Print #intOutputFile, Chr(12)
    '        'Close #intOutputFile
    '        Return True
    '    End Function


    '#End Region

End Module
