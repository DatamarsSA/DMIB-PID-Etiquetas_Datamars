<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cmdImprimirLotes = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmdImprimirCajas = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSemana = New System.Windows.Forms.NumericUpDown()
        Me.txtAño = New System.Windows.Forms.NumericUpDown()
        Me.txtCaja = New System.Windows.Forms.NumericUpDown()
        Me.txtLote = New System.Windows.Forms.NumericUpDown()
        Me.numLotesPorCaja = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.numAñoCad = New System.Windows.Forms.NumericUpDown()
        Me.numMesCad = New System.Windows.Forms.NumericUpDown()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.numLotes = New System.Windows.Forms.NumericUpDown()
        Me.numCajas = New System.Windows.Forms.NumericUpDown()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtCodPedido = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtGlobalFirstCode = New System.Windows.Forms.TextBox()
        Me.numUnidadesPorlote = New System.Windows.Forms.NumericUpDown()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtLastCodigoChip = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtNumPedido = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtNombreCliente = New System.Windows.Forms.TextBox()
        Me.txtSiglaPais = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtNombProducto = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.txtRefProducto = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.lbEtiLoAl = New System.Windows.Forms.Label()
        Me.lbEtiLoAn = New System.Windows.Forms.Label()
        Me.lbEtiCaAn = New System.Windows.Forms.Label()
        Me.lbEtiCaAl = New System.Windows.Forms.Label()
        Me.cmdImpManual = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.picEtiquetaCaja = New System.Windows.Forms.PictureBox()
        Me.picEtiquetaLote = New System.Windows.Forms.PictureBox()
        Me.numCajasPorPalet = New System.Windows.Forms.NumericUpDown()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.numPalets = New System.Windows.Forms.NumericUpDown()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.cmdImprimirPalets = New System.Windows.Forms.Button()
        Me.txtnumsap = New System.Windows.Forms.TextBox()
        CType(Me.txtSemana, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAño, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCaja, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLote, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numLotesPorCaja, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numAñoCad, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numMesCad, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numLotes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numCajas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numUnidadesPorlote, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picEtiquetaCaja, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picEtiquetaLote, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numCajasPorPalet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numPalets, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdImprimirLotes
        '
        Me.cmdImprimirLotes.BackColor = System.Drawing.Color.Transparent
        Me.cmdImprimirLotes.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.cmdImprimirLotes.ForeColor = System.Drawing.Color.Maroon
        Me.cmdImprimirLotes.Location = New System.Drawing.Point(263, 409)
        Me.cmdImprimirLotes.Name = "cmdImprimirLotes"
        Me.cmdImprimirLotes.Size = New System.Drawing.Size(100, 23)
        Me.cmdImprimirLotes.TabIndex = 21
        Me.cmdImprimirLotes.Text = "Lotes"
        Me.cmdImprimirLotes.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Maroon
        Me.Label1.Location = New System.Drawing.Point(12, 248)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 13)
        Me.Label1.TabIndex = 34
        Me.Label1.Text = "Lote inicial"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.Maroon
        Me.Label2.Location = New System.Drawing.Point(12, 222)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "Caja inicial"
        '
        'cmdImprimirCajas
        '
        Me.cmdImprimirCajas.ForeColor = System.Drawing.Color.Maroon
        Me.cmdImprimirCajas.Location = New System.Drawing.Point(263, 437)
        Me.cmdImprimirCajas.Name = "cmdImprimirCajas"
        Me.cmdImprimirCajas.Size = New System.Drawing.Size(100, 23)
        Me.cmdImprimirCajas.TabIndex = 22
        Me.cmdImprimirCajas.Text = "Cajas"
        Me.cmdImprimirCajas.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Maroon
        Me.Label3.Location = New System.Drawing.Point(12, 196)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(26, 13)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Año"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Maroon
        Me.Label4.Location = New System.Drawing.Point(12, 170)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 31
        Me.Label4.Text = "Semana"
        '
        'txtSemana
        '
        Me.txtSemana.Enabled = False
        Me.txtSemana.Location = New System.Drawing.Point(133, 168)
        Me.txtSemana.Maximum = New Decimal(New Integer() {53, 0, 0, 0})
        Me.txtSemana.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.txtSemana.Name = "txtSemana"
        Me.txtSemana.Size = New System.Drawing.Size(120, 20)
        Me.txtSemana.TabIndex = 9
        Me.txtSemana.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtSemana.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'txtAño
        '
        Me.txtAño.Enabled = False
        Me.txtAño.Location = New System.Drawing.Point(133, 194)
        Me.txtAño.Name = "txtAño"
        Me.txtAño.Size = New System.Drawing.Size(120, 20)
        Me.txtAño.TabIndex = 10
        Me.txtAño.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCaja
        '
        Me.txtCaja.Enabled = False
        Me.txtCaja.Location = New System.Drawing.Point(133, 220)
        Me.txtCaja.Maximum = New Decimal(New Integer() {9999, 0, 0, 0})
        Me.txtCaja.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.txtCaja.Name = "txtCaja"
        Me.txtCaja.Size = New System.Drawing.Size(120, 20)
        Me.txtCaja.TabIndex = 11
        Me.txtCaja.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtCaja.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'txtLote
        '
        Me.txtLote.Enabled = False
        Me.txtLote.Location = New System.Drawing.Point(133, 246)
        Me.txtLote.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.txtLote.Name = "txtLote"
        Me.txtLote.Size = New System.Drawing.Size(120, 20)
        Me.txtLote.TabIndex = 12
        Me.txtLote.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtLote.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'numLotesPorCaja
        '
        Me.numLotesPorCaja.Enabled = False
        Me.numLotesPorCaja.Location = New System.Drawing.Point(133, 299)
        Me.numLotesPorCaja.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numLotesPorCaja.Name = "numLotesPorCaja"
        Me.numLotesPorCaja.Size = New System.Drawing.Size(120, 20)
        Me.numLotesPorCaja.TabIndex = 14
        Me.numLotesPorCaja.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.numLotesPorCaja.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.Color.Maroon
        Me.Label5.Location = New System.Drawing.Point(12, 301)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(74, 13)
        Me.Label5.TabIndex = 36
        Me.Label5.Text = "Lotes por caja"
        '
        'numAñoCad
        '
        Me.numAñoCad.Enabled = False
        Me.numAñoCad.Location = New System.Drawing.Point(133, 381)
        Me.numAñoCad.Name = "numAñoCad"
        Me.numAñoCad.Size = New System.Drawing.Size(120, 20)
        Me.numAñoCad.TabIndex = 16
        Me.numAñoCad.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'numMesCad
        '
        Me.numMesCad.Enabled = False
        Me.numMesCad.Location = New System.Drawing.Point(133, 355)
        Me.numMesCad.Maximum = New Decimal(New Integer() {12, 0, 0, 0})
        Me.numMesCad.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numMesCad.Name = "numMesCad"
        Me.numMesCad.Size = New System.Drawing.Size(120, 20)
        Me.numMesCad.TabIndex = 15
        Me.numMesCad.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.numMesCad.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.Maroon
        Me.Label6.Location = New System.Drawing.Point(12, 357)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 13)
        Me.Label6.TabIndex = 37
        Me.Label6.Text = "Mes caducidad"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.Color.Maroon
        Me.Label7.Location = New System.Drawing.Point(12, 383)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(79, 13)
        Me.Label7.TabIndex = 38
        Me.Label7.Text = "Año caducidad"
        '
        'numLotes
        '
        Me.numLotes.Enabled = False
        Me.numLotes.Location = New System.Drawing.Point(135, 411)
        Me.numLotes.Maximum = New Decimal(New Integer() {99999, 0, 0, 0})
        Me.numLotes.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numLotes.Name = "numLotes"
        Me.numLotes.Size = New System.Drawing.Size(120, 20)
        Me.numLotes.TabIndex = 17
        Me.numLotes.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.numLotes.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'numCajas
        '
        Me.numCajas.Enabled = False
        Me.numCajas.Location = New System.Drawing.Point(135, 437)
        Me.numCajas.Maximum = New Decimal(New Integer() {99999, 0, 0, 0})
        Me.numCajas.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numCajas.Name = "numCajas"
        Me.numCajas.Size = New System.Drawing.Size(120, 20)
        Me.numCajas.TabIndex = 18
        Me.numCajas.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.numCajas.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.Maroon
        Me.Label8.Location = New System.Drawing.Point(6, 413)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 13)
        Me.Label8.TabIndex = 39
        Me.Label8.Text = "Cantidad lotes a imprimir"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.Maroon
        Me.Label9.Location = New System.Drawing.Point(6, 439)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(123, 13)
        Me.Label9.TabIndex = 40
        Me.Label9.Text = "Cantidad cajas a imprimir"
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Items.AddRange(New Object() {"Std", "tipo3", "tipo5", "tipo7", "tipo9", "tipo10", "tipo12", "tipo13", "tipo15", "report35", "report37", "report39", "report43"})
        Me.ListBox1.Location = New System.Drawing.Point(517, 392)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.ScrollAlwaysVisible = True
        Me.ListBox1.Size = New System.Drawing.Size(120, 30)
        Me.ListBox1.TabIndex = 19
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.ForeColor = System.Drawing.Color.Maroon
        Me.Label10.Location = New System.Drawing.Point(527, 376)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(99, 13)
        Me.Label10.TabIndex = 41
        Me.Label10.Text = "Etiquetas Tipo Lote"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.ForeColor = System.Drawing.Color.Maroon
        Me.Label11.Location = New System.Drawing.Point(794, 376)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(102, 13)
        Me.Label11.TabIndex = 42
        Me.Label11.Text = "Etiquetas  Tipo Caja"
        '
        'ListBox2
        '
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.Items.AddRange(New Object() {"Std", "tipo1", "tipo3", "tipo4", "tipo5", "tipo6", "tipo7", "tipo9", "tipo10", "tipo11", "tipo13", "tipo14", "tipo15", "report36", "report38", "report40", "report42", "report44", "report45"})
        Me.ListBox2.Location = New System.Drawing.Point(786, 392)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.ScrollAlwaysVisible = True
        Me.ListBox2.Size = New System.Drawing.Size(120, 30)
        Me.ListBox2.TabIndex = 20
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.ForeColor = System.Drawing.Color.Maroon
        Me.Label12.Location = New System.Drawing.Point(24, 31)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(109, 13)
        Me.Label12.TabIndex = 23
        Me.Label12.Text = "Codigo Barras Pedido"
        '
        'txtCodPedido
        '
        Me.txtCodPedido.Location = New System.Drawing.Point(135, 28)
        Me.txtCodPedido.Name = "txtCodPedido"
        Me.txtCodPedido.Size = New System.Drawing.Size(100, 20)
        Me.txtCodPedido.TabIndex = 1
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.ForeColor = System.Drawing.Color.Maroon
        Me.Label13.Location = New System.Drawing.Point(24, 134)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(92, 13)
        Me.Label13.TabIndex = 29
        Me.Label13.Text = "Inicio Codigo Chip"
        '
        'txtGlobalFirstCode
        '
        Me.txtGlobalFirstCode.Enabled = False
        Me.txtGlobalFirstCode.Location = New System.Drawing.Point(135, 131)
        Me.txtGlobalFirstCode.Name = "txtGlobalFirstCode"
        Me.txtGlobalFirstCode.Size = New System.Drawing.Size(100, 20)
        Me.txtGlobalFirstCode.TabIndex = 7
        '
        'numUnidadesPorlote
        '
        Me.numUnidadesPorlote.Enabled = False
        Me.numUnidadesPorlote.Location = New System.Drawing.Point(133, 272)
        Me.numUnidadesPorlote.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numUnidadesPorlote.Name = "numUnidadesPorlote"
        Me.numUnidadesPorlote.Size = New System.Drawing.Size(120, 20)
        Me.numUnidadesPorlote.TabIndex = 13
        Me.numUnidadesPorlote.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.numUnidadesPorlote.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.ForeColor = System.Drawing.Color.Maroon
        Me.Label14.Location = New System.Drawing.Point(12, 274)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(91, 13)
        Me.Label14.TabIndex = 35
        Me.Label14.Text = "Jeringas por Lote "
        '
        'txtLastCodigoChip
        '
        Me.txtLastCodigoChip.Enabled = False
        Me.txtLastCodigoChip.Location = New System.Drawing.Point(347, 132)
        Me.txtLastCodigoChip.Name = "txtLastCodigoChip"
        Me.txtLastCodigoChip.Size = New System.Drawing.Size(100, 20)
        Me.txtLastCodigoChip.TabIndex = 8
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.ForeColor = System.Drawing.Color.Maroon
        Me.Label15.Location = New System.Drawing.Point(259, 135)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(81, 13)
        Me.Label15.TabIndex = 30
        Me.Label15.Text = "Fin Codigo Chip"
        '
        'txtNumPedido
        '
        Me.txtNumPedido.Enabled = False
        Me.txtNumPedido.Location = New System.Drawing.Point(347, 28)
        Me.txtNumPedido.Name = "txtNumPedido"
        Me.txtNumPedido.Size = New System.Drawing.Size(100, 20)
        Me.txtNumPedido.TabIndex = 2
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.ForeColor = System.Drawing.Color.Maroon
        Me.Label16.Location = New System.Drawing.Point(261, 31)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(40, 13)
        Me.Label16.TabIndex = 24
        Me.Label16.Text = "Pedido"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.ForeColor = System.Drawing.Color.Maroon
        Me.Label17.Location = New System.Drawing.Point(24, 68)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(79, 13)
        Me.Label17.TabIndex = 25
        Me.Label17.Text = "Nombre Cliente"
        '
        'txtNombreCliente
        '
        Me.txtNombreCliente.Enabled = False
        Me.txtNombreCliente.Location = New System.Drawing.Point(109, 65)
        Me.txtNombreCliente.Name = "txtNombreCliente"
        Me.txtNombreCliente.Size = New System.Drawing.Size(146, 20)
        Me.txtNombreCliente.TabIndex = 3
        '
        'txtSiglaPais
        '
        Me.txtSiglaPais.Enabled = False
        Me.txtSiglaPais.Location = New System.Drawing.Point(348, 66)
        Me.txtSiglaPais.Name = "txtSiglaPais"
        Me.txtSiglaPais.Size = New System.Drawing.Size(99, 20)
        Me.txtSiglaPais.TabIndex = 4
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.ForeColor = System.Drawing.Color.Maroon
        Me.Label18.Location = New System.Drawing.Point(262, 69)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(62, 13)
        Me.Label18.TabIndex = 26
        Me.Label18.Text = "Pais Cliente"
        '
        'txtNombProducto
        '
        Me.txtNombProducto.Enabled = False
        Me.txtNombProducto.Location = New System.Drawing.Point(134, 101)
        Me.txtNombProducto.Name = "txtNombProducto"
        Me.txtNombProducto.Size = New System.Drawing.Size(99, 20)
        Me.txtNombProducto.TabIndex = 5
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.ForeColor = System.Drawing.Color.Maroon
        Me.Label19.Location = New System.Drawing.Point(24, 104)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(90, 13)
        Me.Label19.TabIndex = 27
        Me.Label19.Text = "Nombre Producto"
        '
        'txtRefProducto
        '
        Me.txtRefProducto.Enabled = False
        Me.txtRefProducto.Location = New System.Drawing.Point(348, 101)
        Me.txtRefProducto.Name = "txtRefProducto"
        Me.txtRefProducto.Size = New System.Drawing.Size(99, 20)
        Me.txtRefProducto.TabIndex = 6
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.ForeColor = System.Drawing.Color.Maroon
        Me.Label20.Location = New System.Drawing.Point(259, 104)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(73, 13)
        Me.Label20.TabIndex = 28
        Me.Label20.Text = "Ref. Producto"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.ForeColor = System.Drawing.Color.Maroon
        Me.Label21.Location = New System.Drawing.Point(539, 127)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(92, 13)
        Me.Label21.TabIndex = 45
        Me.Label21.Text = "ETIQUETA LOTE"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.ForeColor = System.Drawing.Color.Maroon
        Me.Label22.Location = New System.Drawing.Point(809, 123)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(90, 13)
        Me.Label22.TabIndex = 46
        Me.Label22.Text = "ETIQUETA CAJA"
        '
        'lbEtiLoAl
        '
        Me.lbEtiLoAl.AutoSize = True
        Me.lbEtiLoAl.Location = New System.Drawing.Point(452, 264)
        Me.lbEtiLoAl.Name = "lbEtiLoAl"
        Me.lbEtiLoAl.Size = New System.Drawing.Size(41, 13)
        Me.lbEtiLoAl.TabIndex = 47
        Me.lbEtiLoAl.Text = "110mm"
        Me.lbEtiLoAl.Visible = False
        '
        'lbEtiLoAn
        '
        Me.lbEtiLoAn.AutoSize = True
        Me.lbEtiLoAn.Location = New System.Drawing.Point(560, 149)
        Me.lbEtiLoAn.Name = "lbEtiLoAn"
        Me.lbEtiLoAn.Size = New System.Drawing.Size(35, 13)
        Me.lbEtiLoAn.TabIndex = 48
        Me.lbEtiLoAn.Text = "60mm"
        Me.lbEtiLoAn.Visible = False
        '
        'lbEtiCaAn
        '
        Me.lbEtiCaAn.AutoSize = True
        Me.lbEtiCaAn.Location = New System.Drawing.Point(839, 149)
        Me.lbEtiCaAn.Name = "lbEtiCaAn"
        Me.lbEtiCaAn.Size = New System.Drawing.Size(35, 13)
        Me.lbEtiCaAn.TabIndex = 49
        Me.lbEtiCaAn.Text = "60mm"
        Me.lbEtiCaAn.Visible = False
        '
        'lbEtiCaAl
        '
        Me.lbEtiCaAl.AutoSize = True
        Me.lbEtiCaAl.Location = New System.Drawing.Point(708, 264)
        Me.lbEtiCaAl.Name = "lbEtiCaAl"
        Me.lbEtiCaAl.Size = New System.Drawing.Size(41, 13)
        Me.lbEtiCaAl.TabIndex = 50
        Me.lbEtiCaAl.Text = "110mm"
        Me.lbEtiCaAl.Visible = False
        '
        'cmdImpManual
        '
        Me.cmdImpManual.ForeColor = System.Drawing.Color.Maroon
        Me.cmdImpManual.Location = New System.Drawing.Point(383, 409)
        Me.cmdImpManual.Name = "cmdImpManual"
        Me.cmdImpManual.Size = New System.Drawing.Size(75, 77)
        Me.cmdImpManual.TabIndex = 52
        Me.cmdImpManual.Text = "Modificar Impresion Etiquetas "
        Me.cmdImpManual.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Etiquetas_Datamars.My.Resources.Resources.logo_datamars_color
        Me.PictureBox1.Location = New System.Drawing.Point(497, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(456, 85)
        Me.PictureBox1.TabIndex = 51
        Me.PictureBox1.TabStop = False
        '
        'picEtiquetaCaja
        '
        Me.picEtiquetaCaja.BackColor = System.Drawing.Color.Transparent
        Me.picEtiquetaCaja.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picEtiquetaCaja.Location = New System.Drawing.Point(753, 165)
        Me.picEtiquetaCaja.Name = "picEtiquetaCaja"
        Me.picEtiquetaCaja.Size = New System.Drawing.Size(200, 175)
        Me.picEtiquetaCaja.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picEtiquetaCaja.TabIndex = 44
        Me.picEtiquetaCaja.TabStop = False
        '
        'picEtiquetaLote
        '
        Me.picEtiquetaLote.BackColor = System.Drawing.Color.Transparent
        Me.picEtiquetaLote.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picEtiquetaLote.Location = New System.Drawing.Point(497, 165)
        Me.picEtiquetaLote.Name = "picEtiquetaLote"
        Me.picEtiquetaLote.Size = New System.Drawing.Size(200, 175)
        Me.picEtiquetaLote.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picEtiquetaLote.TabIndex = 43
        Me.picEtiquetaLote.TabStop = False
        '
        'numCajasPorPalet
        '
        Me.numCajasPorPalet.Enabled = False
        Me.numCajasPorPalet.Location = New System.Drawing.Point(133, 329)
        Me.numCajasPorPalet.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numCajasPorPalet.Name = "numCajasPorPalet"
        Me.numCajasPorPalet.Size = New System.Drawing.Size(120, 20)
        Me.numCajasPorPalet.TabIndex = 53
        Me.numCajasPorPalet.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.numCajasPorPalet.Value = New Decimal(New Integer() {50, 0, 0, 0})
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.ForeColor = System.Drawing.Color.Maroon
        Me.Label23.Location = New System.Drawing.Point(12, 331)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(78, 13)
        Me.Label23.TabIndex = 54
        Me.Label23.Text = "Cajas por Palet"
        '
        'numPalets
        '
        Me.numPalets.Enabled = False
        Me.numPalets.Location = New System.Drawing.Point(135, 463)
        Me.numPalets.Maximum = New Decimal(New Integer() {99999, 0, 0, 0})
        Me.numPalets.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numPalets.Name = "numPalets"
        Me.numPalets.Size = New System.Drawing.Size(120, 20)
        Me.numPalets.TabIndex = 55
        Me.numPalets.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.numPalets.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.ForeColor = System.Drawing.Color.Maroon
        Me.Label24.Location = New System.Drawing.Point(6, 465)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(126, 13)
        Me.Label24.TabIndex = 56
        Me.Label24.Text = "Cantidad palets a imprimir"
        '
        'cmdImprimirPalets
        '
        Me.cmdImprimirPalets.ForeColor = System.Drawing.Color.Maroon
        Me.cmdImprimirPalets.Location = New System.Drawing.Point(263, 463)
        Me.cmdImprimirPalets.Name = "cmdImprimirPalets"
        Me.cmdImprimirPalets.Size = New System.Drawing.Size(100, 23)
        Me.cmdImprimirPalets.TabIndex = 57
        Me.cmdImprimirPalets.Text = "Palets"
        Me.cmdImprimirPalets.UseVisualStyleBackColor = True
        '
        'txtnumsap
        '
        Me.txtnumsap.Location = New System.Drawing.Point(347, 168)
        Me.txtnumsap.Name = "txtnumsap"
        Me.txtnumsap.Size = New System.Drawing.Size(100, 20)
        Me.txtnumsap.TabIndex = 58
        Me.txtnumsap.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(979, 500)
        Me.Controls.Add(Me.txtnumsap)
        Me.Controls.Add(Me.cmdImprimirPalets)
        Me.Controls.Add(Me.numPalets)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.numCajasPorPalet)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.cmdImpManual)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lbEtiCaAl)
        Me.Controls.Add(Me.lbEtiCaAn)
        Me.Controls.Add(Me.lbEtiLoAn)
        Me.Controls.Add(Me.lbEtiLoAl)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.picEtiquetaCaja)
        Me.Controls.Add(Me.picEtiquetaLote)
        Me.Controls.Add(Me.txtRefProducto)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.txtNombProducto)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.txtSiglaPais)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.txtNombreCliente)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.txtNumPedido)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.txtLastCodigoChip)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.numUnidadesPorlote)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.txtGlobalFirstCode)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.txtCodPedido)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.numLotes)
        Me.Controls.Add(Me.numCajas)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.numAñoCad)
        Me.Controls.Add(Me.numMesCad)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.numLotesPorCaja)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtLote)
        Me.Controls.Add(Me.txtCaja)
        Me.Controls.Add(Me.txtAño)
        Me.Controls.Add(Me.txtSemana)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cmdImprimirCajas)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdImprimirLotes)
        Me.Name = "Form1"
        Me.Text = "Etiquetas Datamars"
        CType(Me.txtSemana, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAño, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCaja, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLote, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numLotesPorCaja, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numAñoCad, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numMesCad, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numLotes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numCajas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numUnidadesPorlote, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picEtiquetaCaja, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picEtiquetaLote, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numCajasPorPalet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numPalets, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdImprimirLotes As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents cmdImprimirCajas As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents txtSemana As NumericUpDown
    Friend WithEvents txtAño As NumericUpDown
    Friend WithEvents txtCaja As NumericUpDown
    Friend WithEvents txtLote As NumericUpDown
    Friend WithEvents numLotesPorCaja As NumericUpDown
    Friend WithEvents Label5 As Label
    Friend WithEvents numAñoCad As NumericUpDown
    Friend WithEvents numMesCad As NumericUpDown
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents numLotes As NumericUpDown
    Friend WithEvents numCajas As NumericUpDown
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents Label12 As Label
    Friend WithEvents txtCodPedido As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents txtGlobalFirstCode As TextBox
    Friend WithEvents numUnidadesPorlote As NumericUpDown
    Friend WithEvents Label14 As Label
    Friend WithEvents txtLastCodigoChip As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents txtNumPedido As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents txtNombreCliente As TextBox
    Friend WithEvents txtSiglaPais As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents txtNombProducto As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents txtRefProducto As TextBox
    Friend WithEvents Label20 As Label
    Friend WithEvents picEtiquetaLote As PictureBox
    Friend WithEvents picEtiquetaCaja As PictureBox
    Friend WithEvents Label21 As Label
    Friend WithEvents Label22 As Label
    Friend WithEvents lbEtiLoAl As Label
    Friend WithEvents lbEtiLoAn As Label
    Friend WithEvents lbEtiCaAn As Label
    Friend WithEvents lbEtiCaAl As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents cmdImpManual As Button
    Friend WithEvents numCajasPorPalet As NumericUpDown
    Friend WithEvents Label23 As Label
    Friend WithEvents numPalets As NumericUpDown
    Friend WithEvents Label24 As Label
    Friend WithEvents cmdImprimirPalets As Button
    Friend WithEvents txtnumsap As TextBox
End Class
