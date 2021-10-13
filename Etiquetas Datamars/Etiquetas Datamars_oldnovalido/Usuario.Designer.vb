<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Usuario
    Inherits System.Windows.Forms.UserControl

    'UserControl reemplaza a Dispose para limpiar la lista de componentes.
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
        Me.lblNombre = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblNombre
        '
        Me.lblNombre.Font = New System.Drawing.Font("Calibri", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNombre.Image = Global.Etiquetas_Datamars.My.Resources.Resources.icoUser120BBB
        Me.lblNombre.Location = New System.Drawing.Point(0, 0)
        Me.lblNombre.Name = "lblNombre"
        Me.lblNombre.Size = New System.Drawing.Size(120, 120)
        Me.lblNombre.TabIndex = 0
        Me.lblNombre.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Usuario
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.lblNombre)
        Me.Margin = New System.Windows.Forms.Padding(10)
        Me.Name = "Usuario"
        Me.Size = New System.Drawing.Size(120, 120)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblNombre As System.Windows.Forms.Label

End Class
