<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUsuarios
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
        Me.panel = New System.Windows.Forms.FlowLayoutPanel()
        Me.SuspendLayout()
        '
        'panel
        '
        Me.panel.AutoScroll = True
        Me.panel.BackColor = System.Drawing.Color.White
        Me.panel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.panel.Location = New System.Drawing.Point(0, 0)
        Me.panel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.panel.Name = "panel"
        Me.panel.Size = New System.Drawing.Size(983, 482)
        Me.panel.TabIndex = 0
        '
        'frmUsuarios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(983, 482)
        Me.Controls.Add(Me.panel)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmUsuarios"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Choose a user to start"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents panel As System.Windows.Forms.FlowLayoutPanel
End Class
