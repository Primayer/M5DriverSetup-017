<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormM5Install
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
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

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormM5Install))
        Me.M5_Setup1 = New LibM5DriverInstall.M5_Setup
        Me.SuspendLayout()
        '
        'M5_Setup1
        '
        Me.M5_Setup1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.M5_Setup1.EnableAutoInstall = True
        Me.M5_Setup1.Location = New System.Drawing.Point(0, 0)
        Me.M5_Setup1.Name = "M5_Setup1"
        Me.M5_Setup1.Size = New System.Drawing.Size(721, 391)
        Me.M5_Setup1.TabIndex = 0
        '
        'FormM5Install
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(721, 391)
        Me.Controls.Add(Me.M5_Setup1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormM5Install"
        Me.Text = "M5 driver installer"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents M5_Setup1 As LibM5DriverInstall.M5_Setup
End Class
