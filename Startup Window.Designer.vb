<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class StartupWindow
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(StartupWindow))
        Me.OpenSourceDBFWindow = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'OpenSourceDBFWindow
        '
        Me.OpenSourceDBFWindow.Location = New System.Drawing.Point(15, 15)
        Me.OpenSourceDBFWindow.Name = "OpenSourceDBFWindow"
        Me.OpenSourceDBFWindow.Size = New System.Drawing.Size(194, 119)
        Me.OpenSourceDBFWindow.TabIndex = 0
        Me.OpenSourceDBFWindow.Text = "Edit SOURCE.DBF File"
        Me.OpenSourceDBFWindow.UseVisualStyleBackColor = True
        '
        'StartupWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(231, 160)
        Me.Controls.Add(Me.OpenSourceDBFWindow)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "StartupWindow"
        Me.Padding = New System.Windows.Forms.Padding(12)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Xerox Bus Editor"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents OpenSourceDBFWindow As Button
End Class
