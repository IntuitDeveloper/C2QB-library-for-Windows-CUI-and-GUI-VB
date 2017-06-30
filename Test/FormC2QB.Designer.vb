<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormC2QB
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.C2QB = New System.Windows.Forms.Button()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.SuspendLayout()
        '
        'C2QB
        '
        Me.C2QB.Location = New System.Drawing.Point(0, 0)
        Me.C2QB.Name = "C2QB"
        Me.C2QB.Size = New System.Drawing.Size(291, 253)
        Me.C2QB.TabIndex = 0
        Me.C2QB.Text = "C2QB"
        Me.C2QB.UseVisualStyleBackColor = True
        '
        'FormC2QB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(506, 459)
        Me.Controls.Add(Me.C2QB)
        Me.Name = "FormC2QB"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents C2QB As Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
