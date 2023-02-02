<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3))
        Me.PDF1 = New AxAcroPDFLib.AxAcroPDF()
        CType(Me.PDF1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PDF1
        '
        Me.PDF1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PDF1.Enabled = True
        Me.PDF1.Location = New System.Drawing.Point(0, 0)
        Me.PDF1.Name = "PDF1"
        Me.PDF1.OcxState = CType(resources.GetObject("PDF1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.PDF1.Size = New System.Drawing.Size(1292, 1049)
        Me.PDF1.TabIndex = 0
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1292, 1049)
        Me.Controls.Add(Me.PDF1)
        Me.Name = "Form3"
        Me.Text = "Manual"
        CType(Me.PDF1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PDF1 As AxAcroPDFLib.AxAcroPDF
End Class
