<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmTest04
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
        lblFormSpec = New System.Windows.Forms.Label()
        LbxFormSpec = New System.Windows.Forms.TextBox()
        LbxSpecOps = New System.Windows.Forms.TextBox()
        lblSpecOps = New System.Windows.Forms.Label()
        LbxSpecSel = New System.Windows.Forms.TextBox()
        lblSpecSel = New System.Windows.Forms.Label()
        SuspendLayout()
        ' 
        ' lblFormSpec
        ' 
        lblFormSpec.AutoSize = True
        lblFormSpec.Location = New System.Drawing.Point(12, 9)
        lblFormSpec.Name = "lblFormSpec"
        lblFormSpec.Size = New System.Drawing.Size(41, 15)
        lblFormSpec.TabIndex = 0
        lblFormSpec.Text = "Label1"
        ' 
        ' LbxFormSpec
        ' 
        LbxFormSpec.Enabled = True
        LbxFormSpec.Height = 15
        LbxFormSpec.Location = New System.Drawing.Point(24, 49)
        LbxFormSpec.Name = "LbxFormSpec"
        LbxFormSpec.Size = New System.Drawing.Size(120, 94)
        LbxFormSpec.TabIndex = 1
        ' 
        ' LbxSpecOps
        ' 
        LbxSpecOps.Enabled = True
        LbxSpecOps.Height = 15
        LbxSpecOps.Location = New System.Drawing.Point(24, 194)
        LbxSpecOps.Name = "LbxSpecOps"
        LbxSpecOps.Size = New System.Drawing.Size(120, 94)
        LbxSpecOps.TabIndex = 3
        ' 
        ' lblSpecOps
        ' 
        lblSpecOps.AutoSize = True
        lblSpecOps.Location = New System.Drawing.Point(12, 154)
        lblSpecOps.Name = "lblSpecOps"
        lblSpecOps.Size = New System.Drawing.Size(41, 15)
        lblSpecOps.TabIndex = 2
        lblSpecOps.Text = "Label2"
        ' 
        ' LbxSpecSel
        ' 
        LbxSpecSel.Enabled = True
        LbxSpecSel.Height = 15
        LbxSpecSel.Location = New System.Drawing.Point(24, 336)
        LbxSpecSel.Name = "LbxSpecSel"
        LbxSpecSel.Size = New System.Drawing.Size(120, 94)
        LbxSpecSel.TabIndex = 5
        ' 
        ' lblSpecSel
        ' 
        lblSpecSel.AutoSize = True
        lblSpecSel.Location = New System.Drawing.Point(12, 296)
        lblSpecSel.Name = "lblSpecSel"
        lblSpecSel.Size = New System.Drawing.Size(41, 15)
        lblSpecSel.TabIndex = 4
        lblSpecSel.Text = "Label3"
        ' 
        ' FmTest04
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7F, 15F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        ClientSize = New System.Drawing.Size(384, 463)
        Controls.Add(LbxSpecSel)
        Controls.Add(lblSpecSel)
        Controls.Add(LbxSpecOps)
        Controls.Add(lblSpecOps)
        Controls.Add(LbxFormSpec)
        Controls.Add(lblFormSpec)
        Name = "fmTest04"
        Text = "fmTest04"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents lblFormSpec As System.Windows.Forms.Label
    Friend WithEvents LbxFormSpec As System.Windows.Forms.TextBox
    Friend WithEvents LbxSpecOps As System.Windows.Forms.TextBox
    Friend WithEvents lblSpecOps As System.Windows.Forms.Label
    Friend WithEvents LbxSpecSel As System.Windows.Forms.TextBox
    Friend WithEvents lblSpecSel As System.Windows.Forms.Label
End Class
