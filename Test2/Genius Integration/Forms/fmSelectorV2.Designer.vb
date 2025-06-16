<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class fmSelectorV2
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
        tbxView = New System.Windows.Forms.TextBox()
        lsbSelection = New System.Windows.Forms.ListBox()
        btnCancel = New System.Windows.Forms.Button()
        btnOk = New System.Windows.Forms.Button()
        SuspendLayout()
        ' 
        ' tbxView
        ' 
        tbxView.Location = New System.Drawing.Point(284, 134)
        tbxView.Name = "tbxView"
        tbxView.Size = New System.Drawing.Size(100, 23)
        tbxView.TabIndex = 0
        ' 
        ' lsbSelection
        ' 
        lsbSelection.FormattingEnabled = True
        lsbSelection.ItemHeight = 15
        lsbSelection.Location = New System.Drawing.Point(284, 213)
        lsbSelection.Name = "lsbSelection"
        lsbSelection.Size = New System.Drawing.Size(120, 94)
        lsbSelection.TabIndex = 1
        ' 
        ' btnCancel
        ' 
        btnCancel.Location = New System.Drawing.Point(308, 337)
        btnCancel.Name = "btnCancel"
        btnCancel.Size = New System.Drawing.Size(75, 23)
        btnCancel.TabIndex = 2
        btnCancel.Text = "Button1"
        btnCancel.UseVisualStyleBackColor = True
        ' 
        ' btnOk
        ' 
        btnOk.Location = New System.Drawing.Point(438, 337)
        btnOk.Name = "btnOk"
        btnOk.Size = New System.Drawing.Size(75, 23)
        btnOk.TabIndex = 3
        btnOk.Text = "Button2"
        btnOk.UseVisualStyleBackColor = True
        ' 
        ' fmSelectorV2
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7.0F, 15.0F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        ClientSize = New System.Drawing.Size(800, 450)
        Controls.Add(btnOk)
        Controls.Add(btnCancel)
        Controls.Add(lsbSelection)
        Controls.Add(tbxView)
        Name = "fmSelectorV2"
        Text = "Form1"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents tbxView As System.Windows.Forms.TextBox
    Friend WithEvents lsbSelection As System.Windows.Forms.ListBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
End Class
