<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FmSelectorList
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
        btnCancel = New System.Windows.Forms.Button()
        btnOk = New System.Windows.Forms.Button()
        lsbSelection = New System.Windows.Forms.TextBox()
        SuspendLayout()
        ' 
        ' btnCancel
        ' 
        btnCancel.Location = New System.Drawing.Point(57, 210)
        btnCancel.Name = "btnCancel"
        btnCancel.Size = New System.Drawing.Size(75, 23)
        btnCancel.TabIndex = 0
        btnCancel.Text = "Button1"
        btnCancel.UseVisualStyleBackColor = True
        ' 
        ' btnOk
        ' 
        btnOk.Location = New System.Drawing.Point(204, 213)
        btnOk.Name = "btnOk"
        btnOk.Size = New System.Drawing.Size(75, 23)
        btnOk.TabIndex = 1
        btnOk.Text = "Button2"
        btnOk.UseVisualStyleBackColor = True
        ' 
        ' lsbSelection
        ' 
        lsbSelection.Enabled = True
        lsbSelection.Height = 15
        lsbSelection.Location = New System.Drawing.Point(57, 28)
        lsbSelection.Name = "lsbSelection"
        lsbSelection.Size = New System.Drawing.Size(222, 154)
        lsbSelection.TabIndex = 2
        ' 
        ' FmSelectorList
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7.0F, 15.0F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        ClientSize = New System.Drawing.Size(347, 249)
        Controls.Add(lsbSelection)
        Controls.Add(btnOk)
        Controls.Add(btnCancel)
        Name = "FmSelectorList"
        Text = "Form1"
        ResumeLayout(False)
    End Sub

    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents lsbSelection As System.Windows.Forms.TextBox
End Class
