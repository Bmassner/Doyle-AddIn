<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class fmGetList
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
        cmdCancel = New System.Windows.Forms.Button()
        cmdOk = New System.Windows.Forms.Button()
        lbTxIn = New System.Windows.Forms.Label()
        txIn = New System.Windows.Forms.ListBox()
        SuspendLayout()
        ' 
        ' cmdCancel
        ' 
        cmdCancel.Location = New System.Drawing.Point(12, 307)
        cmdCancel.Name = "cmdCancel"
        cmdCancel.Size = New System.Drawing.Size(75, 23)
        cmdCancel.TabIndex = 0
        cmdCancel.Text = "Cancel"
        cmdCancel.UseVisualStyleBackColor = True
        ' 
        ' cmdOk
        ' 
        cmdOk.Location = New System.Drawing.Point(88, 307)
        cmdOk.Name = "cmdOk"
        cmdOk.Size = New System.Drawing.Size(75, 23)
        cmdOk.TabIndex = 1
        cmdOk.Text = "Ok"
        cmdOk.UseVisualStyleBackColor = True
        ' 
        ' lbTxIn
        ' 
        lbTxIn.AutoSize = True
        lbTxIn.Location = New System.Drawing.Point(30, 146)
        lbTxIn.Name = "lbTxIn"
        lbTxIn.Size = New System.Drawing.Size(126, 15)
        lbTxIn.TabIndex = 2
        lbTxIn.Text = "Paste or Type List Here"
        ' 
        ' txIn
        ' 
        txIn.FormattingEnabled = True
        txIn.ItemHeight = 15
        txIn.Location = New System.Drawing.Point(12, 12)
        txIn.Name = "txIn"
        txIn.Size = New System.Drawing.Size(150, 289)
        txIn.TabIndex = 4
        ' 
        ' fmGetList
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7.0F, 15.0F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        BackColor = Drawing.SystemColors.ButtonFace
        ClientSize = New System.Drawing.Size(180, 339)
        Controls.Add(txIn)
        Controls.Add(lbTxIn)
        Controls.Add(cmdOk)
        Controls.Add(cmdCancel)
        Name = "fmGetList"
        Text = "List Entry"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents lbTxIn As System.Windows.Forms.Label
    Friend WithEvents txIn As System.Windows.Forms.ListBox
End Class
