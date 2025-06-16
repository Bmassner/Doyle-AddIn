<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FmTest2
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
        cmdLt = New System.Windows.Forms.Button()
        cmdCt = New System.Windows.Forms.Button()
        cmdRt = New System.Windows.Forms.Button()
        lbMsg = New System.Windows.Forms.Label()
        imThmNail = New System.Windows.Forms.PictureBox()
        lblNoImg = New System.Windows.Forms.Label()
        CType(imThmNail, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' cmdLt
        ' 
        cmdLt.Location = New System.Drawing.Point(12, 342)
        cmdLt.Name = "cmdLt"
        cmdLt.Size = New System.Drawing.Size(75, 23)
        cmdLt.TabIndex = 0
        cmdLt.Text = "Button1"
        cmdLt.UseVisualStyleBackColor = True
        ' 
        ' cmdCt
        ' 
        cmdCt.Location = New System.Drawing.Point(93, 342)
        cmdCt.Name = "cmdCt"
        cmdCt.Size = New System.Drawing.Size(75, 23)
        cmdCt.TabIndex = 1
        cmdCt.Text = "Button2"
        cmdCt.UseVisualStyleBackColor = True
        ' 
        ' cmdRt
        ' 
        cmdRt.Location = New System.Drawing.Point(182, 342)
        cmdRt.Name = "cmdRt"
        cmdRt.Size = New System.Drawing.Size(75, 23)
        cmdRt.TabIndex = 2
        cmdRt.Text = "Button3"
        cmdRt.UseVisualStyleBackColor = True
        ' 
        ' lbMsg
        ' 
        lbMsg.AutoSize = True
        lbMsg.Location = New System.Drawing.Point(81, 248)
        lbMsg.Name = "lbMsg"
        lbMsg.Size = New System.Drawing.Size(41, 15)
        lbMsg.TabIndex = 3
        lbMsg.Text = "Label1"
        ' 
        ' imThmNail
        ' 
        imThmNail.Location = New System.Drawing.Point(12, 12)
        imThmNail.Name = "imThmNail"
        imThmNail.Size = New System.Drawing.Size(245, 190)
        imThmNail.TabIndex = 4
        imThmNail.TabStop = False
        ' 
        ' lblNoImg
        ' 
        lblNoImg.AutoSize = True
        lblNoImg.Location = New System.Drawing.Point(103, 94)
        lblNoImg.Name = "lblNoImg"
        lblNoImg.Size = New System.Drawing.Size(41, 15)
        lblNoImg.TabIndex = 5
        lblNoImg.Text = "Label2"
        ' 
        ' FmTest2
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7.0F, 15.0F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        ClientSize = New System.Drawing.Size(281, 404)
        Controls.Add(lblNoImg)
        Controls.Add(imThmNail)
        Controls.Add(lbMsg)
        Controls.Add(cmdRt)
        Controls.Add(cmdCt)
        Controls.Add(cmdLt)
        Name = "fmTest2"
        Text = "Form1"
        CType(imThmNail, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents cmdLt As System.Windows.Forms.Button
    Friend WithEvents cmdCt As System.Windows.Forms.Button
    Friend WithEvents cmdRt As System.Windows.Forms.Button
    Friend WithEvents lbMsg As System.Windows.Forms.Label
    Friend WithEvents imThmNail As System.Windows.Forms.PictureBox
    Friend WithEvents lblNoImg As System.Windows.Forms.Label
End Class
