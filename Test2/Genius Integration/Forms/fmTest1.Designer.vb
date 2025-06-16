<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FmTest1
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
        imThmNail = New System.Windows.Forms.PictureBox()
        lbMsg = New System.Windows.Forms.Label()
        lbMtFamily = New System.Windows.Forms.Label()
        LbxFamily = New System.Windows.Forms.TextBox()
        lblFamily = New System.Windows.Forms.Label()
        dbFamily = New System.Windows.Forms.ComboBox()
        lblItem = New System.Windows.Forms.Label()
        LbxItem = New System.Windows.Forms.TextBox()
        lblNoImg = New System.Windows.Forms.Label()
        CType(imThmNail, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' imThmNail
        ' 
        imThmNail.Location = New System.Drawing.Point(12, 12)
        imThmNail.Name = "imThmNail"
        imThmNail.Size = New System.Drawing.Size(237, 201)
        imThmNail.TabIndex = 0
        imThmNail.TabStop = False
        ' 
        ' lbMsg
        ' 
        lbMsg.AutoSize = True
        lbMsg.Location = New System.Drawing.Point(293, 35)
        lbMsg.Name = "lbMsg"
        lbMsg.Size = New System.Drawing.Size(41, 15)
        lbMsg.TabIndex = 1
        lbMsg.Text = "Label1"
        ' 
        ' lbMtFamily
        ' 
        lbMtFamily.AutoSize = True
        lbMtFamily.Location = New System.Drawing.Point(306, 98)
        lbMtFamily.Name = "lbMtFamily"
        lbMtFamily.Size = New System.Drawing.Size(41, 15)
        lbMtFamily.TabIndex = 2
        lbMtFamily.Text = "Label1"
        ' 
        ' LbxFamily
        ' 
        LbxFamily.Enabled = True
        LbxFamily.Height = 15
        LbxFamily.Location = New System.Drawing.Point(293, 116)
        LbxFamily.Name = "LbxFamily"
        LbxFamily.Size = New System.Drawing.Size(120, 94)
        LbxFamily.TabIndex = 3
        ' 
        ' lblFamily
        ' 
        lblFamily.AutoSize = True
        lblFamily.Location = New System.Drawing.Point(12, 241)
        lblFamily.Name = "lblFamily"
        lblFamily.Size = New System.Drawing.Size(41, 15)
        lblFamily.TabIndex = 4
        lblFamily.Text = "Label1"
        ' 
        ' dbFamily
        ' 
        dbFamily.FormattingEnabled = True
        dbFamily.Location = New System.Drawing.Point(12, 268)
        dbFamily.Name = "dbFamily"
        dbFamily.Size = New System.Drawing.Size(121, 23)
        dbFamily.TabIndex = 5
        ' 
        ' lblItem
        ' 
        lblItem.AutoSize = True
        lblItem.Location = New System.Drawing.Point(12, 294)
        lblItem.Name = "lblItem"
        lblItem.Size = New System.Drawing.Size(41, 15)
        lblItem.TabIndex = 6
        lblItem.Text = "Label1"
        ' 
        ' LbxItem
        ' 
        LbxItem.Enabled = True
        LbxItem.Height = 15
        LbxItem.Location = New System.Drawing.Point(12, 312)
        LbxItem.Name = "LbxItem"
        LbxItem.Size = New System.Drawing.Size(425, 184)
        LbxItem.TabIndex = 7
        ' 
        ' lblNoImg
        ' 
        lblNoImg.AutoSize = True
        lblNoImg.Location = New System.Drawing.Point(92, 116)
        lblNoImg.Name = "lblNoImg"
        lblNoImg.Size = New System.Drawing.Size(41, 15)
        lblNoImg.TabIndex = 8
        lblNoImg.Text = "Label1"
        ' 
        ' FmTest1
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7.0F, 15.0F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        ClientSize = New System.Drawing.Size(461, 504)
        Controls.Add(lblNoImg)
        Controls.Add(LbxItem)
        Controls.Add(lblItem)
        Controls.Add(dbFamily)
        Controls.Add(lblFamily)
        Controls.Add(LbxFamily)
        Controls.Add(lbMtFamily)
        Controls.Add(lbMsg)
        Controls.Add(imThmNail)
        Name = "fmTest1"
        Text = "Form1"
        CType(imThmNail, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents imThmNail As System.Windows.Forms.PictureBox
    Friend WithEvents lbMsg As System.Windows.Forms.Label
    Friend WithEvents lbMtFamily As System.Windows.Forms.Label
    Friend WithEvents LbxFamily As System.Windows.Forms.TextBox
    Friend WithEvents lblFamily As System.Windows.Forms.Label
    Friend WithEvents dbFamily As System.Windows.Forms.ComboBox
    Friend WithEvents lblItem As System.Windows.Forms.Label
    Friend WithEvents LbxItem As System.Windows.Forms.TextBox
    Friend WithEvents lblNoImg As System.Windows.Forms.Label
End Class
