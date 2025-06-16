<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FmMatlQty
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
        LblPartNumber = New System.Windows.Forms.Label()
        imThmNail = New System.Windows.Forms.PictureBox()
        lblPartInfo = New System.Windows.Forms.Label()
        lblMatlNumber = New System.Windows.Forms.Label()
        lblMatlInfo = New System.Windows.Forms.Label()
        lblMatlQty = New System.Windows.Forms.Label()
        txbMatlQty = New System.Windows.Forms.TextBox()
        cbxUnitQty = New System.Windows.Forms.ComboBox()
        cmdOK = New System.Windows.Forms.Button()
        cmdCancel = New System.Windows.Forms.Button()
        LbxMatlQty = New System.Windows.Forms.TextBox()
        CType(imThmNail, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' LblPartNumber
        ' 
        LblPartNumber.AccessibleName = "<partNumber>"
        LblPartNumber.Location = New System.Drawing.Point(18, 6)
        LblPartNumber.Name = "LblPartNumber"
        LblPartNumber.Size = New System.Drawing.Size(186, 18)
        LblPartNumber.TabIndex = 0
        LblPartNumber.Tag = "<partNumber>"
        ' 
        ' imThmNail
        ' 
        imThmNail.Location = New System.Drawing.Point(12, 24)
        imThmNail.Name = "imThmNail"
        imThmNail.Size = New System.Drawing.Size(186, 156)
        imThmNail.TabIndex = 1
        imThmNail.TabStop = False
        ' 
        ' lblPartInfo
        ' 
        lblPartInfo.Location = New System.Drawing.Point(72, 183)
        lblPartInfo.Name = "lblPartInfo"
        lblPartInfo.Size = New System.Drawing.Size(41, 15)
        lblPartInfo.TabIndex = 2
        lblPartInfo.Text = "Label2"
        ' 
        ' lblMatlNumber
        ' 
        lblMatlNumber.Location = New System.Drawing.Point(72, 198)
        lblMatlNumber.Name = "lblMatlNumber"
        lblMatlNumber.Size = New System.Drawing.Size(41, 15)
        lblMatlNumber.TabIndex = 3
        lblMatlNumber.Text = "Label3"
        ' 
        ' lblMatlInfo
        ' 
        lblMatlInfo.Location = New System.Drawing.Point(18, 225)
        lblMatlInfo.Name = "lblMatlInfo"
        lblMatlInfo.Size = New System.Drawing.Size(100, 23)
        lblMatlInfo.TabIndex = 4
        lblMatlInfo.Text = "Label4"
        ' 
        ' lblMatlQty
        ' 
        lblMatlQty.Location = New System.Drawing.Point(18, 240)
        lblMatlQty.Name = "lblMatlQty"
        lblMatlQty.Size = New System.Drawing.Size(41, 15)
        lblMatlQty.TabIndex = 5
        lblMatlQty.Text = "Label5"
        ' 
        ' txbMatlQty
        ' 
        txbMatlQty.Location = New System.Drawing.Point(12, 258)
        txbMatlQty.Name = "txbMatlQty"
        txbMatlQty.Size = New System.Drawing.Size(100, 23)
        txbMatlQty.TabIndex = 6
        ' 
        ' cbxUnitQty
        ' 
        cbxUnitQty.FormattingEnabled = True
        cbxUnitQty.Location = New System.Drawing.Point(12, 287)
        cbxUnitQty.Name = "cbxUnitQty"
        cbxUnitQty.Size = New System.Drawing.Size(121, 23)
        cbxUnitQty.TabIndex = 7
        ' 
        ' cmdOK
        ' 
        cmdOK.Location = New System.Drawing.Point(18, 382)
        cmdOK.Name = "cmdOK"
        cmdOK.Size = New System.Drawing.Size(75, 23)
        cmdOK.TabIndex = 8
        cmdOK.Text = "Button1"
        cmdOK.UseVisualStyleBackColor = True
        ' 
        ' cmdCancel
        ' 
        cmdCancel.Location = New System.Drawing.Point(105, 382)
        cmdCancel.Name = "cmdCancel"
        cmdCancel.Size = New System.Drawing.Size(75, 23)
        cmdCancel.TabIndex = 9
        cmdCancel.Text = "Button2"
        cmdCancel.UseVisualStyleBackColor = True
        ' 
        ' LbxMatlQty
        ' 
        LbxMatlQty.Enabled = True
        LbxMatlQty.Height = 15
        LbxMatlQty.Location = New System.Drawing.Point(139, 258)
        LbxMatlQty.Name = "LbxMatlQty"
        LbxMatlQty.Size = New System.Drawing.Size(120, 94)
        LbxMatlQty.TabIndex = 10
        ' 
        ' FmMatlQty
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7F, 15F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        ClientSize = New System.Drawing.Size(272, 450)
        Controls.Add(LbxMatlQty)
        Controls.Add(cmdCancel)
        Controls.Add(cmdOK)
        Controls.Add(cbxUnitQty)
        Controls.Add(txbMatlQty)
        Controls.Add(lblMatlQty)
        Controls.Add(lblMatlInfo)
        Controls.Add(lblMatlNumber)
        Controls.Add(lblPartInfo)
        Controls.Add(imThmNail)
        Controls.Add(LblPartNumber)
        Name = "fmMatlQty"
        Text = "Set/Verify Material Quantity"
        CType(imThmNail, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents LblPartNumber As System.Windows.Forms.Label
    Friend WithEvents imThmNail As System.Windows.Forms.PictureBox
    Friend WithEvents lblPartInfo As System.Windows.Forms.Label
    Friend WithEvents lblMatlNumber As System.Windows.Forms.Label
    Friend WithEvents lblMatlInfo As System.Windows.Forms.Label
    Friend WithEvents lblMatlQty As System.Windows.Forms.Label
    Friend WithEvents txbMatlQty As System.Windows.Forms.TextBox
    Friend WithEvents cbxUnitQty As System.Windows.Forms.ComboBox
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents LbxMatlQty As System.Windows.Forms.TextBox
End Class
