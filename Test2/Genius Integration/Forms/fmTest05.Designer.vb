<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FmTest05
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
        TbsItemGrps = New System.Windows.Forms.TabControl()
        LbxItems = New System.Windows.Forms.TabPage()
        TabPage2 = New System.Windows.Forms.TabPage()
        LblPartNum = New System.Windows.Forms.Label()
        LblDesc = New System.Windows.Forms.Label()
        ImgOfItem = New System.Windows.Forms.PictureBox()
        CmdOpenItem = New System.Windows.Forms.Button()
        cmdEndCancel = New System.Windows.Forms.Button()
        cmdEndSave = New System.Windows.Forms.Button()
        TbsItemGrps.SuspendLayout()
        CType(ImgOfItem, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' TbsItemGrps
        ' 
        TbsItemGrps.Controls.Add(LbxItems)
        TbsItemGrps.Controls.Add(TabPage2)
        TbsItemGrps.Location = New System.Drawing.Point(12, 12)
        TbsItemGrps.Name = "TbsItemGrps"
        TbsItemGrps.SelectedIndex = 0
        TbsItemGrps.Size = New System.Drawing.Size(238, 426)
        TbsItemGrps.TabIndex = 0
        ' 
        ' LbxItems
        ' 
        LbxItems.Location = New System.Drawing.Point(4, 24)
        LbxItems.Name = "LbxItems"
        LbxItems.Padding = New System.Windows.Forms.Padding(3)
        LbxItems.Size = New System.Drawing.Size(230, 398)
        LbxItems.TabIndex = 0
        LbxItems.Text = "TabPage1"
        LbxItems.UseVisualStyleBackColor = True
        ' 
        ' TabPage2
        ' 
        TabPage2.Location = New System.Drawing.Point(4, 24)
        TabPage2.Name = "TabPage2"
        TabPage2.Padding = New System.Windows.Forms.Padding(3)
        TabPage2.Size = New System.Drawing.Size(230, 398)
        TabPage2.TabIndex = 1
        TabPage2.Text = "TabPage2"
        TabPage2.UseVisualStyleBackColor = True
        ' 
        ' LblPartNum
        ' 
        LblPartNum.AutoSize = True
        LblPartNum.Location = New System.Drawing.Point(273, 36)
        LblPartNum.Name = "LblPartNum"
        LblPartNum.Size = New System.Drawing.Size(41, 15)
        LblPartNum.TabIndex = 1
        LblPartNum.Text = "Label1"
        ' 
        ' LblDesc
        ' 
        LblDesc.AutoSize = True
        LblDesc.Location = New System.Drawing.Point(273, 62)
        LblDesc.Name = "LblDesc"
        LblDesc.Size = New System.Drawing.Size(41, 15)
        LblDesc.TabIndex = 2
        LblDesc.Text = "Label2"
        ' 
        ' ImgOfItem
        ' 
        ImgOfItem.Location = New System.Drawing.Point(320, 62)
        ImgOfItem.Name = "ImgOfItem"
        ImgOfItem.Size = New System.Drawing.Size(111, 109)
        ImgOfItem.TabIndex = 3
        ImgOfItem.TabStop = False
        ' 
        ' CmdOpenItem
        ' 
        CmdOpenItem.Location = New System.Drawing.Point(356, 32)
        CmdOpenItem.Name = "CmdOpenItem"
        CmdOpenItem.Size = New System.Drawing.Size(75, 23)
        CmdOpenItem.TabIndex = 4
        CmdOpenItem.Text = "Button1"
        CmdOpenItem.UseVisualStyleBackColor = True
        ' 
        ' cmdEndCancel
        ' 
        cmdEndCancel.Location = New System.Drawing.Point(273, 411)
        cmdEndCancel.Name = "cmdEndCancel"
        cmdEndCancel.Size = New System.Drawing.Size(75, 23)
        cmdEndCancel.TabIndex = 5
        cmdEndCancel.Text = "Button2"
        cmdEndCancel.UseVisualStyleBackColor = True
        ' 
        ' cmdEndSave
        ' 
        cmdEndSave.Location = New System.Drawing.Point(356, 411)
        cmdEndSave.Name = "cmdEndSave"
        cmdEndSave.Size = New System.Drawing.Size(75, 23)
        cmdEndSave.TabIndex = 6
        cmdEndSave.Text = "Button3"
        cmdEndSave.UseVisualStyleBackColor = True
        ' 
        ' FmTest05
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7F, 15F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        ClientSize = New System.Drawing.Size(443, 450)
        Controls.Add(cmdEndSave)
        Controls.Add(cmdEndCancel)
        Controls.Add(CmdOpenItem)
        Controls.Add(ImgOfItem)
        Controls.Add(LblDesc)
        Controls.Add(LblPartNum)
        Controls.Add(TbsItemGrps)
        Name = "FmTest05"
        Text = "Form1"
        TbsItemGrps.ResumeLayout(False)
        CType(ImgOfItem, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents TbsItemGrps As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents LblPartNum As System.Windows.Forms.Label
    Friend WithEvents LblDesc As System.Windows.Forms.Label
    Friend WithEvents ImgOfItem As System.Windows.Forms.PictureBox
    Friend WithEvents CmdOpenItem As System.Windows.Forms.Button
    Friend WithEvents cmdEndCancel As System.Windows.Forms.Button
    Friend WithEvents cmdEndSave As System.Windows.Forms.Button
    Friend WithEvents LbxItems As System.Windows.Forms.TabPage
End Class
