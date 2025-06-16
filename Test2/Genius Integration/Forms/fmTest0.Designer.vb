<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class fmTest0
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
        imTNail = New System.Windows.Forms.PictureBox()
        CType(imTNail, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' imTNail
        ' 
        imTNail.Location = New System.Drawing.Point(170, 144)
        imTNail.Name = "imTNail"
        imTNail.Size = New System.Drawing.Size(100, 50)
        imTNail.TabIndex = 0
        imTNail.TabStop = False
        ' 
        ' fmTest0
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7.0F, 15.0F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        ClientSize = New System.Drawing.Size(377, 381)
        Controls.Add(imTNail)
        Name = "fmTest0"
        Text = "Form1"
        CType(imTNail, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
    End Sub

    Friend WithEvents imTNail As System.Windows.Forms.PictureBox
End Class
