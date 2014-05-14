<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.btnRun = New System.Windows.Forms.Button()
        Me.picImage = New System.Windows.Forms.PictureBox()
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog()
        Me.btnConvert = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.lblSubject = New System.Windows.Forms.Label()
        Me.lblFrom = New System.Windows.Forms.Label()
        Me.lblAcc = New System.Windows.Forms.Label()
        Me.txtSubject = New System.Windows.Forms.TextBox()
        Me.txtFrom = New System.Windows.Forms.TextBox()
        Me.txtAcc = New System.Windows.Forms.TextBox()
        Me.btnNext = New System.Windows.Forms.Button()
        CType(Me.picImage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(148, 287)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(75, 23)
        Me.btnRun.TabIndex = 0
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'picImage
        '
        Me.picImage.Location = New System.Drawing.Point(246, 0)
        Me.picImage.Name = "picImage"
        Me.picImage.Size = New System.Drawing.Size(160, 179)
        Me.picImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picImage.TabIndex = 1
        Me.picImage.TabStop = False
        '
        'dlgOpen
        '
        Me.dlgOpen.Multiselect = True
        '
        'btnConvert
        '
        Me.btnConvert.Location = New System.Drawing.Point(164, 241)
        Me.btnConvert.Name = "btnConvert"
        Me.btnConvert.Size = New System.Drawing.Size(75, 23)
        Me.btnConvert.TabIndex = 2
        Me.btnConvert.Text = "Convert"
        Me.btnConvert.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(-1, 13)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(414, 342)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnNext)
        Me.TabPage1.Controls.Add(Me.txtAcc)
        Me.TabPage1.Controls.Add(Me.txtFrom)
        Me.TabPage1.Controls.Add(Me.txtSubject)
        Me.TabPage1.Controls.Add(Me.lblAcc)
        Me.TabPage1.Controls.Add(Me.lblFrom)
        Me.TabPage1.Controls.Add(Me.lblSubject)
        Me.TabPage1.Controls.Add(Me.picImage)
        Me.TabPage1.Controls.Add(Me.btnRun)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(406, 316)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Outlook"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnConvert)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(406, 316)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Word -> tif"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(406, 281)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Rightfax"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'lblSubject
        '
        Me.lblSubject.AutoSize = True
        Me.lblSubject.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSubject.Location = New System.Drawing.Point(7, 7)
        Me.lblSubject.Name = "lblSubject"
        Me.lblSubject.Size = New System.Drawing.Size(46, 13)
        Me.lblSubject.TabIndex = 2
        Me.lblSubject.Text = "Subject:"
        '
        'lblFrom
        '
        Me.lblFrom.AutoSize = True
        Me.lblFrom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFrom.Location = New System.Drawing.Point(7, 30)
        Me.lblFrom.Name = "lblFrom"
        Me.lblFrom.Size = New System.Drawing.Size(44, 13)
        Me.lblFrom.TabIndex = 3
        Me.lblFrom.Text = "Sender:"
        '
        'lblAcc
        '
        Me.lblAcc.AutoSize = True
        Me.lblAcc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAcc.Location = New System.Drawing.Point(7, 52)
        Me.lblAcc.Name = "lblAcc"
        Me.lblAcc.Size = New System.Drawing.Size(53, 13)
        Me.lblAcc.TabIndex = 4
        Me.lblAcc.Text = "Account: "
        '
        'txtSubject
        '
        Me.txtSubject.Enabled = False
        Me.txtSubject.Location = New System.Drawing.Point(59, 4)
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.ReadOnly = True
        Me.txtSubject.Size = New System.Drawing.Size(123, 20)
        Me.txtSubject.TabIndex = 5
        '
        'txtFrom
        '
        Me.txtFrom.Enabled = False
        Me.txtFrom.Location = New System.Drawing.Point(59, 27)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.ReadOnly = True
        Me.txtFrom.Size = New System.Drawing.Size(123, 20)
        Me.txtFrom.TabIndex = 6
        '
        'txtAcc
        '
        Me.txtAcc.Location = New System.Drawing.Point(59, 49)
        Me.txtAcc.Name = "txtAcc"
        Me.txtAcc.Size = New System.Drawing.Size(123, 20)
        Me.txtAcc.TabIndex = 7
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(272, 287)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(75, 23)
        Me.btnNext.TabIndex = 8
        Me.btnNext.Text = "Next"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(435, 361)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "Form1"
        Me.Text = "SAMuel"
        CType(Me.picImage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents picImage As System.Windows.Forms.PictureBox
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnConvert As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents lblSubject As System.Windows.Forms.Label
    Friend WithEvents lblFrom As System.Windows.Forms.Label
    Friend WithEvents lblAcc As System.Windows.Forms.Label
    Friend WithEvents txtAcc As System.Windows.Forms.TextBox
    Friend WithEvents txtFrom As System.Windows.Forms.TextBox
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents btnNext As System.Windows.Forms.Button

End Class
