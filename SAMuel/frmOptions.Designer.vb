﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptions
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
        Me.txtSavePath = New System.Windows.Forms.TextBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.btnSelectSaveFolder = New System.Windows.Forms.Button()
        Me.btnSaveOptions = New System.Windows.Forms.Button()
        Me.lblSaveTo = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tabGeneral = New System.Windows.Forms.TabPage()
        Me.tabOutlook = New System.Windows.Forms.TabPage()
        Me.grpTiffOptions = New System.Windows.Forms.GroupBox()
        Me.cmbColorDepth = New System.Windows.Forms.ComboBox()
        Me.cmbCompression = New System.Windows.Forms.ComboBox()
        Me.lblColorDepth = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.grpWatermarkOptions = New System.Windows.Forms.GroupBox()
        Me.lblFont = New System.Windows.Forms.Label()
        Me.cmbFont = New System.Windows.Forms.ComboBox()
        Me.TabControl1.SuspendLayout()
        Me.tabGeneral.SuspendLayout()
        Me.tabOutlook.SuspendLayout()
        Me.grpTiffOptions.SuspendLayout()
        Me.grpWatermarkOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtSavePath
        '
        Me.txtSavePath.Enabled = False
        Me.txtSavePath.Location = New System.Drawing.Point(62, 8)
        Me.txtSavePath.Name = "txtSavePath"
        Me.txtSavePath.Size = New System.Drawing.Size(346, 20)
        Me.txtSavePath.TabIndex = 0
        '
        'btnSelectSaveFolder
        '
        Me.btnSelectSaveFolder.Location = New System.Drawing.Point(414, 6)
        Me.btnSelectSaveFolder.Name = "btnSelectSaveFolder"
        Me.btnSelectSaveFolder.Size = New System.Drawing.Size(75, 23)
        Me.btnSelectSaveFolder.TabIndex = 1
        Me.btnSelectSaveFolder.Text = "Save To.."
        Me.btnSelectSaveFolder.UseVisualStyleBackColor = True
        '
        'btnSaveOptions
        '
        Me.btnSaveOptions.Location = New System.Drawing.Point(211, 199)
        Me.btnSaveOptions.Name = "btnSaveOptions"
        Me.btnSaveOptions.Size = New System.Drawing.Size(75, 23)
        Me.btnSaveOptions.TabIndex = 2
        Me.btnSaveOptions.Text = "Save"
        Me.btnSaveOptions.UseVisualStyleBackColor = True
        '
        'lblSaveTo
        '
        Me.lblSaveTo.AutoSize = True
        Me.lblSaveTo.Location = New System.Drawing.Point(2, 11)
        Me.lblSaveTo.Name = "lblSaveTo"
        Me.lblSaveTo.Size = New System.Drawing.Size(60, 13)
        Me.lblSaveTo.TabIndex = 3
        Me.lblSaveTo.Text = "Save Path:"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tabGeneral)
        Me.TabControl1.Controls.Add(Me.tabOutlook)
        Me.TabControl1.Location = New System.Drawing.Point(1, 1)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(506, 192)
        Me.TabControl1.TabIndex = 4
        '
        'tabGeneral
        '
        Me.tabGeneral.Controls.Add(Me.txtSavePath)
        Me.tabGeneral.Controls.Add(Me.lblSaveTo)
        Me.tabGeneral.Controls.Add(Me.btnSelectSaveFolder)
        Me.tabGeneral.Location = New System.Drawing.Point(4, 22)
        Me.tabGeneral.Name = "tabGeneral"
        Me.tabGeneral.Padding = New System.Windows.Forms.Padding(3)
        Me.tabGeneral.Size = New System.Drawing.Size(498, 166)
        Me.tabGeneral.TabIndex = 0
        Me.tabGeneral.Text = "General"
        Me.tabGeneral.UseVisualStyleBackColor = True
        '
        'tabOutlook
        '
        Me.tabOutlook.Controls.Add(Me.grpWatermarkOptions)
        Me.tabOutlook.Controls.Add(Me.grpTiffOptions)
        Me.tabOutlook.Location = New System.Drawing.Point(4, 22)
        Me.tabOutlook.Name = "tabOutlook"
        Me.tabOutlook.Padding = New System.Windows.Forms.Padding(3)
        Me.tabOutlook.Size = New System.Drawing.Size(498, 166)
        Me.tabOutlook.TabIndex = 1
        Me.tabOutlook.Text = "Outlook"
        Me.tabOutlook.UseVisualStyleBackColor = True
        '
        'grpTiffOptions
        '
        Me.grpTiffOptions.Controls.Add(Me.Label1)
        Me.grpTiffOptions.Controls.Add(Me.lblColorDepth)
        Me.grpTiffOptions.Controls.Add(Me.cmbCompression)
        Me.grpTiffOptions.Controls.Add(Me.cmbColorDepth)
        Me.grpTiffOptions.Location = New System.Drawing.Point(7, 72)
        Me.grpTiffOptions.Name = "grpTiffOptions"
        Me.grpTiffOptions.Size = New System.Drawing.Size(179, 88)
        Me.grpTiffOptions.TabIndex = 0
        Me.grpTiffOptions.TabStop = False
        Me.grpTiffOptions.Text = "Tiff Options"
        '
        'cmbColorDepth
        '
        Me.cmbColorDepth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbColorDepth.FormattingEnabled = True
        Me.cmbColorDepth.Items.AddRange(New Object() {"1-bit", "2-bit", "8-bit", "24-bit", "30-bit"})
        Me.cmbColorDepth.Location = New System.Drawing.Point(77, 19)
        Me.cmbColorDepth.Name = "cmbColorDepth"
        Me.cmbColorDepth.Size = New System.Drawing.Size(96, 21)
        Me.cmbColorDepth.TabIndex = 0
        '
        'cmbCompression
        '
        Me.cmbCompression.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCompression.FormattingEnabled = True
        Me.cmbCompression.Items.AddRange(New Object() {"LZW", "CCITT3", "CCITT4", "RLE", "None"})
        Me.cmbCompression.Location = New System.Drawing.Point(77, 46)
        Me.cmbCompression.Name = "cmbCompression"
        Me.cmbCompression.Size = New System.Drawing.Size(96, 21)
        Me.cmbCompression.TabIndex = 1
        '
        'lblColorDepth
        '
        Me.lblColorDepth.AutoSize = True
        Me.lblColorDepth.Location = New System.Drawing.Point(8, 23)
        Me.lblColorDepth.Name = "lblColorDepth"
        Me.lblColorDepth.Size = New System.Drawing.Size(66, 13)
        Me.lblColorDepth.TabIndex = 2
        Me.lblColorDepth.Text = "Color Depth:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 49)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Compression: "
        '
        'grpWatermarkOptions
        '
        Me.grpWatermarkOptions.Controls.Add(Me.lblFont)
        Me.grpWatermarkOptions.Controls.Add(Me.cmbFont)
        Me.grpWatermarkOptions.Location = New System.Drawing.Point(192, 72)
        Me.grpWatermarkOptions.Name = "grpWatermarkOptions"
        Me.grpWatermarkOptions.Size = New System.Drawing.Size(179, 88)
        Me.grpWatermarkOptions.TabIndex = 4
        Me.grpWatermarkOptions.TabStop = False
        Me.grpWatermarkOptions.Text = "Watermark Options"
        '
        'lblFont
        '
        Me.lblFont.AutoSize = True
        Me.lblFont.Location = New System.Drawing.Point(46, 19)
        Me.lblFont.Name = "lblFont"
        Me.lblFont.Size = New System.Drawing.Size(31, 13)
        Me.lblFont.TabIndex = 2
        Me.lblFont.Text = "Font:"
        '
        'cmbFont
        '
        Me.cmbFont.FormattingEnabled = True
        Me.cmbFont.Items.AddRange(New Object() {"Courier New", "Times New Roman", "fixed sys"})
        Me.cmbFont.Location = New System.Drawing.Point(77, 15)
        Me.cmbFont.Name = "cmbFont"
        Me.cmbFont.Size = New System.Drawing.Size(96, 21)
        Me.cmbFont.TabIndex = 0
        '
        'frmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(506, 227)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnSaveOptions)
        Me.Name = "frmOptions"
        Me.Text = "Options"
        Me.TabControl1.ResumeLayout(False)
        Me.tabGeneral.ResumeLayout(False)
        Me.tabGeneral.PerformLayout()
        Me.tabOutlook.ResumeLayout(False)
        Me.grpTiffOptions.ResumeLayout(False)
        Me.grpTiffOptions.PerformLayout()
        Me.grpWatermarkOptions.ResumeLayout(False)
        Me.grpWatermarkOptions.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtSavePath As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents btnSelectSaveFolder As System.Windows.Forms.Button
    Friend WithEvents btnSaveOptions As System.Windows.Forms.Button
    Friend WithEvents lblSaveTo As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tabGeneral As System.Windows.Forms.TabPage
    Friend WithEvents tabOutlook As System.Windows.Forms.TabPage
    Friend WithEvents grpTiffOptions As System.Windows.Forms.GroupBox
    Friend WithEvents lblColorDepth As System.Windows.Forms.Label
    Friend WithEvents cmbCompression As System.Windows.Forms.ComboBox
    Friend WithEvents cmbColorDepth As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grpWatermarkOptions As System.Windows.Forms.GroupBox
    Friend WithEvents lblFont As System.Windows.Forms.Label
    Friend WithEvents cmbFont As System.Windows.Forms.ComboBox
End Class
