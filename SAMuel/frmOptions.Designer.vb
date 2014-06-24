<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.TabControl1.SuspendLayout()
        Me.tabGeneral.SuspendLayout()
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
        Me.tabOutlook.Location = New System.Drawing.Point(4, 22)
        Me.tabOutlook.Name = "tabOutlook"
        Me.tabOutlook.Padding = New System.Windows.Forms.Padding(3)
        Me.tabOutlook.Size = New System.Drawing.Size(498, 134)
        Me.tabOutlook.TabIndex = 1
        Me.tabOutlook.Text = "Outlook"
        Me.tabOutlook.UseVisualStyleBackColor = True
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
End Class
