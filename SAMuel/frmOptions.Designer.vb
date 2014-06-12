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
        Me.SuspendLayout()
        '
        'txtSavePath
        '
        Me.txtSavePath.Enabled = False
        Me.txtSavePath.Location = New System.Drawing.Point(68, 36)
        Me.txtSavePath.Name = "txtSavePath"
        Me.txtSavePath.Size = New System.Drawing.Size(346, 20)
        Me.txtSavePath.TabIndex = 0
        '
        'btnSelectSaveFolder
        '
        Me.btnSelectSaveFolder.Location = New System.Drawing.Point(420, 34)
        Me.btnSelectSaveFolder.Name = "btnSelectSaveFolder"
        Me.btnSelectSaveFolder.Size = New System.Drawing.Size(75, 23)
        Me.btnSelectSaveFolder.TabIndex = 1
        Me.btnSelectSaveFolder.Text = "Save To.."
        Me.btnSelectSaveFolder.UseVisualStyleBackColor = True
        '
        'btnSaveOptions
        '
        Me.btnSaveOptions.Location = New System.Drawing.Point(188, 192)
        Me.btnSaveOptions.Name = "btnSaveOptions"
        Me.btnSaveOptions.Size = New System.Drawing.Size(75, 23)
        Me.btnSaveOptions.TabIndex = 2
        Me.btnSaveOptions.Text = "Save"
        Me.btnSaveOptions.UseVisualStyleBackColor = True
        '
        'lblSaveTo
        '
        Me.lblSaveTo.AutoSize = True
        Me.lblSaveTo.Location = New System.Drawing.Point(8, 39)
        Me.lblSaveTo.Name = "lblSaveTo"
        Me.lblSaveTo.Size = New System.Drawing.Size(60, 13)
        Me.lblSaveTo.TabIndex = 3
        Me.lblSaveTo.Text = "Save Path:"
        '
        'frmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(505, 227)
        Me.Controls.Add(Me.lblSaveTo)
        Me.Controls.Add(Me.btnSaveOptions)
        Me.Controls.Add(Me.btnSelectSaveFolder)
        Me.Controls.Add(Me.txtSavePath)
        Me.Name = "frmOptions"
        Me.Text = "Options"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtSavePath As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents btnSelectSaveFolder As System.Windows.Forms.Button
    Friend WithEvents btnSaveOptions As System.Windows.Forms.Button
    Friend WithEvents lblSaveTo As System.Windows.Forms.Label
End Class
