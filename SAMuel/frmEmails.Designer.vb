<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEmails
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
        Me.clbSelectedEmails = New System.Windows.Forms.CheckedListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'clbSelectedEmails
        '
        Me.clbSelectedEmails.AllowDrop = True
        Me.clbSelectedEmails.ForeColor = System.Drawing.SystemColors.WindowText
        Me.clbSelectedEmails.FormattingEnabled = True
        Me.clbSelectedEmails.Location = New System.Drawing.Point(-2, 42)
        Me.clbSelectedEmails.Name = "clbSelectedEmails"
        Me.clbSelectedEmails.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.clbSelectedEmails.Size = New System.Drawing.Size(387, 349)
        Me.clbSelectedEmails.TabIndex = 13
        Me.clbSelectedEmails.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(112, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(151, 20)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Emails Processed"
        '
        'frmEmails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(386, 388)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.clbSelectedEmails)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Location = New System.Drawing.Point(1000, 0)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEmails"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents clbSelectedEmails As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
