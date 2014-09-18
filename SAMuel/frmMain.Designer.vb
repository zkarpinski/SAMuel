<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container()
        Dim ListViewItem1 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"test", "test", "test"}, -1)
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lstEmailAttachments = New System.Windows.Forms.ListView()
        Me.hType = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.hName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.chkValidOnly = New System.Windows.Forms.CheckBox()
        Me.groupOLAudit = New System.Windows.Forms.GroupBox()
        Me.rtbEmailBody = New System.Windows.Forms.RichTextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtFrom = New System.Windows.Forms.TextBox()
        Me.lblSubject = New System.Windows.Forms.Label()
        Me.lblFrom = New System.Windows.Forms.Label()
        Me.lblOutlookMessage = New System.Windows.Forms.Label()
        Me.lblAcc = New System.Windows.Forms.Label()
        Me.txtSubject = New System.Windows.Forms.TextBox()
        Me.txtAcc = New System.Windows.Forms.TextBox()
        Me.btnReject = New System.Windows.Forms.Button()
        Me.chkAuditMode = New System.Windows.Forms.CheckBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.tabWordToTiff = New System.Windows.Forms.TabPage()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.rbConvertDOC = New System.Windows.Forms.RadioButton()
        Me.rbConvertIMAGE = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.rbConvertTiff = New System.Windows.Forms.RadioButton()
        Me.rbConvertPDF = New System.Windows.Forms.RadioButton()
        Me.lblDragAndDropWord = New System.Windows.Forms.Label()
        Me.btnConvert = New System.Windows.Forms.Button()
        Me.tabRFax = New System.Windows.Forms.TabPage()
        Me.chkRFSaveRec = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkRFCoverSheet = New System.Windows.Forms.CheckBox()
        Me.txtRFRecFax = New System.Windows.Forms.TextBox()
        Me.lblRFRecFax = New System.Windows.Forms.Label()
        Me.txtRFRecName = New System.Windows.Forms.TextBox()
        Me.lblRFRecName = New System.Windows.Forms.Label()
        Me.grpRFServer = New System.Windows.Forms.GroupBox()
        Me.chkRFNTauth = New System.Windows.Forms.CheckBox()
        Me.lblRFpassword = New System.Windows.Forms.Label()
        Me.txtRFsvr = New System.Windows.Forms.TextBox()
        Me.txtRFpw = New System.Windows.Forms.TextBox()
        Me.lblRFserver = New System.Windows.Forms.Label()
        Me.txtRFuser = New System.Windows.Forms.TextBox()
        Me.lblRFuser = New System.Windows.Forms.Label()
        Me.btnRFax = New System.Windows.Forms.Button()
        Me.tabKofax = New System.Windows.Forms.TabPage()
        Me.lblKFComments = New System.Windows.Forms.Label()
        Me.txtKFComments = New System.Windows.Forms.TextBox()
        Me.gbKFSource = New System.Windows.Forms.GroupBox()
        Me.rbKFUSMail = New System.Windows.Forms.RadioButton()
        Me.rbKFEmail = New System.Windows.Forms.RadioButton()
        Me.lblKFBatchType = New System.Windows.Forms.Label()
        Me.cbKFBatchType = New System.Windows.Forms.ComboBox()
        Me.btnKFImport = New System.Windows.Forms.Button()
        Me.lblKFBatchName = New System.Windows.Forms.Label()
        Me.txtKFBatchName = New System.Windows.Forms.TextBox()
        Me.tabAddContact = New System.Windows.Forms.TabPage()
        Me.btnStopAddContacts = New System.Windows.Forms.Button()
        Me.lbxContactAlerts = New System.Windows.Forms.ListBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.rtbCContact = New System.Windows.Forms.RichTextBox()
        Me.mtxtCAccount = New System.Windows.Forms.MaskedTextBox()
        Me.lblCAccount = New System.Windows.Forms.Label()
        Me.btnCAddContact = New System.Windows.Forms.Button()
        Me.tabDPA = New System.Windows.Forms.TabPage()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.btnBudgetBill = New System.Windows.Forms.Button()
        Me.mtxtDPAAcc = New System.Windows.Forms.MaskedTextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDPAmonthly = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDPAdown = New System.Windows.Forms.TextBox()
        Me.btnDPAprocess = New System.Windows.Forms.Button()
        Me.chkMinPayment = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tabTDrive = New System.Windows.Forms.TabPage()
        Me.btnTDClear = New System.Windows.Forms.Button()
        Me.btnTDCreateEmail = New System.Windows.Forms.Button()
        Me.lvTDriveFiles = New System.Windows.Forms.ListView()
        Me.colType = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colSendTo = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colAcc = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.lblBranding = New System.Windows.Forms.Label()
        Me.timerConvert = New System.Windows.Forms.Timer(Me.components)
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.groupOLAudit.SuspendLayout()
        Me.tabWordToTiff.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.tabRFax.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.grpRFServer.SuspendLayout()
        Me.tabKofax.SuspendLayout()
        Me.gbKFSource.SuspendLayout()
        Me.tabAddContact.SuspendLayout()
        Me.tabDPA.SuspendLayout()
        Me.tabTDrive.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dlgOpen
        '
        Me.dlgOpen.Multiselect = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.tabWordToTiff)
        Me.TabControl1.Controls.Add(Me.tabRFax)
        Me.TabControl1.Controls.Add(Me.tabKofax)
        Me.TabControl1.Controls.Add(Me.tabAddContact)
        Me.TabControl1.Controls.Add(Me.tabDPA)
        Me.TabControl1.Controls.Add(Me.tabTDrive)
        Me.TabControl1.Location = New System.Drawing.Point(-1, 27)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(439, 316)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.BackgroundImage = Global.SAMuel.My.Resources.Resources.ny_map_f
        Me.TabPage1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.TabPage1.Controls.Add(Me.lstEmailAttachments)
        Me.TabPage1.Controls.Add(Me.chkValidOnly)
        Me.TabPage1.Controls.Add(Me.groupOLAudit)
        Me.TabPage1.Controls.Add(Me.btnReject)
        Me.TabPage1.Controls.Add(Me.chkAuditMode)
        Me.TabPage1.Controls.Add(Me.btnCancel)
        Me.TabPage1.Controls.Add(Me.btnNext)
        Me.TabPage1.Controls.Add(Me.btnRun)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(431, 290)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Outlook"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lstEmailAttachments
        '
        Me.lstEmailAttachments.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.hType, Me.hName})
        Me.lstEmailAttachments.FullRowSelect = True
        ListViewItem1.StateImageIndex = 0
        Me.lstEmailAttachments.Items.AddRange(New System.Windows.Forms.ListViewItem() {ListViewItem1})
        Me.lstEmailAttachments.Location = New System.Drawing.Point(229, 7)
        Me.lstEmailAttachments.Name = "lstEmailAttachments"
        Me.lstEmailAttachments.Size = New System.Drawing.Size(196, 254)
        Me.lstEmailAttachments.TabIndex = 18
        Me.lstEmailAttachments.UseCompatibleStateImageBehavior = False
        Me.lstEmailAttachments.View = System.Windows.Forms.View.Details
        '
        'hType
        '
        Me.hType.Text = "Type"
        Me.hType.Width = 46
        '
        'hName
        '
        Me.hName.Text = "Filename"
        Me.hName.Width = 146
        '
        'chkValidOnly
        '
        Me.chkValidOnly.AutoSize = True
        Me.chkValidOnly.Location = New System.Drawing.Point(91, 270)
        Me.chkValidOnly.Name = "chkValidOnly"
        Me.chkValidOnly.Size = New System.Drawing.Size(73, 17)
        Me.chkValidOnly.TabIndex = 17
        Me.chkValidOnly.Text = "Valid Only"
        Me.chkValidOnly.UseVisualStyleBackColor = True
        '
        'groupOLAudit
        '
        Me.groupOLAudit.BackColor = System.Drawing.Color.White
        Me.groupOLAudit.Controls.Add(Me.rtbEmailBody)
        Me.groupOLAudit.Controls.Add(Me.Label5)
        Me.groupOLAudit.Controls.Add(Me.txtFrom)
        Me.groupOLAudit.Controls.Add(Me.lblSubject)
        Me.groupOLAudit.Controls.Add(Me.lblFrom)
        Me.groupOLAudit.Controls.Add(Me.lblOutlookMessage)
        Me.groupOLAudit.Controls.Add(Me.lblAcc)
        Me.groupOLAudit.Controls.Add(Me.txtSubject)
        Me.groupOLAudit.Controls.Add(Me.txtAcc)
        Me.groupOLAudit.Location = New System.Drawing.Point(4, 1)
        Me.groupOLAudit.Name = "groupOLAudit"
        Me.groupOLAudit.Size = New System.Drawing.Size(219, 260)
        Me.groupOLAudit.TabIndex = 16
        Me.groupOLAudit.TabStop = False
        Me.groupOLAudit.Text = "Email"
        '
        'rtbEmailBody
        '
        Me.rtbEmailBody.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rtbEmailBody.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbEmailBody.Location = New System.Drawing.Point(-1, 106)
        Me.rtbEmailBody.Name = "rtbEmailBody"
        Me.rtbEmailBody.ReadOnly = True
        Me.rtbEmailBody.Size = New System.Drawing.Size(219, 154)
        Me.rtbEmailBody.TabIndex = 16
        Me.rtbEmailBody.Text = "This is a sample email body. With multiple lines and jibberish..." & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10) & "Ctrl+Scroll " & _
    "changes font size!" & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10) & "Scroll down to see this!."
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(2, 94)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(34, 13)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Body:"
        '
        'txtFrom
        '
        Me.txtFrom.Location = New System.Drawing.Point(55, 34)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.ReadOnly = True
        Me.txtFrom.Size = New System.Drawing.Size(158, 20)
        Me.txtFrom.TabIndex = 6
        '
        'lblSubject
        '
        Me.lblSubject.AutoSize = True
        Me.lblSubject.BackColor = System.Drawing.Color.Transparent
        Me.lblSubject.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSubject.Location = New System.Drawing.Point(2, 14)
        Me.lblSubject.Name = "lblSubject"
        Me.lblSubject.Size = New System.Drawing.Size(46, 13)
        Me.lblSubject.TabIndex = 2
        Me.lblSubject.Text = "Subject:"
        '
        'lblFrom
        '
        Me.lblFrom.AutoSize = True
        Me.lblFrom.BackColor = System.Drawing.Color.Transparent
        Me.lblFrom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFrom.Location = New System.Drawing.Point(2, 37)
        Me.lblFrom.Name = "lblFrom"
        Me.lblFrom.Size = New System.Drawing.Size(44, 13)
        Me.lblFrom.TabIndex = 3
        Me.lblFrom.Text = "Sender:"
        '
        'lblOutlookMessage
        '
        Me.lblOutlookMessage.AutoSize = True
        Me.lblOutlookMessage.BackColor = System.Drawing.Color.White
        Me.lblOutlookMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOutlookMessage.ForeColor = System.Drawing.Color.Red
        Me.lblOutlookMessage.Location = New System.Drawing.Point(50, 84)
        Me.lblOutlookMessage.Name = "lblOutlookMessage"
        Me.lblOutlookMessage.Size = New System.Drawing.Size(125, 13)
        Me.lblOutlookMessage.TabIndex = 15
        Me.lblOutlookMessage.Text = "Email Error/Warning!"
        '
        'lblAcc
        '
        Me.lblAcc.AutoSize = True
        Me.lblAcc.BackColor = System.Drawing.Color.Transparent
        Me.lblAcc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAcc.Location = New System.Drawing.Point(2, 59)
        Me.lblAcc.Name = "lblAcc"
        Me.lblAcc.Size = New System.Drawing.Size(53, 13)
        Me.lblAcc.TabIndex = 4
        Me.lblAcc.Text = "Account: "
        '
        'txtSubject
        '
        Me.txtSubject.Location = New System.Drawing.Point(55, 11)
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.ReadOnly = True
        Me.txtSubject.Size = New System.Drawing.Size(158, 20)
        Me.txtSubject.TabIndex = 5
        '
        'txtAcc
        '
        Me.txtAcc.Location = New System.Drawing.Point(55, 56)
        Me.txtAcc.Name = "txtAcc"
        Me.txtAcc.Size = New System.Drawing.Size(158, 20)
        Me.txtAcc.TabIndex = 7
        '
        'btnReject
        '
        Me.btnReject.Enabled = False
        Me.btnReject.Location = New System.Drawing.Point(265, 267)
        Me.btnReject.Name = "btnReject"
        Me.btnReject.Size = New System.Drawing.Size(75, 23)
        Me.btnReject.TabIndex = 8
        Me.btnReject.Text = "Reject"
        Me.btnReject.UseVisualStyleBackColor = True
        '
        'chkAuditMode
        '
        Me.chkAuditMode.AutoSize = True
        Me.chkAuditMode.Location = New System.Drawing.Point(4, 270)
        Me.chkAuditMode.Name = "chkAuditMode"
        Me.chkAuditMode.Size = New System.Drawing.Size(80, 17)
        Me.chkAuditMode.TabIndex = 14
        Me.chkAuditMode.Text = "Audit Mode"
        Me.chkAuditMode.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(353, 267)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 11
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(177, 267)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(75, 23)
        Me.btnNext.TabIndex = 10
        Me.btnNext.Text = "&Next"
        Me.btnNext.UseVisualStyleBackColor = True
        Me.btnNext.Visible = False
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(177, 267)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(75, 23)
        Me.btnRun.TabIndex = 0
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'tabWordToTiff
        '
        Me.tabWordToTiff.AllowDrop = True
        Me.tabWordToTiff.Controls.Add(Me.GroupBox4)
        Me.tabWordToTiff.Controls.Add(Me.GroupBox3)
        Me.tabWordToTiff.Controls.Add(Me.lblDragAndDropWord)
        Me.tabWordToTiff.Controls.Add(Me.btnConvert)
        Me.tabWordToTiff.Location = New System.Drawing.Point(4, 22)
        Me.tabWordToTiff.Name = "tabWordToTiff"
        Me.tabWordToTiff.Padding = New System.Windows.Forms.Padding(3)
        Me.tabWordToTiff.Size = New System.Drawing.Size(431, 290)
        Me.tabWordToTiff.TabIndex = 1
        Me.tabWordToTiff.Text = "Convert Files"
        Me.tabWordToTiff.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.rbConvertDOC)
        Me.GroupBox4.Controls.Add(Me.rbConvertIMAGE)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(6, 223)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(81, 66)
        Me.GroupBox4.TabIndex = 7
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Input"
        '
        'rbConvertDOC
        '
        Me.rbConvertDOC.AutoSize = True
        Me.rbConvertDOC.Checked = True
        Me.rbConvertDOC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbConvertDOC.Location = New System.Drawing.Point(6, 19)
        Me.rbConvertDOC.Name = "rbConvertDOC"
        Me.rbConvertDOC.Size = New System.Drawing.Size(74, 17)
        Me.rbConvertDOC.TabIndex = 5
        Me.rbConvertDOC.TabStop = True
        Me.rbConvertDOC.Text = "Word Doc"
        Me.rbConvertDOC.UseVisualStyleBackColor = True
        '
        'rbConvertIMAGE
        '
        Me.rbConvertIMAGE.AutoSize = True
        Me.rbConvertIMAGE.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbConvertIMAGE.Location = New System.Drawing.Point(6, 42)
        Me.rbConvertIMAGE.Name = "rbConvertIMAGE"
        Me.rbConvertIMAGE.Size = New System.Drawing.Size(54, 17)
        Me.rbConvertIMAGE.TabIndex = 4
        Me.rbConvertIMAGE.Text = "Image"
        Me.rbConvertIMAGE.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.rbConvertTiff)
        Me.GroupBox3.Controls.Add(Me.rbConvertPDF)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(87, 223)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(75, 66)
        Me.GroupBox3.TabIndex = 6
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Output"
        '
        'rbConvertTiff
        '
        Me.rbConvertTiff.AutoSize = True
        Me.rbConvertTiff.Checked = True
        Me.rbConvertTiff.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbConvertTiff.Location = New System.Drawing.Point(6, 19)
        Me.rbConvertTiff.Name = "rbConvertTiff"
        Me.rbConvertTiff.Size = New System.Drawing.Size(40, 17)
        Me.rbConvertTiff.TabIndex = 5
        Me.rbConvertTiff.TabStop = True
        Me.rbConvertTiff.Text = "Tiff"
        Me.rbConvertTiff.UseVisualStyleBackColor = True
        '
        'rbConvertPDF
        '
        Me.rbConvertPDF.AutoSize = True
        Me.rbConvertPDF.Enabled = False
        Me.rbConvertPDF.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbConvertPDF.Location = New System.Drawing.Point(6, 42)
        Me.rbConvertPDF.Name = "rbConvertPDF"
        Me.rbConvertPDF.Size = New System.Drawing.Size(46, 17)
        Me.rbConvertPDF.TabIndex = 4
        Me.rbConvertPDF.Text = "PDF"
        Me.rbConvertPDF.UseVisualStyleBackColor = True
        '
        'lblDragAndDropWord
        '
        Me.lblDragAndDropWord.AutoSize = True
        Me.lblDragAndDropWord.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDragAndDropWord.Location = New System.Drawing.Point(9, 129)
        Me.lblDragAndDropWord.Name = "lblDragAndDropWord"
        Me.lblDragAndDropWord.Size = New System.Drawing.Size(416, 13)
        Me.lblDragAndDropWord.TabIndex = 3
        Me.lblDragAndDropWord.Text = "Click Convert or Drag and Drop files to convert into the selected format."
        '
        'btnConvert
        '
        Me.btnConvert.Location = New System.Drawing.Point(177, 265)
        Me.btnConvert.Name = "btnConvert"
        Me.btnConvert.Size = New System.Drawing.Size(75, 23)
        Me.btnConvert.TabIndex = 2
        Me.btnConvert.Text = "Convert"
        Me.btnConvert.UseVisualStyleBackColor = True
        '
        'tabRFax
        '
        Me.tabRFax.Controls.Add(Me.chkRFSaveRec)
        Me.tabRFax.Controls.Add(Me.GroupBox2)
        Me.tabRFax.Controls.Add(Me.grpRFServer)
        Me.tabRFax.Controls.Add(Me.btnRFax)
        Me.tabRFax.Location = New System.Drawing.Point(4, 22)
        Me.tabRFax.Name = "tabRFax"
        Me.tabRFax.Size = New System.Drawing.Size(431, 290)
        Me.tabRFax.TabIndex = 4
        Me.tabRFax.Text = "RightFax"
        Me.tabRFax.UseVisualStyleBackColor = True
        '
        'chkRFSaveRec
        '
        Me.chkRFSaveRec.AutoSize = True
        Me.chkRFSaveRec.Location = New System.Drawing.Point(7, 267)
        Me.chkRFSaveRec.Name = "chkRFSaveRec"
        Me.chkRFSaveRec.Size = New System.Drawing.Size(152, 17)
        Me.chkRFSaveRec.TabIndex = 13
        Me.chkRFSaveRec.Text = "Save recipient info for later"
        Me.chkRFSaveRec.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkRFCoverSheet)
        Me.GroupBox2.Controls.Add(Me.txtRFRecFax)
        Me.GroupBox2.Controls.Add(Me.lblRFRecFax)
        Me.GroupBox2.Controls.Add(Me.txtRFRecName)
        Me.GroupBox2.Controls.Add(Me.lblRFRecName)
        Me.GroupBox2.Location = New System.Drawing.Point(19, 104)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(381, 95)
        Me.GroupBox2.TabIndex = 12
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Fax Info"
        '
        'chkRFCoverSheet
        '
        Me.chkRFCoverSheet.AutoSize = True
        Me.chkRFCoverSheet.Checked = True
        Me.chkRFCoverSheet.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRFCoverSheet.Location = New System.Drawing.Point(17, 78)
        Me.chkRFCoverSheet.Name = "chkRFCoverSheet"
        Me.chkRFCoverSheet.Size = New System.Drawing.Size(123, 17)
        Me.chkRFCoverSheet.TabIndex = 11
        Me.chkRFCoverSheet.Text = "Include Cover Sheet"
        Me.chkRFCoverSheet.UseVisualStyleBackColor = True
        Me.chkRFCoverSheet.Visible = False
        '
        'txtRFRecFax
        '
        Me.txtRFRecFax.Location = New System.Drawing.Point(98, 48)
        Me.txtRFRecFax.Name = "txtRFRecFax"
        Me.txtRFRecFax.Size = New System.Drawing.Size(100, 20)
        Me.txtRFRecFax.TabIndex = 6
        '
        'lblRFRecFax
        '
        Me.lblRFRecFax.AutoSize = True
        Me.lblRFRecFax.Location = New System.Drawing.Point(6, 50)
        Me.lblRFRecFax.Name = "lblRFRecFax"
        Me.lblRFRecFax.Size = New System.Drawing.Size(75, 13)
        Me.lblRFRecFax.TabIndex = 7
        Me.lblRFRecFax.Text = "Recipient Fax:"
        '
        'txtRFRecName
        '
        Me.txtRFRecName.Location = New System.Drawing.Point(98, 22)
        Me.txtRFRecName.Name = "txtRFRecName"
        Me.txtRFRecName.Size = New System.Drawing.Size(100, 20)
        Me.txtRFRecName.TabIndex = 4
        '
        'lblRFRecName
        '
        Me.lblRFRecName.AutoSize = True
        Me.lblRFRecName.Location = New System.Drawing.Point(6, 25)
        Me.lblRFRecName.Name = "lblRFRecName"
        Me.lblRFRecName.Size = New System.Drawing.Size(86, 13)
        Me.lblRFRecName.TabIndex = 5
        Me.lblRFRecName.Text = "Recipient Name:"
        '
        'grpRFServer
        '
        Me.grpRFServer.Controls.Add(Me.chkRFNTauth)
        Me.grpRFServer.Controls.Add(Me.lblRFpassword)
        Me.grpRFServer.Controls.Add(Me.txtRFsvr)
        Me.grpRFServer.Controls.Add(Me.txtRFpw)
        Me.grpRFServer.Controls.Add(Me.lblRFserver)
        Me.grpRFServer.Controls.Add(Me.txtRFuser)
        Me.grpRFServer.Controls.Add(Me.lblRFuser)
        Me.grpRFServer.Location = New System.Drawing.Point(19, 3)
        Me.grpRFServer.Name = "grpRFServer"
        Me.grpRFServer.Size = New System.Drawing.Size(381, 95)
        Me.grpRFServer.TabIndex = 11
        Me.grpRFServer.TabStop = False
        Me.grpRFServer.Text = "Server Info"
        '
        'chkRFNTauth
        '
        Me.chkRFNTauth.AutoSize = True
        Me.chkRFNTauth.Location = New System.Drawing.Point(22, 65)
        Me.chkRFNTauth.Name = "chkRFNTauth"
        Me.chkRFNTauth.Size = New System.Drawing.Size(134, 17)
        Me.chkRFNTauth.TabIndex = 8
        Me.chkRFNTauth.Text = "Use NT Authentication"
        Me.chkRFNTauth.UseVisualStyleBackColor = True
        Me.chkRFNTauth.Visible = False
        '
        'lblRFpassword
        '
        Me.lblRFpassword.AutoSize = True
        Me.lblRFpassword.Location = New System.Drawing.Point(193, 51)
        Me.lblRFpassword.Name = "lblRFpassword"
        Me.lblRFpassword.Size = New System.Drawing.Size(56, 13)
        Me.lblRFpassword.TabIndex = 10
        Me.lblRFpassword.Text = "Password:"
        '
        'txtRFsvr
        '
        Me.txtRFsvr.Location = New System.Drawing.Point(64, 22)
        Me.txtRFsvr.Name = "txtRFsvr"
        Me.txtRFsvr.Size = New System.Drawing.Size(100, 20)
        Me.txtRFsvr.TabIndex = 4
        '
        'txtRFpw
        '
        Me.txtRFpw.Location = New System.Drawing.Point(255, 48)
        Me.txtRFpw.Name = "txtRFpw"
        Me.txtRFpw.Size = New System.Drawing.Size(100, 20)
        Me.txtRFpw.TabIndex = 9
        '
        'lblRFserver
        '
        Me.lblRFserver.AutoSize = True
        Me.lblRFserver.Location = New System.Drawing.Point(19, 25)
        Me.lblRFserver.Name = "lblRFserver"
        Me.lblRFserver.Size = New System.Drawing.Size(41, 13)
        Me.lblRFserver.TabIndex = 5
        Me.lblRFserver.Text = "Server:"
        '
        'txtRFuser
        '
        Me.txtRFuser.Location = New System.Drawing.Point(255, 22)
        Me.txtRFuser.Name = "txtRFuser"
        Me.txtRFuser.Size = New System.Drawing.Size(100, 20)
        Me.txtRFuser.TabIndex = 6
        '
        'lblRFuser
        '
        Me.lblRFuser.AutoSize = True
        Me.lblRFuser.Location = New System.Drawing.Point(191, 25)
        Me.lblRFuser.Name = "lblRFuser"
        Me.lblRFuser.Size = New System.Drawing.Size(58, 13)
        Me.lblRFuser.TabIndex = 7
        Me.lblRFuser.Text = "Username:"
        '
        'btnRFax
        '
        Me.btnRFax.Location = New System.Drawing.Point(177, 265)
        Me.btnRFax.Name = "btnRFax"
        Me.btnRFax.Size = New System.Drawing.Size(75, 23)
        Me.btnRFax.TabIndex = 3
        Me.btnRFax.Text = "Fax"
        Me.btnRFax.UseVisualStyleBackColor = True
        '
        'tabKofax
        '
        Me.tabKofax.Controls.Add(Me.lblKFComments)
        Me.tabKofax.Controls.Add(Me.txtKFComments)
        Me.tabKofax.Controls.Add(Me.gbKFSource)
        Me.tabKofax.Controls.Add(Me.lblKFBatchType)
        Me.tabKofax.Controls.Add(Me.cbKFBatchType)
        Me.tabKofax.Controls.Add(Me.btnKFImport)
        Me.tabKofax.Controls.Add(Me.lblKFBatchName)
        Me.tabKofax.Controls.Add(Me.txtKFBatchName)
        Me.tabKofax.Location = New System.Drawing.Point(4, 22)
        Me.tabKofax.Name = "tabKofax"
        Me.tabKofax.Padding = New System.Windows.Forms.Padding(3)
        Me.tabKofax.Size = New System.Drawing.Size(431, 290)
        Me.tabKofax.TabIndex = 5
        Me.tabKofax.Text = "Kofax It"
        Me.tabKofax.UseVisualStyleBackColor = True
        '
        'lblKFComments
        '
        Me.lblKFComments.AutoSize = True
        Me.lblKFComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKFComments.Location = New System.Drawing.Point(12, 94)
        Me.lblKFComments.Name = "lblKFComments"
        Me.lblKFComments.Size = New System.Drawing.Size(59, 13)
        Me.lblKFComments.TabIndex = 19
        Me.lblKFComments.Text = "Comments:"
        '
        'txtKFComments
        '
        Me.txtKFComments.Location = New System.Drawing.Point(87, 89)
        Me.txtKFComments.Name = "txtKFComments"
        Me.txtKFComments.Size = New System.Drawing.Size(165, 20)
        Me.txtKFComments.TabIndex = 18
        '
        'gbKFSource
        '
        Me.gbKFSource.Controls.Add(Me.rbKFUSMail)
        Me.gbKFSource.Controls.Add(Me.rbKFEmail)
        Me.gbKFSource.Location = New System.Drawing.Point(273, 25)
        Me.gbKFSource.Name = "gbKFSource"
        Me.gbKFSource.Size = New System.Drawing.Size(106, 58)
        Me.gbKFSource.TabIndex = 17
        Me.gbKFSource.TabStop = False
        Me.gbKFSource.Text = "Source"
        '
        'rbKFUSMail
        '
        Me.rbKFUSMail.AutoSize = True
        Me.rbKFUSMail.Location = New System.Drawing.Point(7, 38)
        Me.rbKFUSMail.Name = "rbKFUSMail"
        Me.rbKFUSMail.Size = New System.Drawing.Size(62, 17)
        Me.rbKFUSMail.TabIndex = 1
        Me.rbKFUSMail.Text = "US Mail"
        Me.rbKFUSMail.UseVisualStyleBackColor = True
        '
        'rbKFEmail
        '
        Me.rbKFEmail.AutoSize = True
        Me.rbKFEmail.Checked = True
        Me.rbKFEmail.Location = New System.Drawing.Point(7, 20)
        Me.rbKFEmail.Name = "rbKFEmail"
        Me.rbKFEmail.Size = New System.Drawing.Size(50, 17)
        Me.rbKFEmail.TabIndex = 0
        Me.rbKFEmail.TabStop = True
        Me.rbKFEmail.Text = "Email"
        Me.rbKFEmail.UseVisualStyleBackColor = True
        '
        'lblKFBatchType
        '
        Me.lblKFBatchType.AutoSize = True
        Me.lblKFBatchType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKFBatchType.Location = New System.Drawing.Point(12, 51)
        Me.lblKFBatchType.Name = "lblKFBatchType"
        Me.lblKFBatchType.Size = New System.Drawing.Size(34, 13)
        Me.lblKFBatchType.TabIndex = 16
        Me.lblKFBatchType.Text = "Type:"
        '
        'cbKFBatchType
        '
        Me.cbKFBatchType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbKFBatchType.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbKFBatchType.FormattingEnabled = True
        Me.cbKFBatchType.Items.AddRange(New Object() {"0003 - BILLING DISPUTE", "0014 - DENIAL OF SERVICE", "0084 - DPA-SYRACUSE PROCESSING", "0026 - LIFE SUPPORT", "0049 - MEDICAL-INFO ONLY", "0046 - NO ACTION-INFO ONLY", "0033 - SAMS"})
        Me.cbKFBatchType.Location = New System.Drawing.Point(87, 48)
        Me.cbKFBatchType.Name = "cbKFBatchType"
        Me.cbKFBatchType.Size = New System.Drawing.Size(165, 20)
        Me.cbKFBatchType.TabIndex = 15
        '
        'btnKFImport
        '
        Me.btnKFImport.Location = New System.Drawing.Point(177, 265)
        Me.btnKFImport.Name = "btnKFImport"
        Me.btnKFImport.Size = New System.Drawing.Size(75, 23)
        Me.btnKFImport.TabIndex = 14
        Me.btnKFImport.Text = "Import"
        Me.btnKFImport.UseVisualStyleBackColor = True
        '
        'lblKFBatchName
        '
        Me.lblKFBatchName.AutoSize = True
        Me.lblKFBatchName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKFBatchName.Location = New System.Drawing.Point(12, 25)
        Me.lblKFBatchName.Name = "lblKFBatchName"
        Me.lblKFBatchName.Size = New System.Drawing.Size(69, 13)
        Me.lblKFBatchName.TabIndex = 13
        Me.lblKFBatchName.Text = "Batch Name:"
        '
        'txtKFBatchName
        '
        Me.txtKFBatchName.Location = New System.Drawing.Point(87, 22)
        Me.txtKFBatchName.Name = "txtKFBatchName"
        Me.txtKFBatchName.Size = New System.Drawing.Size(165, 20)
        Me.txtKFBatchName.TabIndex = 12
        '
        'tabAddContact
        '
        Me.tabAddContact.AllowDrop = True
        Me.tabAddContact.Controls.Add(Me.btnStopAddContacts)
        Me.tabAddContact.Controls.Add(Me.lbxContactAlerts)
        Me.tabAddContact.Controls.Add(Me.Label4)
        Me.tabAddContact.Controls.Add(Me.rtbCContact)
        Me.tabAddContact.Controls.Add(Me.mtxtCAccount)
        Me.tabAddContact.Controls.Add(Me.lblCAccount)
        Me.tabAddContact.Controls.Add(Me.btnCAddContact)
        Me.tabAddContact.Location = New System.Drawing.Point(4, 22)
        Me.tabAddContact.Name = "tabAddContact"
        Me.tabAddContact.Size = New System.Drawing.Size(431, 290)
        Me.tabAddContact.TabIndex = 6
        Me.tabAddContact.Text = "Add Contact"
        Me.tabAddContact.UseVisualStyleBackColor = True
        '
        'btnStopAddContacts
        '
        Me.btnStopAddContacts.Location = New System.Drawing.Point(301, 264)
        Me.btnStopAddContacts.Name = "btnStopAddContacts"
        Me.btnStopAddContacts.Size = New System.Drawing.Size(75, 23)
        Me.btnStopAddContacts.TabIndex = 11
        Me.btnStopAddContacts.Text = "Stop"
        Me.btnStopAddContacts.UseVisualStyleBackColor = True
        '
        'lbxContactAlerts
        '
        Me.lbxContactAlerts.FormattingEnabled = True
        Me.lbxContactAlerts.Location = New System.Drawing.Point(35, 72)
        Me.lbxContactAlerts.Name = "lbxContactAlerts"
        Me.lbxContactAlerts.Size = New System.Drawing.Size(349, 95)
        Me.lbxContactAlerts.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(188, 196)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Contact:"
        '
        'rtbCContact
        '
        Me.rtbCContact.Location = New System.Drawing.Point(35, 212)
        Me.rtbCContact.Name = "rtbCContact"
        Me.rtbCContact.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical
        Me.rtbCContact.Size = New System.Drawing.Size(349, 47)
        Me.rtbCContact.TabIndex = 8
        Me.rtbCContact.Text = ""
        '
        'mtxtCAccount
        '
        Me.mtxtCAccount.BackColor = System.Drawing.Color.Yellow
        Me.mtxtCAccount.Location = New System.Drawing.Point(77, 31)
        Me.mtxtCAccount.Mask = "00000-99999"
        Me.mtxtCAccount.Name = "mtxtCAccount"
        Me.mtxtCAccount.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.mtxtCAccount.Size = New System.Drawing.Size(100, 20)
        Me.mtxtCAccount.TabIndex = 6
        Me.mtxtCAccount.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.mtxtCAccount.TextMaskFormat = System.Windows.Forms.MaskFormat.ExcludePromptAndLiterals
        '
        'lblCAccount
        '
        Me.lblCAccount.AutoSize = True
        Me.lblCAccount.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCAccount.Location = New System.Drawing.Point(14, 34)
        Me.lblCAccount.Name = "lblCAccount"
        Me.lblCAccount.Size = New System.Drawing.Size(63, 13)
        Me.lblCAccount.TabIndex = 7
        Me.lblCAccount.Text = "Account #: "
        '
        'btnCAddContact
        '
        Me.btnCAddContact.Location = New System.Drawing.Point(177, 265)
        Me.btnCAddContact.Name = "btnCAddContact"
        Me.btnCAddContact.Size = New System.Drawing.Size(75, 23)
        Me.btnCAddContact.TabIndex = 0
        Me.btnCAddContact.Text = "Add Contact"
        Me.btnCAddContact.UseVisualStyleBackColor = True
        '
        'tabDPA
        '
        Me.tabDPA.Controls.Add(Me.CheckBox1)
        Me.tabDPA.Controls.Add(Me.btnBudgetBill)
        Me.tabDPA.Controls.Add(Me.mtxtDPAAcc)
        Me.tabDPA.Controls.Add(Me.GroupBox1)
        Me.tabDPA.Controls.Add(Me.Label3)
        Me.tabDPA.Controls.Add(Me.txtDPAmonthly)
        Me.tabDPA.Controls.Add(Me.Label2)
        Me.tabDPA.Controls.Add(Me.txtDPAdown)
        Me.tabDPA.Controls.Add(Me.btnDPAprocess)
        Me.tabDPA.Controls.Add(Me.chkMinPayment)
        Me.tabDPA.Controls.Add(Me.Label1)
        Me.tabDPA.Location = New System.Drawing.Point(4, 22)
        Me.tabDPA.Name = "tabDPA"
        Me.tabDPA.Size = New System.Drawing.Size(431, 290)
        Me.tabDPA.TabIndex = 3
        Me.tabDPA.Text = "DPA"
        Me.tabDPA.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(6, 87)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(90, 17)
        Me.CheckBox1.TabIndex = 14
        Me.CheckBox1.Text = "Budget Billing"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'btnBudgetBill
        '
        Me.btnBudgetBill.Location = New System.Drawing.Point(273, 265)
        Me.btnBudgetBill.Name = "btnBudgetBill"
        Me.btnBudgetBill.Size = New System.Drawing.Size(75, 23)
        Me.btnBudgetBill.TabIndex = 13
        Me.btnBudgetBill.Text = "Budget Bill"
        Me.btnBudgetBill.UseVisualStyleBackColor = True
        Me.btnBudgetBill.Visible = False
        '
        'mtxtDPAAcc
        '
        Me.mtxtDPAAcc.BackColor = System.Drawing.Color.Yellow
        Me.mtxtDPAAcc.Location = New System.Drawing.Point(66, 15)
        Me.mtxtDPAAcc.Mask = "00000-99999"
        Me.mtxtDPAAcc.Name = "mtxtDPAAcc"
        Me.mtxtDPAAcc.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.mtxtDPAAcc.Size = New System.Drawing.Size(100, 20)
        Me.mtxtDPAAcc.TabIndex = 1
        Me.mtxtDPAAcc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.mtxtDPAAcc.TextMaskFormat = System.Windows.Forms.MaskFormat.ExcludePromptAndLiterals
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(9, 116)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(419, 148)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Account Info:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(192, 44)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(47, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Monthly:"
        '
        'txtDPAmonthly
        '
        Me.txtDPAmonthly.Location = New System.Drawing.Point(239, 41)
        Me.txtDPAmonthly.Name = "txtDPAmonthly"
        Me.txtDPAmonthly.Size = New System.Drawing.Size(100, 20)
        Me.txtDPAmonthly.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(24, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Down:"
        '
        'txtDPAdown
        '
        Me.txtDPAdown.Location = New System.Drawing.Point(66, 41)
        Me.txtDPAdown.Name = "txtDPAdown"
        Me.txtDPAdown.Size = New System.Drawing.Size(100, 20)
        Me.txtDPAdown.TabIndex = 2
        '
        'btnDPAprocess
        '
        Me.btnDPAprocess.Location = New System.Drawing.Point(177, 265)
        Me.btnDPAprocess.Name = "btnDPAprocess"
        Me.btnDPAprocess.Size = New System.Drawing.Size(75, 23)
        Me.btnDPAprocess.TabIndex = 10
        Me.btnDPAprocess.Text = "Process DPA"
        Me.btnDPAprocess.UseVisualStyleBackColor = True
        '
        'chkMinPayment
        '
        Me.chkMinPayment.AutoSize = True
        Me.chkMinPayment.Location = New System.Drawing.Point(6, 67)
        Me.chkMinPayment.Name = "chkMinPayment"
        Me.chkMinPayment.Size = New System.Drawing.Size(117, 17)
        Me.chkMinPayment.TabIndex = 4
        Me.chkMinPayment.Text = "Minimum Payment?"
        Me.chkMinPayment.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Account #: "
        '
        'tabTDrive
        '
        Me.tabTDrive.AllowDrop = True
        Me.tabTDrive.Controls.Add(Me.btnTDClear)
        Me.tabTDrive.Controls.Add(Me.btnTDCreateEmail)
        Me.tabTDrive.Controls.Add(Me.lvTDriveFiles)
        Me.tabTDrive.Location = New System.Drawing.Point(4, 22)
        Me.tabTDrive.Name = "tabTDrive"
        Me.tabTDrive.Size = New System.Drawing.Size(431, 290)
        Me.tabTDrive.TabIndex = 7
        Me.tabTDrive.Text = "T: Drive"
        Me.tabTDrive.UseVisualStyleBackColor = True
        '
        'btnTDClear
        '
        Me.btnTDClear.Location = New System.Drawing.Point(356, 128)
        Me.btnTDClear.Name = "btnTDClear"
        Me.btnTDClear.Size = New System.Drawing.Size(64, 22)
        Me.btnTDClear.TabIndex = 2
        Me.btnTDClear.Text = "Clear"
        Me.btnTDClear.UseVisualStyleBackColor = True
        '
        'btnTDCreateEmail
        '
        Me.btnTDCreateEmail.Location = New System.Drawing.Point(162, 264)
        Me.btnTDCreateEmail.Name = "btnTDCreateEmail"
        Me.btnTDCreateEmail.Size = New System.Drawing.Size(89, 23)
        Me.btnTDCreateEmail.TabIndex = 1
        Me.btnTDCreateEmail.Text = "Create Emails"
        Me.btnTDCreateEmail.UseVisualStyleBackColor = True
        '
        'lvTDriveFiles
        '
        Me.lvTDriveFiles.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colType, Me.colSendTo, Me.colAcc, Me.colName})
        Me.lvTDriveFiles.FullRowSelect = True
        Me.lvTDriveFiles.Location = New System.Drawing.Point(4, 14)
        Me.lvTDriveFiles.Name = "lvTDriveFiles"
        Me.lvTDriveFiles.Size = New System.Drawing.Size(424, 108)
        Me.lvTDriveFiles.TabIndex = 0
        Me.lvTDriveFiles.UseCompatibleStateImageBehavior = False
        Me.lvTDriveFiles.View = System.Windows.Forms.View.Details
        '
        'colType
        '
        Me.colType.Text = "Type"
        Me.colType.Width = 55
        '
        'colSendTo
        '
        Me.colSendTo.Text = "Send To"
        Me.colSendTo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.colSendTo.Width = 100
        '
        'colAcc
        '
        Me.colAcc.Text = "Account"
        Me.colAcc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.colAcc.Width = 85
        '
        'colName
        '
        Me.colName.Text = "Name"
        Me.colName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.colName.Width = 150
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(-1, 342)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(439, 23)
        Me.ProgressBar.TabIndex = 4
        Me.ProgressBar.Value = 88
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.BackColor = System.Drawing.Color.Transparent
        Me.lblStatus.ForeColor = System.Drawing.Color.Black
        Me.lblStatus.Location = New System.Drawing.Point(191, 346)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(38, 13)
        Me.lblStatus.TabIndex = 3
        Me.lblStatus.Text = "DONE"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(435, 24)
        Me.MenuStrip1.TabIndex = 5
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "&Exit"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OptionsToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.ToolsToolStripMenuItem.Text = "&Tools"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.AboutToolStripMenuItem.Text = "&About"
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "SAMuel"
        Me.NotifyIcon1.Visible = True
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'lblBranding
        '
        Me.lblBranding.AutoSize = True
        Me.lblBranding.BackColor = System.Drawing.Color.Transparent
        Me.lblBranding.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBranding.Location = New System.Drawing.Point(330, 356)
        Me.lblBranding.Name = "lblBranding"
        Me.lblBranding.Size = New System.Drawing.Size(108, 9)
        Me.lblBranding.TabIndex = 6
        Me.lblBranding.Text = "Created by: Zachary Karpinski"
        '
        'timerConvert
        '
        Me.timerConvert.Interval = 1000
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(435, 363)
        Me.Controls.Add(Me.lblBranding)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.ProgressBar)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmMain"
        Me.Text = "SAMuel"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.groupOLAudit.ResumeLayout(False)
        Me.groupOLAudit.PerformLayout()
        Me.tabWordToTiff.ResumeLayout(False)
        Me.tabWordToTiff.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.tabRFax.ResumeLayout(False)
        Me.tabRFax.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.grpRFServer.ResumeLayout(False)
        Me.grpRFServer.PerformLayout()
        Me.tabKofax.ResumeLayout(False)
        Me.tabKofax.PerformLayout()
        Me.gbKFSource.ResumeLayout(False)
        Me.gbKFSource.PerformLayout()
        Me.tabAddContact.ResumeLayout(False)
        Me.tabAddContact.PerformLayout()
        Me.tabDPA.ResumeLayout(False)
        Me.tabDPA.PerformLayout()
        Me.tabTDrive.ResumeLayout(False)
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnConvert As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents tabWordToTiff As System.Windows.Forms.TabPage
    Friend WithEvents lblSubject As System.Windows.Forms.Label
    Friend WithEvents lblFrom As System.Windows.Forms.Label
    Friend WithEvents lblAcc As System.Windows.Forms.Label
    Friend WithEvents txtAcc As System.Windows.Forms.TextBox
    Friend WithEvents txtFrom As System.Windows.Forms.TextBox
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents btnReject As System.Windows.Forms.Button
    Friend WithEvents tabDPA As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnDPAprocess As System.Windows.Forms.Button
    Friend WithEvents chkMinPayment As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDPAmonthly As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDPAdown As System.Windows.Forms.TextBox
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents mtxtDPAAcc As System.Windows.Forms.MaskedTextBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnBudgetBill As System.Windows.Forms.Button
    Friend WithEvents tabRFax As System.Windows.Forms.TabPage
    Friend WithEvents ToolsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnRFax As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents lblRFserver As System.Windows.Forms.Label
    Friend WithEvents txtRFsvr As System.Windows.Forms.TextBox
    Friend WithEvents lblRFuser As System.Windows.Forms.Label
    Friend WithEvents txtRFuser As System.Windows.Forms.TextBox
    Friend WithEvents chkRFNTauth As System.Windows.Forms.CheckBox
    Friend WithEvents lblRFpassword As System.Windows.Forms.Label
    Friend WithEvents txtRFpw As System.Windows.Forms.TextBox
    Friend WithEvents grpRFServer As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtRFRecName As System.Windows.Forms.TextBox
    Friend WithEvents lblRFRecName As System.Windows.Forms.Label
    Friend WithEvents txtRFRecFax As System.Windows.Forms.TextBox
    Friend WithEvents lblRFRecFax As System.Windows.Forms.Label
    Friend WithEvents chkRFCoverSheet As System.Windows.Forms.CheckBox
    Friend WithEvents chkRFSaveRec As System.Windows.Forms.CheckBox
    Friend WithEvents lblDragAndDropWord As System.Windows.Forms.Label
    Friend WithEvents tabKofax As System.Windows.Forms.TabPage
    Friend WithEvents btnKFImport As System.Windows.Forms.Button
    Friend WithEvents lblKFBatchName As System.Windows.Forms.Label
    Friend WithEvents txtKFBatchName As System.Windows.Forms.TextBox
    Friend WithEvents chkAuditMode As System.Windows.Forms.CheckBox
    Friend WithEvents lblOutlookMessage As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents lblKFBatchType As System.Windows.Forms.Label
    Friend WithEvents cbKFBatchType As System.Windows.Forms.ComboBox
    Friend WithEvents gbKFSource As System.Windows.Forms.GroupBox
    Friend WithEvents rbKFUSMail As System.Windows.Forms.RadioButton
    Friend WithEvents rbKFEmail As System.Windows.Forms.RadioButton
    Friend WithEvents lblKFComments As System.Windows.Forms.Label
    Friend WithEvents txtKFComments As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents tabAddContact As System.Windows.Forms.TabPage
    Friend WithEvents btnCAddContact As System.Windows.Forms.Button
    Friend WithEvents mtxtCAccount As System.Windows.Forms.MaskedTextBox
    Friend WithEvents lblCAccount As System.Windows.Forms.Label
    Friend WithEvents rtbCContact As System.Windows.Forms.RichTextBox
    Friend WithEvents rbConvertTiff As System.Windows.Forms.RadioButton
    Friend WithEvents rbConvertPDF As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents groupOLAudit As System.Windows.Forms.GroupBox
    Friend WithEvents chkValidOnly As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents rbConvertDOC As System.Windows.Forms.RadioButton
    Friend WithEvents rbConvertIMAGE As System.Windows.Forms.RadioButton
    Friend WithEvents lblBranding As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lbxContactAlerts As System.Windows.Forms.ListBox
    Friend WithEvents btnStopAddContacts As System.Windows.Forms.Button
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lstEmailAttachments As System.Windows.Forms.ListView
    Friend WithEvents hType As System.Windows.Forms.ColumnHeader
    Friend WithEvents hName As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents rtbEmailBody As System.Windows.Forms.RichTextBox
    Friend WithEvents tabTDrive As System.Windows.Forms.TabPage
    Friend WithEvents lvTDriveFiles As System.Windows.Forms.ListView
    Friend WithEvents colType As System.Windows.Forms.ColumnHeader
    Friend WithEvents colSendTo As System.Windows.Forms.ColumnHeader
    Friend WithEvents colAcc As System.Windows.Forms.ColumnHeader
    Friend WithEvents colName As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnTDCreateEmail As System.Windows.Forms.Button
    Friend WithEvents btnTDClear As System.Windows.Forms.Button
    Friend WithEvents timerConvert As System.Windows.Forms.Timer

End Class
