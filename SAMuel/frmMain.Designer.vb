﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.btnRun = New System.Windows.Forms.Button()
        Me.picImage = New System.Windows.Forms.PictureBox()
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog()
        Me.btnConvert = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.rtbEmailBody = New System.Windows.Forms.RichTextBox()
        Me.btnReject = New System.Windows.Forms.Button()
        Me.txtAcc = New System.Windows.Forms.TextBox()
        Me.txtFrom = New System.Windows.Forms.TextBox()
        Me.txtSubject = New System.Windows.Forms.TextBox()
        Me.lblAcc = New System.Windows.Forms.Label()
        Me.lblFrom = New System.Windows.Forms.Label()
        Me.lblSubject = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
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
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.picImage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.tabDPA.SuspendLayout()
        Me.tabRFax.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.grpRFServer.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
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
        Me.btnConvert.Location = New System.Drawing.Point(177, 267)
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
        Me.TabControl1.Controls.Add(Me.tabDPA)
        Me.TabControl1.Controls.Add(Me.tabRFax)
        Me.TabControl1.Location = New System.Drawing.Point(-1, 27)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(439, 316)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnCancel)
        Me.TabPage1.Controls.Add(Me.btnNext)
        Me.TabPage1.Controls.Add(Me.rtbEmailBody)
        Me.TabPage1.Controls.Add(Me.btnReject)
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
        Me.TabPage1.Size = New System.Drawing.Size(431, 290)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Outlook"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(286, 267)
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
        Me.btnNext.Text = "Next"
        Me.btnNext.UseVisualStyleBackColor = True
        Me.btnNext.Visible = False
        '
        'rtbEmailBody
        '
        Me.rtbEmailBody.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbEmailBody.Location = New System.Drawing.Point(4, 100)
        Me.rtbEmailBody.Name = "rtbEmailBody"
        Me.rtbEmailBody.Size = New System.Drawing.Size(219, 154)
        Me.rtbEmailBody.TabIndex = 9
        Me.rtbEmailBody.Text = ""
        '
        'btnReject
        '
        Me.btnReject.Enabled = False
        Me.btnReject.Location = New System.Drawing.Point(286, 185)
        Me.btnReject.Name = "btnReject"
        Me.btnReject.Size = New System.Drawing.Size(75, 23)
        Me.btnReject.TabIndex = 8
        Me.btnReject.Text = "Reject"
        Me.btnReject.UseVisualStyleBackColor = True
        '
        'txtAcc
        '
        Me.txtAcc.Location = New System.Drawing.Point(59, 49)
        Me.txtAcc.Name = "txtAcc"
        Me.txtAcc.Size = New System.Drawing.Size(123, 20)
        Me.txtAcc.TabIndex = 7
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
        'txtSubject
        '
        Me.txtSubject.Enabled = False
        Me.txtSubject.Location = New System.Drawing.Point(59, 4)
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.ReadOnly = True
        Me.txtSubject.Size = New System.Drawing.Size(123, 20)
        Me.txtSubject.TabIndex = 5
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
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnConvert)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(431, 290)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Word -> tif"
        Me.TabPage2.UseVisualStyleBackColor = True
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
        Me.btnBudgetBill.Location = New System.Drawing.Point(273, 267)
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
        Me.btnDPAprocess.Location = New System.Drawing.Point(177, 267)
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
        Me.btnRFax.Location = New System.Drawing.Point(177, 267)
        Me.btnRFax.Name = "btnRFax"
        Me.btnRFax.Size = New System.Drawing.Size(75, 23)
        Me.btnRFax.TabIndex = 3
        Me.btnRFax.Text = "Fax"
        Me.btnRFax.UseVisualStyleBackColor = True
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(-1, 343)
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
        Me.lblStatus.Location = New System.Drawing.Point(203, 346)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(38, 13)
        Me.lblStatus.TabIndex = 3
        Me.lblStatus.Text = "DONE"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.AboutToolStripMenuItem})
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
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(435, 361)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.ProgressBar)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmMain"
        Me.Text = "SAMuel"
        CType(Me.picImage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.tabDPA.ResumeLayout(False)
        Me.tabDPA.PerformLayout()
        Me.tabRFax.ResumeLayout(False)
        Me.tabRFax.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.grpRFServer.ResumeLayout(False)
        Me.grpRFServer.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents picImage As System.Windows.Forms.PictureBox
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnConvert As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
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
    Friend WithEvents rtbEmailBody As System.Windows.Forms.RichTextBox
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents mtxtDPAAcc As System.Windows.Forms.MaskedTextBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
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

End Class
