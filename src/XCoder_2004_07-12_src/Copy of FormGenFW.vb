

Public Class FormGenFW
    Inherits XinCoder.FormChildBase

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents LabelStep4 As System.Windows.Forms.Label
    Friend WithEvents LabelStep3 As System.Windows.Forms.Label
    Friend WithEvents LabelStep2 As System.Windows.Forms.Label
    Friend WithEvents LabelStep1 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialogUploadImage As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TabPageOption1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption5 As System.Windows.Forms.TabPage
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormGenFW))
        Me.LabelStep4 = New System.Windows.Forms.Label()
        Me.LabelStep3 = New System.Windows.Forms.Label()
        Me.LabelStep2 = New System.Windows.Forms.Label()
        Me.LabelStep1 = New System.Windows.Forms.Label()
        Me.OpenFileDialogUploadImage = New System.Windows.Forms.OpenFileDialog()
        Me.TabPageOption1 = New System.Windows.Forms.TabPage()
        Me.TabPageOption2 = New System.Windows.Forms.TabPage()
        Me.TabPageOption3 = New System.Windows.Forms.TabPage()
        Me.TabPageOption4 = New System.Windows.Forms.TabPage()
        Me.TabPageOption5 = New System.Windows.Forms.TabPage()
        Me.GroupBoxStep.SuspendLayout()
        Me.TabControlOptions.SuspendLayout()
        Me.TabPageFileList.SuspendLayout()
        Me.PanelInstr.SuspendLayout()
        Me.PanelStep.SuspendLayout()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControlDisplay.SuspendLayout()
        Me.PanelDispTitle.SuspendLayout()
        Me.GroupBoxInstr.SuspendLayout()
        Me.PanelDisplay.SuspendLayout()
        Me.PanelStepDetail.SuspendLayout()
        Me.PanelInstrTitle.SuspendLayout()
        Me.TabPageQuickView.SuspendLayout()
        Me.PanelStepTitle.SuspendLayout()
        Me.TabPageSource.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(125, 457)
        Me.btnNext.Visible = True
        '
        'LabelStepInstr
        '
        Me.LabelStepInstr.Visible = True
        '
        'GroupBoxStep
        '
        Me.GroupBoxStep.Visible = True
        '
        'TabControlOptions
        '
        Me.TabControlOptions.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPageOption1, Me.TabPageOption2, Me.TabPageOption3, Me.TabPageOption4, Me.TabPageOption5})
        Me.TabControlOptions.ItemSize = New System.Drawing.Size(115, 21)
        Me.TabControlOptions.Visible = True
        '
        'TreeView1
        '
        Me.TreeView1.Visible = True
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(25, 457)
        Me.btnBack.Visible = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Size = New System.Drawing.Size(529, 100)
        Me.RichTextBox1.Visible = True
        '
        'PanelInstr
        '
        Me.PanelInstr.Location = New System.Drawing.Point(253, 159)
        Me.PanelInstr.Visible = True
        '
        'SplitterStep
        '
        Me.SplitterStep.Size = New System.Drawing.Size(3, 559)
        Me.SplitterStep.Visible = True
        '
        'PanelStep
        '
        Me.PanelStep.Size = New System.Drawing.Size(250, 559)
        Me.PanelStep.Visible = True
        '
        'AxWebBrowser1
        '
        Me.AxWebBrowser1.Size = New System.Drawing.Size(529, 100)
        Me.AxWebBrowser1.Visible = True
        '
        'TabControlDisplay
        '
        Me.TabControlDisplay.Size = New System.Drawing.Size(537, 129)
        Me.TabControlDisplay.Visible = True
        '
        'btnCloseDisp
        '
        Me.btnCloseDisp.Visible = True
        '
        'LabelDispTitle
        '
        Me.LabelDispTitle.Visible = True
        '
        'PanelDispTitle
        '
        Me.PanelDispTitle.Visible = True
        '
        'GroupBoxInstr
        '
        Me.GroupBoxInstr.Visible = True
        '
        'PanelDisplay
        '
        Me.PanelDisplay.Size = New System.Drawing.Size(539, 156)
        Me.PanelDisplay.Visible = True
        '
        'btnCloseStep
        '
        Me.btnCloseStep.Visible = True
        '
        'LabelInstr
        '
        Me.LabelInstr.Visible = True
        '
        'ImageListChild
        '
        Me.ImageListChild.ImageStream = CType(resources.GetObject("ImageListChild.ImageStream"), System.Windows.Forms.ImageListStreamer)
        '
        'SplitterInstrDetail
        '
        Me.SplitterInstrDetail.Visible = True
        '
        'PanelStepDetail
        '
        Me.PanelStepDetail.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelStep4, Me.LabelStep3, Me.LabelStep2, Me.LabelStep1})
        Me.PanelStepDetail.Size = New System.Drawing.Size(248, 532)
        Me.PanelStepDetail.Visible = True
        '
        'LabelInstrTitle
        '
        Me.LabelInstrTitle.BackColor = System.Drawing.Color.Blue
        Me.LabelInstrTitle.ForeColor = System.Drawing.Color.White
        Me.LabelInstrTitle.Visible = True
        '
        'btnCloseInstr
        '
        Me.btnCloseInstr.Visible = True
        '
        'PanelInstrTitle
        '
        Me.PanelInstrTitle.Visible = True
        '
        'LabelStepTitle
        '
        Me.LabelStepTitle.Visible = True
        '
        'TabPageQuickView
        '
        Me.TabPageQuickView.Size = New System.Drawing.Size(529, 100)
        '
        'PanelStepTitle
        '
        Me.PanelStepTitle.Visible = True
        '
        'SplitterInstr
        '
        Me.SplitterInstr.Location = New System.Drawing.Point(253, 156)
        Me.SplitterInstr.Visible = True
        '
        'TabPageSource
        '
        Me.TabPageSource.Size = New System.Drawing.Size(529, 100)
        '
        'LabelStep4
        '
        Me.LabelStep4.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep4.Location = New System.Drawing.Point(24, 290)
        Me.LabelStep4.Name = "LabelStep4"
        Me.LabelStep4.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep4.TabIndex = 31
        Me.LabelStep4.Text = "4. Generate Files"
        Me.LabelStep4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep3
        '
        Me.LabelStep3.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep3.Location = New System.Drawing.Point(24, 266)
        Me.LabelStep3.Name = "LabelStep3"
        Me.LabelStep3.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep3.TabIndex = 30
        Me.LabelStep3.Text = "3. Rename Files"
        Me.LabelStep3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep2
        '
        Me.LabelStep2.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep2.Location = New System.Drawing.Point(24, 242)
        Me.LabelStep2.Name = "LabelStep2"
        Me.LabelStep2.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep2.TabIndex = 29
        Me.LabelStep2.Text = "2. Change File Content"
        Me.LabelStep2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep1
        '
        Me.LabelStep1.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep1.Location = New System.Drawing.Point(24, 218)
        Me.LabelStep1.Name = "LabelStep1"
        Me.LabelStep1.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep1.TabIndex = 28
        Me.LabelStep1.Text = "1. Create Database"
        Me.LabelStep1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OpenFileDialogUploadImage
        '
        Me.OpenFileDialogUploadImage.Filter = "JPEG files (*.jpg)|*.jpg|GIF files (*.gif)|*.gif"
        Me.OpenFileDialogUploadImage.Multiselect = True
        Me.OpenFileDialogUploadImage.Title = "Upload Images"
        '
        'TabPageOption1
        '
        Me.TabPageOption1.Location = New System.Drawing.Point(4, 46)
        Me.TabPageOption1.Name = "TabPageOption1"
        Me.TabPageOption1.Size = New System.Drawing.Size(529, 248)
        Me.TabPageOption1.TabIndex = 0
        Me.TabPageOption1.Text = "Create Database"
        '
        'TabPageOption2
        '
        Me.TabPageOption2.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption2.Name = "TabPageOption2"
        Me.TabPageOption2.Size = New System.Drawing.Size(529, 269)
        Me.TabPageOption2.TabIndex = 1
        Me.TabPageOption2.Text = "Change File Content"
        '
        'TabPageOption3
        '
        Me.TabPageOption3.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption3.Name = "TabPageOption3"
        Me.TabPageOption3.Size = New System.Drawing.Size(529, 269)
        Me.TabPageOption3.TabIndex = 2
        Me.TabPageOption3.Text = "Add Registration Form"
        '
        'TabPageOption4
        '
        Me.TabPageOption4.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption4.Name = "TabPageOption4"
        Me.TabPageOption4.Size = New System.Drawing.Size(529, 269)
        Me.TabPageOption4.TabIndex = 3
        Me.TabPageOption4.Text = "Rename Files"
        '
        'TabPageOption5
        '
        Me.TabPageOption5.Location = New System.Drawing.Point(4, 46)
        Me.TabPageOption5.Name = "TabPageOption5"
        Me.TabPageOption5.Size = New System.Drawing.Size(529, 248)
        Me.TabPageOption5.TabIndex = 4
        Me.TabPageOption5.Text = "Create Files"
        '
        'FormGenFW
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 559)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDisplay, Me.SplitterInstr, Me.PanelInstr, Me.SplitterStep, Me.PanelStep})
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FormGenFW"
        Me.GroupBoxStep.ResumeLayout(False)
        Me.TabControlOptions.ResumeLayout(False)
        Me.TabPageFileList.ResumeLayout(False)
        Me.PanelInstr.ResumeLayout(False)
        Me.PanelStep.ResumeLayout(False)
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControlDisplay.ResumeLayout(False)
        Me.PanelDispTitle.ResumeLayout(False)
        Me.GroupBoxInstr.ResumeLayout(False)
        Me.PanelDisplay.ResumeLayout(False)
        Me.PanelStepDetail.ResumeLayout(False)
        Me.PanelInstrTitle.ResumeLayout(False)
        Me.TabPageQuickView.ResumeLayout(False)
        Me.PanelStepTitle.ResumeLayout(False)
        Me.TabPageSource.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private NUM_STEPS As Integer = 4
    Private prevSelectedIndex As Integer = 0 ' of TabControlOptions

    ' Used by button "View" to tmporarily view the pages.
    Private tmpFileFolder As String
    Private tmpFileSubFolder As String
    Private tmpImageFolder As String

    Private curIEUrl As String
    Private selectedRadioButton As String

    Private Sub FormGenFW_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.moduleType = ModuleMain.xcModuleType.xcWebSiteFramework
        'Me.initModuleContent()
    End Sub

    Private Sub FormGenFW_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Dim width As Integer = Me.TabControlOptions.Width - 40
        Me.LabelCreateDB.Width = width
        Me.LabelODBCInfo.Width = width
        Me.LabelSetFileNames.Width = width
        Me.LabelAfterBuild.Width = width
        Me.LabelOutputPath.Width = width
        Me.GroupBox1.Width = width - 20
        Me.GroupBox2.Width = width - 20
    End Sub



    Protected Friend Overrides Sub saveModule()
        Dim str As String = ""
        str += "CreateDate=" & Now() & vbCrLf
        str += "ModuleName=" & Me.moduleName & vbCrLf
        str += "ModuleType=" & Me.moduleType & vbCrLf
        str += vbCrLf
        str += "StepNumber=" & Me.stepNumber & vbCrLf
        str += vbCrLf
        str += "ODBCName=" & Me.TextBoxODBCName.Text & vbCrLf
        str += "ODBCLogin=" & Me.TextBoxLogin.Text & vbCrLf
        str += "ODBCPassword=" & Me.TextBoxPassword.Text & vbCrLf
        str += vbCrLf
        str += "indexFileName=" & Me.indexFileNode.Text & vbCrLf
        str += "cssFileName=" & Me.cssFileNode.Text & vbCrLf
        str += "adminFolderName=" & Me.adminFolderNode.Text & vbCrLf
        str += "adminHomeFileName=" & Me.adminHomeFileNode.Text & vbCrLf
        str += "adminApproveFileName=" & Me.adminApproveFileNode.Text & vbCrLf
        str += "aspBinFolderName=" & Me.aspBinFolderNode.Text & vbCrLf
        str += "aspBinAdovbsFileName=" & Me.aspBinAdovbsFileNode.Text & vbCrLf
        str += "aspBinDBConfigFileName=" & Me.aspBinDBConfigFileNode.Text & vbCrLf
        str += "aspBinAspFuncsLibFileName=" & Me.aspBinAspFuncLibFileNode.Text & vbCrLf
        str += "loginFolderName=" & Me.loginFolderNode.Text & vbCrLf
        str += "loginLoginFileName=" & Me.loginLoginFileNode.Text & vbCrLf
        str += "loginLogoutFileName=" & Me.loginLogoutFileNode.Text & vbCrLf
        str += "loginAdminSecFileName=" & Me.loginAdminSecFileNode.Text & vbCrLf
        str += "loginMemSecFileName=" & Me.loginMemSecFileNode.Text & vbCrLf
        str += "memberFolderName=" & Me.memberFolderNode.Text & vbCrLf
        str += "memberHomeFileName=" & Me.memberHomeFileNode.Text & vbCrLf
        str += "rootIncFolderName=" & Me.rootIncFolderNode.Text & vbCrLf
        str += "rootIncHeaderFileName=" & Me.rootIncHeaderFileNode.Text & vbCrLf
        str += "rootIncFooterFileName=" & Me.rootIncFooterFileNode.Text & vbCrLf
        str += "imageFolderName=" & Me.imagesFolderNode.Text & vbCrLf
        str += "indexFileContent=" & Me.moduleName & ".index.txt" & vbCrLf
        str += "cssFileContent=" & Me.moduleName & ".css.txt" & vbCrLf
        str += "rootIncHeaderFileContent=" & Me.moduleName & ".header.txt" & vbCrLf
        str += "rootIncFooterFileContent=" & Me.moduleName & ".footer.txt" & vbCrLf
        str += vbCrLf
        ' interface status
        str += "curIEUrl=" & Me.curIEUrl & vbCrLf
        str += "curRadioButton=" & Me.selectedRadioButton & vbCrLf
        str += "btnChgContentStatus=" & Me.btnChgContent.Enabled & vbCrLf
        str += "curDispTab=" & Me.TabControlDisplay.SelectedTab.Name & vbCrLf

        IOManager.SaveTextToFile(str, Me.moduleFullPath)

        IOManager.SaveTextToFile(Me.getIndexFile, IOManager.getFolder(Me.moduleFullPath) & Me.moduleName & ".index.txt")
        IOManager.SaveTextToFile(Me.getCssFile, IOManager.getFolder(Me.moduleFullPath) & Me.moduleName & ".css.txt")
        IOManager.SaveTextToFile(Me.getRootIncHeaderFile, IOManager.getFolder(Me.moduleFullPath) & Me.moduleName & ".header.txt")
        IOManager.SaveTextToFile(Me.getRootIncFooterFile, IOManager.getFolder(Me.moduleFullPath) & Me.moduleName & ".footer.txt")

        Me.infoSaved = True
    End Sub


    Private Sub initTreeNodesText()
        ' Treeview nodes
        indexFileNode.Text = "index.asp"
        cssFileNode.Text = "xc.css"
        adminFolderNode.Text = "admin"
        adminHomeFileNode.Text = "index.asp"
        adminApproveFileNode.Text = "approveAdmin.asp"
        aspBinFolderNode.Text = "asp-bin"
        aspBinAdovbsFileNode.Text = "adovbs.inc"
        aspBinDBConfigFileNode.Text = "db_config.asp"
        aspBinAspFuncLibFileNode.Text = "aspFuncLib.asp"
        loginFolderNode.Text = "login"
        loginLoginFileNode.Text = "login.asp"
        loginLogoutFileNode.Text = "logout.asp"
        loginAdminSecFileNode.Text = "admin_Security.asp"
        loginMemSecFileNode.Text = "mem_security.asp"
        memberFolderNode.Text = "member"
        memberHomeFileNode.Text = "index.asp"
        rootIncFolderNode.Text = "rootInclude"
        rootIncHeaderFileNode.Text = "header.asp"
        rootIncFooterFileNode.Text = "footer.asp"
        imagesFolderNode.Text = "images"
    End Sub


    ' Use moduleInfoLines to initialized module.
    Protected Overrides Sub initModuleContent()
        tmpFileFolder = SetupManager.ROOT_PATH & "tmp\" & Me.projectName & "\" & Me.moduleName & "\"
        tmpFileSubFolder = SetupManager.ROOT_PATH & "tmp\" & Me.projectName & "\" & Me.moduleName & "\sub\"
        tmpImageFolder = SetupManager.ROOT_PATH & "tmp\" & Me.projectName & "\" & Me.moduleName & "\images\"
        IOManager.CreateFolder(tmpFileFolder)
        IOManager.CreateFolder(tmpFileSubFolder)
        IOManager.CreateFolder(tmpImageFolder)
        'MsgBox(Me.moduleName & ", tmpfilefolder=" & tmpFileFolder & ", tmpImagefolder=" & tmpImageFolder)

        Me.LabelOutputPath.Text = Me.outputPath
        Me.stepNumber = 1 ' later should be from file
        Me.initTreeNodesText() ' text may change from file

        Dim lines() As String
        If Me.readModuleInfo(lines) = True Then
            ' init variables
            Dim i As Integer
            Dim maxI As Integer = lines.Length - 1
            Dim line As String
            Dim lcase_line As String

            For i = 0 To maxI
                line = lines(i)
                lcase_line = LCase(line)
                If lcase_line.StartsWith("stepnumber") Then
                    Me.stepNumber = CType(MyUtil.getStrAfterDelimiter(line, "="), Integer)
                ElseIf lcase_line.StartsWith("odbcname") Then
                    Me.TextBoxODBCName.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("odbclogin") Then
                    Me.TextBoxLogin.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("odbcpassword") Then
                    Me.TextBoxPassword.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("indexfilename") Then
                    Me.indexFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("cssfilename") Then
                    Me.cssFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminfoldername") Then
                    Me.adminFolderNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminhomefilename") Then
                    Me.adminHomeFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminapprovefilename") Then
                    Me.adminApproveFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("aspbinfoldername") Then
                    Me.aspBinFolderNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("aspbinadovbsfilename") Then
                    Me.aspBinAdovbsFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("aspbindbconfigfilename") Then
                    Me.aspBinDBConfigFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("aspbinaspfuncslibfilename") Then
                    Me.aspBinAspFuncLibFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("loginfoldername") Then
                    Me.loginFolderNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("loginloginfilename") Then
                    Me.loginLoginFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("loginlogoutfilename") Then
                    Me.loginLogoutFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("loginadminsecfilename") Then
                    Me.loginAdminSecFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("loginmemsecfilename") Then
                    Me.loginMemSecFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("memberfoldername") Then
                    Me.memberFolderNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("memberhomefilename") Then
                    Me.memberHomeFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("rootincfoldername") Then
                    Me.rootIncFolderNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("rootincheaderfilename") Then
                    Me.rootIncHeaderFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("rootincfooterfilename") Then
                    Me.rootIncFooterFileNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("imagefoldername") Then
                    Me.imagesFolderNode.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("indexfilecontent") Then
                    Me.setIndexFile(IOManager.GetFileContents(IOManager.getFolder(Me.moduleFullPath) & MyUtil.getStrAfterDelimiter(line, "=")))
                ElseIf lcase_line.StartsWith("cssfilecontent") Then
                    Me.setCssFile(IOManager.GetFileContents(IOManager.getFolder(Me.moduleFullPath) & MyUtil.getStrAfterDelimiter(line, "=")))
                ElseIf lcase_line.StartsWith("rootincheaderfilecontent") Then
                    Me.setRootIncHeaderFile(IOManager.GetFileContents(IOManager.getFolder(Me.moduleFullPath) & MyUtil.getStrAfterDelimiter(line, "=")))
                ElseIf lcase_line.StartsWith("rootincfooterfilecontent") Then
                    Me.setRootIncFooterFile(IOManager.GetFileContents(IOManager.getFolder(Me.moduleFullPath) & MyUtil.getStrAfterDelimiter(line, "=")))
                ElseIf lcase_line.StartsWith("curradiobutton") Then
                    If MyUtil.getStrAfterDelimiter(line, "=") <> "" Then
                        CType(MyUtil.getControlFromName(Me.GroupBox1, MyUtil.getStrAfterDelimiter(line, "=")), RadioButton).Checked = True
                    End If
                ElseIf lcase_line.StartsWith("curieurl") Then
                    Me.curIEUrl = MyUtil.getStrAfterDelimiter(line, "=")
                    If Me.curIEUrl <> "" AndAlso IOManager.fileExists(Me.curIEUrl) Then Me.AxWebBrowser1.Navigate(Me.curIEUrl)
                ElseIf lcase_line.StartsWith("btnchgcontentstatus") Then
                    Me.btnChgContent.Enabled = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("curdisptab") Then
                    Me.TabControlDisplay.SelectedTab = CType(MyUtil.getControlFromName(Me.TabControlDisplay, MyUtil.getStrAfterDelimiter(line, "=")), TabPage)
                End If
            Next
        End If

        constructTreeViewFileList(Me.TreeView1)
        Me.processStep()

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Treeview nodes
    Private indexFileNode As New TreeNode("index.asp", 0, 0)
    Private cssFileNode As New TreeNode("xc.asp", 0, 0)
    Private adminFolderNode As New TreeNode("admin", 2, 2)
    Private adminHomeFileNode As New TreeNode("index.asp", 0, 0)
    Private adminApproveFileNode As New TreeNode("approveAdmin.asp", 0, 0)
    Private aspBinFolderNode As New TreeNode("asp-bin", 2, 2)
    Private aspBinAdovbsFileNode As New TreeNode("adovbs.inc", 0, 0)
    Private aspBinDBConfigFileNode As New TreeNode("db_config.asp", 0, 0)
    Private aspBinAspFuncLibFileNode As New TreeNode("aspFuncLib.asp", 0, 0)
    Private loginFolderNode As New TreeNode("login", 2, 2)
    Private loginLoginFileNode As New TreeNode("login.asp", 0, 0)
    Private loginLogoutFileNode As New TreeNode("logout.asp", 0, 0)
    Private loginAdminSecFileNode As New TreeNode("admin_Security.asp", 0, 0)
    Private loginMemSecFileNode As New TreeNode("mem_security.asp", 0, 0)
    Private memberFolderNode As New TreeNode("member", 2, 2)
    Private memberHomeFileNode As New TreeNode("index.asp", 0, 0)
    Private rootIncFolderNode As New TreeNode("rootInclude", 2, 2)
    Private rootIncHeaderFileNode As New TreeNode("header.asp", 0, 0)
    Private rootIncFooterFileNode As New TreeNode("footer.asp", 0, 0)
    Private imagesFolderNode As New TreeNode("images", 2, 2)


    Friend Sub constructTreeViewFileList(ByRef TreeViewFileList As TreeView)
        Try
            TreeViewFileList.Nodes.Clear()

            Me.rootIncFolderNode.Nodes.Add(Me.rootIncHeaderFileNode)
            Me.rootIncFolderNode.Nodes.Add(Me.rootIncFooterFileNode)

            Me.adminFolderNode.Nodes.Add(Me.adminHomeFileNode)
            Me.adminFolderNode.Nodes.Add(Me.adminApproveFileNode)

            Me.aspBinFolderNode.Nodes.Add(Me.aspBinAdovbsFileNode)
            Me.aspBinFolderNode.Nodes.Add(Me.aspBinDBConfigFileNode)
            Me.aspBinFolderNode.Nodes.Add(Me.aspBinAspFuncLibFileNode)

            Me.loginFolderNode.Nodes.Add(Me.loginLoginFileNode)
            Me.loginFolderNode.Nodes.Add(Me.loginLogoutFileNode)
            Me.loginFolderNode.Nodes.Add(Me.loginAdminSecFileNode)
            Me.loginFolderNode.Nodes.Add(Me.loginMemSecFileNode)

            Me.memberFolderNode.Nodes.Add(Me.memberHomeFileNode)

            Dim rootNode As New TreeNode(outputPath, 2, 2)

            rootNode.Nodes.Add(Me.indexFileNode)
            rootNode.Nodes.Add(Me.cssFileNode)
            rootNode.Nodes.Add(Me.adminFolderNode)
            rootNode.Nodes.Add(Me.aspBinFolderNode)
            rootNode.Nodes.Add(Me.loginFolderNode)
            rootNode.Nodes.Add(Me.memberFolderNode)
            rootNode.Nodes.Add(Me.rootIncFolderNode)
            rootNode.Nodes.Add(Me.imagesFolderNode)

            rootNode.ExpandAll()

            TreeViewFileList.Nodes.Add(rootNode)
        Catch ex As Exception
            MsgBox("Cannot create initial node:" & ex.ToString)
            End
        End Try

        'Me.TabPageFileList.Show()
        'Me.TabControlDisplay.SelectedTab = Me.TabPageFileList

    End Sub


    Private Sub TreeView1_BeforeLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TreeView1.BeforeLabelEdit
        e.CancelEdit = (e.Node.Text = Me.outputPath)
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' These are functions to get the real contents for each file
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''
    ' Generate Web Site Framework
    '
    Public Overrides Sub generateModule()
        ' Clear generated files of last time.
        If IOManager.folderExists(outputPath) Then
            IOManager.DeleteFolder(outputPath)
            IOManager.CreateFolder(outputPath)
        End If

        'MsgBox("validate database information")
        If Me.prevSelectedIndex = 0 Then
            If Trim(Me.TextBoxODBCName.Text) = "" Or _
               Trim(Me.TextBoxLogin.Text) = "" Or _
               Trim(Me.TextBoxPassword.Text) = "" Then
                MsgBox("It's fine to leave ODBC DSN information empty. But you may need to manually edit file asp-bin/database_config.asp later.")
            End If
        End If

        Dim progressBar As New FormProgressBar()
        progressBar.setText(Me.moduleName & SetupManager.MODULE_FILE_EXTENSION)
        progressBar.Show()
        'progressBar.Owner = Me
        progressBar.setProgress(1)

        If outputPath.EndsWith("\") = False Then
            outputPath &= "\"
        End If
        Dim adminFolder As String = Me.adminFolderNode.Text & "\"
        IOManager.CreateFolder(outputPath & adminFolder)
        IOManager.SaveTextToFile(Me.getAdminHomeFile, _
                  outputPath & adminFolder & Me.adminHomeFileNode.Text)
        IOManager.SaveTextToFile(Me.getAdminApproveFile, _
                  outputPath & adminFolder & Me.adminApproveFileNode.Text)

        progressBar.setProgress(2)

        Dim aspBinFolder As String = Me.aspBinFolderNode.Text & "\"
        IOManager.CreateFolder(outputPath & aspBinFolder)
        IOManager.SaveTextToFile(Me.getAspBinAdovbsFile, _
                  outputPath & aspBinFolder & Me.aspBinAdovbsFileNode.Text)
        IOManager.SaveTextToFile(Me.getAspBinAspFuncLibFile, _
                  outputPath & aspBinFolder & Me.aspBinAspFuncLibFileNode.Text)
        IOManager.SaveTextToFile(getAspBinDBConfigFile(), _
                  outputPath & aspBinFolder & Me.aspBinDBConfigFileNode.Text)

        progressBar.setProgress(1)

        Dim rootIncFolder As String = Me.rootIncFolderNode.Text & "\"
        IOManager.CreateFolder(outputPath & rootIncFolder)
        IOManager.SaveTextToFile(Me.getRootIncHeaderFile, _
                  outputPath & rootIncFolder & Me.rootIncHeaderFileNode.Text)
        IOManager.SaveTextToFile(Me.getRootIncFooterFile, _
                  outputPath & rootIncFolder & Me.rootIncFooterFileNode.Text)

        progressBar.setProgress(1)

        Dim memFolder As String = Me.memberFolderNode.Text & "\"
        IOManager.CreateFolder(outputPath & memFolder)
        IOManager.SaveTextToFile(Me.getMemHomeFile, _
                  outputPath & memFolder & Me.memberHomeFileNode.Text)

        progressBar.setProgress(1)

        Dim loginFolder As String = Me.loginFolderNode.Text & "\"
        IOManager.CreateFolder(outputPath & loginFolder)
        IOManager.SaveTextToFile(Me.getLoginLoginFile, _
                  outputPath & loginFolder & Me.loginLoginFileNode.Text)
        IOManager.SaveTextToFile(Me.getLoginLogoutFile, _
                  outputPath & loginFolder & Me.loginLogoutFileNode.Text)
        IOManager.SaveTextToFile(Me.getLoginAdminSecFile, _
                  outputPath & loginFolder & Me.loginAdminSecFileNode.Text)
        IOManager.SaveTextToFile(Me.getLoginMemSecFile, _
                  outputPath & loginFolder & Me.loginMemSecFileNode.Text)

        progressBar.setProgress(2)

        Dim imageFolder As String = Me.imagesFolderNode.Text & "\"
        IOManager.CreateFolder(outputPath & imageFolder)
        ' If there are uploaded image files, copy to target folder
        If IOManager.folderExists(tmpImageFolder) Then
            Dim images As String() = IO.Directory.GetFiles(tmpImageFolder)
            Dim image As String
            For Each image In images
                IO.File.Copy(image, outputPath & imageFolder & IOManager.getFile(image), True)
            Next
        End If


        progressBar.setProgress(1)

        IOManager.SaveTextToFile(Me.getCssFile, outputPath & Me.cssFileNode.Text)
        IOManager.SaveTextToFile(Me.getIndexFile, outputPath & Me.indexFileNode.Text)

        progressBar.setProgress(1)
        progressBar.Hide()
        progressBar.Close()
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Interface functions
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        stepnumber = stepnumber - 1
        Me.processStep()
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        stepnumber = stepnumber + 1
        Me.processStep()
    End Sub

    Private Sub processStep()
        ' Reset interface for this step
        If Me.stepNumber = 1 Then   ' first step
            Me.btnBack.Visible = False
            Me.btnNext.Visible = True
        ElseIf Me.stepNumber = NUM_STEPS Then    ' last step
            Me.btnBack.Visible = True
            Me.btnNext.Visible = False
        Else
            Me.btnBack.Visible = True
            Me.btnNext.Visible = True
        End If
        Me.TabControlOptions.SelectedTab = _
           CType(MyUtil.getControlFromName(Me.TabControlOptions, "TabpageOption" & stepnumber), TabPage)
        Me.setStepLabelsColor()

    End Sub

    Private Sub setStepLabelsColor()
        Dim i As Integer
        Dim curLabel As Label
        For i = 1 To NUM_STEPS
            curLabel = CType(MyUtil.getControlFromName(Me.PanelStepDetail, "LabelStep" & i), Label)
            If Me.stepNumber >= i Then ' done
                setStepLabelTrue(curLabel)
            Else    ' not done yet
                setStepLabelFalse(curLabel)
            End If
        Next
    End Sub

    Private Sub setStepLabelTrue(ByRef lbl As Label)
        lbl.BackColor = Color.Green
        lbl.ForeColor = Color.White
    End Sub

    Private Sub setStepLabelFalse(ByRef lbl As Label)
        lbl.BackColor = Color.LightGray
        lbl.ForeColor = Color.Black
    End Sub


    Private Sub ButtonCreateDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.curIEUrl = SetupManager.WEBSITE_LIB_PATH & "dbsetup.html"
        AxWebBrowser1.Navigate(Me.curIEUrl)
        Me.TabControlDisplay.SelectedTab = Me.TabPageQuickView
    End Sub


    Private Sub TabControlOptions_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControlOptions.SelectedIndexChanged
        Select Case Me.TabControlOptions.SelectedIndex
            Case 0
                Me.stepNumber = 1
                Me.LabelStepInstr.Text = "You should create the database for your website first. Follow the instruction in Instruction Explorer."
            Case 1
                Me.stepNumber = 2
                Me.TabControlDisplay.SelectedTab = Me.TabPageFileList
                Me.LabelStepInstr.Text = "You don't really need to change filenames for the module to work. But in case you prefer other names, you can do it."
            Case 2
                Me.stepNumber = 3
                Me.LabelStepInstr.Text = "The files NOT shown in Instruction Explorer are not necessary to change."
            Case 3
                Me.stepNumber = 4
                Me.LabelStepInstr.Text = "The shortcut to build project is F5. Only opening modules in a project are built. The shortcut to build only the currently active module is F6."
        End Select
        Me.processStep()
        prevSelectedIndex = Me.TabControlOptions.SelectedIndex
    End Sub

    Private Sub testODBC()
        Dim wwwroot As String = "c:\inetpub\wwwroot\"
        If IOManager.folderExists(wwwroot) Then
            Dim tmpFolder As String = wwwroot & "tmp\"
            If Not IOManager.folderExists(tmpFolder) Then
                IOManager.CreateFolder(tmpFolder)
                Dim str As String = ""
                str += ""

            End If
        End If
    End Sub




    ' event handler
    'Friend Sub ProcessExited(ByVal sender As Object, _
    '   ByVal e As System.EventArgs)
    '    Dim myProcess As Process = DirectCast( _
    '       sender, Process)
    '    MessageBox.Show("The process exited, raising " & _
    '       "the Exited event at: " & myProcess.ExitTime & _
    ''       "." & System.Environment.NewLine & _
    '       "Exit Code: " & myProcess.ExitCode)
    '    myProcess.Close()
    'End Sub

    Private Sub btnOpenOutputFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If IOManager.folderExists(Me.outputPath) Then
            Dim myProcess As New Process()
            myProcess = System.Diagnostics.Process.Start(Me.outputPath)
            'myProcess.StartInfo.FileName = Me.outputPath '& "a.txt"

            ' allow the process to raise events
            'myProcess.EnableRaisingEvents = True
            ' add an Exited event handler
            'AddHandler myProcess.Exited, AddressOf Me.ProcessExited
            ' start the process
            'myProcess.Start()
        Else
            MsgBox("Output folder is not created yet")
        End If
    End Sub

    Private Sub btnUseDefaultFileNames_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.initTreeNodesText()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''' content of files
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private indexFileContent As String
    Private cssFileContent As String
    Private rootIncHeaderFileContent As String
    Private rootIncFooterFileContent As String

    'Private indexFileNode As New TreeNode("index.asp", 0, 0)
    Private Function getIndexFile() As String
        'Dim str As String = ""
        Return Me.indexFileContent
    End Function

    Private Sub setIndexFile(ByVal content As String)
        Me.indexFileContent = content
    End Sub

    'Private cssFileNode As New TreeNode("xc.asp", 0, 0)
    Private Function getCssFile() As String
        'Dim str As String = ""
        Return Me.cssFileContent
    End Function

    Private Sub setCssFile(ByVal content As String)
        Me.cssFileContent = content
    End Sub

    'Private adminHomeFileNode As New TreeNode("index.asp", 0, 0)
    Private Function getAdminHomeFile() As String
        Dim str As String = ""
        Return str
    End Function

    'Private adminApproveFileNode As New TreeNode("approveAdmin.asp", 0, 0)
    Private Function getAdminApproveFile() As String
        Dim str As String = ""
        Return str
    End Function

    'Private aspBinAdovbsFileNode As New TreeNode("adovbs.inc", 0, 0)
    Private Function getAspBinAdovbsFile() As String
        Dim str As String = ""
        str = IOManager.GetFileContents(SetupManager.WEBSITE_LIB_PATH & "files/asp-bin/adovbs.inc")
        Return str
    End Function

    'Private aspBinDBConfigFileNode As New TreeNode("db_config.asp", 0, 0)
    Private Function getAspBinDBConfigFile() As String
        Dim str As String = ""
        str += "<%" & vbCrLf
        str += "''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str += "' SET LOGIN INFORMATION " & vbCrLf
        str += "''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str += "Dim db_db, db_login, db_password" & vbCrLf
        str += vbCrLf
        str += "db_db = " & Chr(34) & Trim(Me.TextBoxODBCName.Text) & Chr(34) & vbCrLf
        str += "db_login = " & Chr(34) & Trim(Me.TextBoxLogin.Text) & Chr(34) & vbCrLf
        str += "db_password = " & Chr(34) & Trim(Me.TextBoxPassword.Text) & Chr(34) & vbCrLf
        str += "%>"
        Return str
    End Function

    'Private aspBinAspFuncLibFileNode As New TreeNode("aspFuncLib.asp", 0, 0)
    Private Function getAspBinAspFuncLibFile() As String
        Dim str As String = ""
        str += "<!--#include file=" & Chr(34) & Me.aspBinAdovbsFileNode.Text & Chr(34) & "-->" & vbCrLf
        str += "<!--#include file=" & Chr(34) & Me.aspBinDBConfigFileNode.Text & Chr(34) & "-->" & vbCrLf
        str += vbCrLf
        str += IOManager.GetFileContents(SetupManager.WEBSITE_LIB_PATH & "files/asp-bin/xcDatabase.asp")
        Return str
    End Function

    'Private loginLoginFileNode As New TreeNode("login.asp", 0, 0)
    Private Function getLoginLoginFile() As String
        Dim str As String = ""
        Return str
    End Function

    'Private loginLogoutFileNode As New TreeNode("logout.asp", 0, 0)
    Private Function getLoginLogoutFile() As String
        Dim str As String = ""
        Return str
    End Function

    'Private loginAdminSecFileNode As New TreeNode("admin_Security.asp", 0, 0)
    Private Function getLoginAdminSecFile() As String
        Dim str As String = ""
        Return str
    End Function

    'Private loginMemSecFileNode As New TreeNode("mem_security.asp", 0, 0)
    Private Function getLoginMemSecFile() As String
        Dim str As String = ""
        Return str
    End Function

    'Private memberHomeFileNode As New TreeNode("index.asp", 0, 0)
    Private Function getMemHomeFile() As String
        Dim str As String = ""
        Return str
    End Function

    'Private rootIncHeaderFileNode As New TreeNode("header.asp", 0, 0)
    Private Function getRootIncHeaderFile() As String
        Return Me.rootIncHeaderFileContent
    End Function

    Private Sub setRootIncHeaderFile(ByVal content As String)
        Me.rootIncHeaderFileContent = content
    End Sub

    'Private rootIncFooterFileNode As New TreeNode("footer.asp", 0, 0)
    Private Function getRootIncFooterFile() As String
        Return Me.rootIncFooterFileContent
    End Function

    Private Sub setRootIncFooterFile(ByVal content As String)
        Me.rootIncFooterFileContent = content
    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions to change file contents
    ''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub RadioButtonIndex_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.RadioButtonIndex.Checked = True Then
            Me.RichTextBox1.Text = Me.getIndexFile()
            Me.TabControlDisplay.SelectedTab = Me.TabPageSource
            Me.btnChgContent.Enabled = True
            Me.btnView.Enabled = True
        End If
    End Sub

    Private Sub RadioButtonCss_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.RadioButtonCss.Checked = True Then
            Me.RichTextBox1.Text = Me.getCssFile()
            Me.TabControlDisplay.SelectedTab = Me.TabPageSource
            Me.btnChgContent.Enabled = True
            Me.btnView.Enabled = False
        End If
    End Sub

    Private Sub RadioButtonHeader_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.RadioButtonHeader.Checked = True Then
            Me.RichTextBox1.Text = Me.getRootIncHeaderFile()
            Me.TabControlDisplay.SelectedTab = Me.TabPageSource
            Me.btnChgContent.Enabled = True
            Me.btnView.Enabled = True
        End If
    End Sub

    Private Sub RadioButtonFooter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.RadioButtonFooter.Checked = True Then
            Me.RichTextBox1.Text = Me.getRootIncFooterFile()
            Me.TabControlDisplay.SelectedTab = Me.TabPageSource
            Me.btnChgContent.Enabled = True
            Me.btnView.Enabled = True
        End If
    End Sub

    Private Sub btnChgContent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.saveFileContent()
    End Sub


    Private Sub saveFileContent()
        If Me.RadioButtonIndex.Checked = True Then
            Me.setIndexFile(Me.RichTextBox1.Text)
            selectedRadioButton = Me.RadioButtonIndex.Name
        ElseIf Me.RadioButtonCss.Checked = True Then
            Me.setCssFile(Me.RichTextBox1.Text)
            selectedRadioButton = Me.RadioButtonCss.Name
        ElseIf Me.RadioButtonHeader.Checked = True Then
            Me.setRootIncHeaderFile(Me.RichTextBox1.Text)
            selectedRadioButton = Me.RadioButtonHeader.Name
        ElseIf Me.RadioButtonFooter.Checked = True Then
            Me.setRootIncFooterFile(Me.RichTextBox1.Text)
            selectedRadioButton = Me.RadioButtonFooter.Name
        End If
        Me.btnChgContent.Enabled = False
    End Sub


    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.saveFileContent()
        Dim tmpFilePath As String
        If Me.RadioButtonIndex.Checked = True Then
            tmpFilePath = tmpFileFolder & "index.html"
            IOManager.SaveTextToFile(Me.getIndexFile, tmpFilePath)
        ElseIf Me.RadioButtonHeader.Checked = True Or Me.RadioButtonFooter.Checked = True Then
            tmpFilePath = tmpFileSubFolder & "HeaderFooter.html"
            IOManager.SaveTextToFile(Me.getRootIncHeaderFile & vbCrLf & Me.getRootIncFooterFile, tmpFilePath)
        End If
        Me.curIEUrl = tmpFilePath
        Me.AxWebBrowser1.Navigate(Me.curIEUrl)
        Me.TabControlDisplay.SelectedTab = Me.TabPageQuickView
    End Sub



    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.OpenFileDialogUploadImage.ShowDialog(Me)
    End Sub

    Private Sub OpenFileDialogUploadImage_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialogUploadImage.FileOk
        'Dim tmpImageFolder As String = SetupManager.ROOT_PATH & "tmp\" & Me.moduleName & "\images\"
        IOManager.CreateFolder(tmpImageFolder)
        Dim imageFile As String
        For Each imageFile In Me.OpenFileDialogUploadImage.FileNames
            IO.File.Copy(imageFile, tmpImageFolder & IOManager.getFile(imageFile), True)
        Next
    End Sub

End Class
