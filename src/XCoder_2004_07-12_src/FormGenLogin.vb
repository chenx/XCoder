

Public Class FormGenLogin
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
    Friend WithEvents btnMetadata As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBoxPassword As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxLogin As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxCaseSensitive As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents btnUseDefaultFileNames As System.Windows.Forms.Button
    Friend WithEvents LabelSetFileNames As System.Windows.Forms.Label
    Friend WithEvents btnOpenOutputFolder As System.Windows.Forms.Button
    Friend WithEvents LabelOutputPath As System.Windows.Forms.Label
    Friend WithEvents LabelAfterBuild As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LabelStep3 As System.Windows.Forms.Label
    Friend WithEvents LabelStep2 As System.Windows.Forms.Label
    Friend WithEvents LabelStep1 As System.Windows.Forms.Label
    Friend WithEvents TabPageOption1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption3 As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormGenLogin))
        Me.LabelStep3 = New System.Windows.Forms.Label()
        Me.LabelStep2 = New System.Windows.Forms.Label()
        Me.LabelStep1 = New System.Windows.Forms.Label()
        Me.TabPageOption1 = New System.Windows.Forms.TabPage()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBoxCaseSensitive = New System.Windows.Forms.CheckBox()
        Me.TextBoxLogin = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBoxPassword = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnMetadata = New System.Windows.Forms.Button()
        Me.TabPageOption2 = New System.Windows.Forms.TabPage()
        Me.btnUseDefaultFileNames = New System.Windows.Forms.Button()
        Me.LabelSetFileNames = New System.Windows.Forms.Label()
        Me.TabPageOption3 = New System.Windows.Forms.TabPage()
        Me.btnOpenOutputFolder = New System.Windows.Forms.Button()
        Me.LabelOutputPath = New System.Windows.Forms.Label()
        Me.LabelAfterBuild = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.PanelStep.SuspendLayout()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControlDisplay.SuspendLayout()
        Me.PanelInstr.SuspendLayout()
        Me.PanelDispTitle.SuspendLayout()
        Me.PanelDisplay.SuspendLayout()
        Me.PanelStepDetail.SuspendLayout()
        Me.PanelInstrTitle.SuspendLayout()
        Me.PanelStepTitle.SuspendLayout()
        Me.TabPageSource.SuspendLayout()
        Me.TabPageQuickView.SuspendLayout()
        Me.GroupBoxInstr.SuspendLayout()
        Me.TabPageFileList.SuspendLayout()
        Me.TabControlOptions.SuspendLayout()
        Me.GroupBoxStep.SuspendLayout()
        Me.TabPageOption1.SuspendLayout()
        Me.TabPageOption2.SuspendLayout()
        Me.TabPageOption3.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelStep
        '
        Me.PanelStep.Size = New System.Drawing.Size(250, 559)
        Me.PanelStep.Visible = True
        '
        'SplitterStep
        '
        Me.SplitterStep.Size = New System.Drawing.Size(3, 559)
        Me.SplitterStep.Visible = True
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
        Me.btnCloseDisp.Location = New System.Drawing.Point(512, 0)
        Me.btnCloseDisp.Visible = True
        '
        'LabelDispTitle
        '
        Me.LabelDispTitle.Size = New System.Drawing.Size(512, 25)
        Me.LabelDispTitle.Visible = True
        '
        'PanelInstr
        '
        Me.PanelInstr.Location = New System.Drawing.Point(253, 159)
        Me.PanelInstr.Size = New System.Drawing.Size(539, 400)
        Me.PanelInstr.Visible = True
        '
        'PanelDispTitle
        '
        Me.PanelDispTitle.Size = New System.Drawing.Size(537, 25)
        Me.PanelDispTitle.Visible = True
        '
        'PanelDisplay
        '
        Me.PanelDisplay.Location = New System.Drawing.Point(253, 0)
        Me.PanelDisplay.Size = New System.Drawing.Size(539, 156)
        Me.PanelDisplay.Visible = True
        '
        'btnCloseStep
        '
        Me.btnCloseStep.Visible = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Visible = True
        '
        'ImageListChild
        '
        Me.ImageListChild.ImageStream = CType(resources.GetObject("ImageListChild.ImageStream"), System.Windows.Forms.ImageListStreamer)
        '
        'PanelStepDetail
        '
        Me.PanelStepDetail.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelStep3, Me.LabelStep2, Me.LabelStep1})
        Me.PanelStepDetail.Size = New System.Drawing.Size(248, 532)
        Me.PanelStepDetail.Visible = True
        '
        'LabelInstrTitle
        '
        Me.LabelInstrTitle.BackColor = System.Drawing.Color.Blue
        Me.LabelInstrTitle.ForeColor = System.Drawing.Color.White
        Me.LabelInstrTitle.Size = New System.Drawing.Size(512, 25)
        Me.LabelInstrTitle.Visible = True
        '
        'btnCloseInstr
        '
        Me.btnCloseInstr.Location = New System.Drawing.Point(512, 0)
        Me.btnCloseInstr.Visible = True
        '
        'PanelInstrTitle
        '
        Me.PanelInstrTitle.Size = New System.Drawing.Size(537, 25)
        Me.PanelInstrTitle.Visible = True
        '
        'LabelStepTitle
        '
        Me.LabelStepTitle.Visible = True
        '
        'PanelStepTitle
        '
        Me.PanelStepTitle.Visible = True
        '
        'SplitterInstr
        '
        Me.SplitterInstr.Location = New System.Drawing.Point(253, 156)
        Me.SplitterInstr.Size = New System.Drawing.Size(539, 3)
        Me.SplitterInstr.Visible = True
        '
        'TabPageQuickView
        '
        Me.TabPageQuickView.Size = New System.Drawing.Size(529, 100)
        '
        'SplitterInstrDetail
        '
        Me.SplitterInstrDetail.Size = New System.Drawing.Size(537, 3)
        Me.SplitterInstrDetail.Visible = True
        '
        'LabelInstr
        '
        Me.LabelInstr.Size = New System.Drawing.Size(531, 51)
        Me.LabelInstr.Visible = True
        '
        'GroupBoxInstr
        '
        Me.GroupBoxInstr.Size = New System.Drawing.Size(537, 72)
        Me.GroupBoxInstr.Visible = True
        '
        'TreeView1
        '
        Me.TreeView1.Size = New System.Drawing.Size(532, 474)
        Me.TreeView1.Visible = True
        '
        'TabPageFileList
        '
        Me.TabPageFileList.Size = New System.Drawing.Size(532, 474)
        '
        'TabControlOptions
        '
        Me.TabControlOptions.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPageOption1, Me.TabPageOption2, Me.TabPageOption3})
        Me.TabControlOptions.ItemSize = New System.Drawing.Size(0, 21)
        Me.TabControlOptions.Size = New System.Drawing.Size(537, 373)
        Me.TabControlOptions.Visible = True
        '
        'GroupBoxStep
        '
        Me.GroupBoxStep.Visible = True
        '
        'LabelStepInstr
        '
        Me.LabelStepInstr.Visible = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(125, 457)
        Me.btnNext.Visible = True
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(25, 457)
        Me.btnBack.Visible = True
        '
        'LabelStep3
        '
        Me.LabelStep3.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep3.Location = New System.Drawing.Point(24, 264)
        Me.LabelStep3.Name = "LabelStep3"
        Me.LabelStep3.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep3.TabIndex = 31
        Me.LabelStep3.Text = "3. Generate Files"
        Me.LabelStep3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep2
        '
        Me.LabelStep2.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep2.Location = New System.Drawing.Point(24, 240)
        Me.LabelStep2.Name = "LabelStep2"
        Me.LabelStep2.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep2.TabIndex = 30
        Me.LabelStep2.Text = "2. Rename Files"
        Me.LabelStep2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep1
        '
        Me.LabelStep1.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep1.Location = New System.Drawing.Point(24, 216)
        Me.LabelStep1.Name = "LabelStep1"
        Me.LabelStep1.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep1.TabIndex = 29
        Me.LabelStep1.Text = "1. Member Reg Information"
        Me.LabelStep1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TabPageOption1
        '
        Me.TabPageOption1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label3, Me.CheckBox1, Me.CheckBoxCaseSensitive, Me.TextBoxLogin, Me.Label2, Me.TextBoxPassword, Me.Label1, Me.btnMetadata})
        Me.TabPageOption1.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption1.Name = "TabPageOption1"
        Me.TabPageOption1.Size = New System.Drawing.Size(529, 344)
        Me.TabPageOption1.TabIndex = 1
        Me.TabPageOption1.Text = "Detail Information"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(48, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(448, 40)
        Me.Label3.TabIndex = 47
        Me.Label3.Text = "After adding registration form metadata to the Source window about, press the but" & _
        "ton below:"
        '
        'CheckBox1
        '
        Me.CheckBox1.Location = New System.Drawing.Point(56, 208)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(368, 24)
        Me.CheckBox1.TabIndex = 46
        Me.CheckBox1.Text = "Registration needs administrator approval"
        '
        'CheckBoxCaseSensitive
        '
        Me.CheckBoxCaseSensitive.Location = New System.Drawing.Point(56, 176)
        Me.CheckBoxCaseSensitive.Name = "CheckBoxCaseSensitive"
        Me.CheckBoxCaseSensitive.Size = New System.Drawing.Size(352, 24)
        Me.CheckBoxCaseSensitive.TabIndex = 45
        Me.CheckBoxCaseSensitive.Text = "Login is case sensitive"
        '
        'TextBoxLogin
        '
        Me.TextBoxLogin.Location = New System.Drawing.Point(168, 112)
        Me.TextBoxLogin.Name = "TextBoxLogin"
        Me.TextBoxLogin.Size = New System.Drawing.Size(184, 22)
        Me.TextBoxLogin.TabIndex = 44
        Me.TextBoxLogin.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(48, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 23)
        Me.Label2.TabIndex = 43
        Me.Label2.Text = "Login Field:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBoxPassword
        '
        Me.TextBoxPassword.Location = New System.Drawing.Point(168, 136)
        Me.TextBoxPassword.Name = "TextBoxPassword"
        Me.TextBoxPassword.Size = New System.Drawing.Size(184, 22)
        Me.TextBoxPassword.TabIndex = 42
        Me.TextBoxPassword.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(48, 136)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 23)
        Me.Label1.TabIndex = 41
        Me.Label1.Text = "Password Field:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnMetadata
        '
        Me.btnMetadata.Location = New System.Drawing.Point(48, 72)
        Me.btnMetadata.Name = "btnMetadata"
        Me.btnMetadata.Size = New System.Drawing.Size(184, 23)
        Me.btnMetadata.TabIndex = 0
        Me.btnMetadata.Text = "Add Form Metadata"
        '
        'TabPageOption2
        '
        Me.TabPageOption2.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnUseDefaultFileNames, Me.LabelSetFileNames})
        Me.TabPageOption2.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption2.Name = "TabPageOption2"
        Me.TabPageOption2.Size = New System.Drawing.Size(532, 344)
        Me.TabPageOption2.TabIndex = 2
        Me.TabPageOption2.Text = "Rename files"
        '
        'btnUseDefaultFileNames
        '
        Me.btnUseDefaultFileNames.Location = New System.Drawing.Point(32, 104)
        Me.btnUseDefaultFileNames.Name = "btnUseDefaultFileNames"
        Me.btnUseDefaultFileNames.Size = New System.Drawing.Size(160, 23)
        Me.btnUseDefaultFileNames.TabIndex = 4
        Me.btnUseDefaultFileNames.Text = "Use default names"
        '
        'LabelSetFileNames
        '
        Me.LabelSetFileNames.Location = New System.Drawing.Point(24, 24)
        Me.LabelSetFileNames.Name = "LabelSetFileNames"
        Me.LabelSetFileNames.Size = New System.Drawing.Size(464, 64)
        Me.LabelSetFileNames.TabIndex = 3
        Me.LabelSetFileNames.Text = "All the files to be generated are shown in the TreeView list above. If you want t" & _
        "o change the name of a file or folder, double click on a node or right click on " & _
        "a node and choose ""Rename"" from the menu."
        '
        'TabPageOption3
        '
        Me.TabPageOption3.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnOpenOutputFolder, Me.LabelOutputPath, Me.LabelAfterBuild, Me.Label6})
        Me.TabPageOption3.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption3.Name = "TabPageOption3"
        Me.TabPageOption3.Size = New System.Drawing.Size(532, 344)
        Me.TabPageOption3.TabIndex = 3
        Me.TabPageOption3.Text = "Generate Files"
        '
        'btnOpenOutputFolder
        '
        Me.btnOpenOutputFolder.Location = New System.Drawing.Point(24, 192)
        Me.btnOpenOutputFolder.Name = "btnOpenOutputFolder"
        Me.btnOpenOutputFolder.Size = New System.Drawing.Size(176, 23)
        Me.btnOpenOutputFolder.TabIndex = 10
        Me.btnOpenOutputFolder.Text = "Open Output Folder..."
        '
        'LabelOutputPath
        '
        Me.LabelOutputPath.Location = New System.Drawing.Point(24, 104)
        Me.LabelOutputPath.Name = "LabelOutputPath"
        Me.LabelOutputPath.Size = New System.Drawing.Size(472, 80)
        Me.LabelOutputPath.TabIndex = 9
        Me.LabelOutputPath.Text = "..."
        '
        'LabelAfterBuild
        '
        Me.LabelAfterBuild.Location = New System.Drawing.Point(24, 56)
        Me.LabelAfterBuild.Name = "LabelAfterBuild"
        Me.LabelAfterBuild.Size = New System.Drawing.Size(480, 32)
        Me.LabelAfterBuild.TabIndex = 8
        Me.LabelAfterBuild.Text = "You can copy the output module to other places you want. The output path of the m" & _
        "odule is:"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(480, 32)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Use the ""Build"" menu or toolbar triangle button to generate files."
        '
        'FormGenLogin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 559)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDisplay, Me.SplitterInstr, Me.PanelInstr, Me.SplitterStep, Me.PanelStep})
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FormGenLogin"
        Me.Text = "Member Registration (Login)"
        Me.PanelStep.ResumeLayout(False)
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControlDisplay.ResumeLayout(False)
        Me.PanelInstr.ResumeLayout(False)
        Me.PanelDispTitle.ResumeLayout(False)
        Me.PanelDisplay.ResumeLayout(False)
        Me.PanelStepDetail.ResumeLayout(False)
        Me.PanelInstrTitle.ResumeLayout(False)
        Me.PanelStepTitle.ResumeLayout(False)
        Me.TabPageSource.ResumeLayout(False)
        Me.TabPageQuickView.ResumeLayout(False)
        Me.GroupBoxInstr.ResumeLayout(False)
        Me.TabPageFileList.ResumeLayout(False)
        Me.TabControlOptions.ResumeLayout(False)
        Me.GroupBoxStep.ResumeLayout(False)
        Me.TabPageOption1.ResumeLayout(False)
        Me.TabPageOption2.ResumeLayout(False)
        Me.TabPageOption3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private nodeRegFolder As New TreeNode("memRegFolder", 2, 2)
    Private nodeReg As New TreeNode("register.asp", 0, 0)
    Private nodeRegConfirm As New TreeNode("registerConfirm.asp", 0, 0)
    Private nodeMemRegEdit As New TreeNode("memRegEdit.asp", 0, 0)
    Private nodeMemRegConfirm As New TreeNode("memRegConfirm.asp", 0, 0)
    Private nodeMemChgPassword As New TreeNode("memChgPass.asp", 0, 0)
    Private nodeAdmRegList As New TreeNode("adminRegList.asp", 0, 0)
    Private nodeAdmRegEdit As New TreeNode("adminRegEdit.asp", 0, 0)
    Private nodeAdmRegConfirm As New TreeNode("adminRegConfirm.asp", 0, 0)
    Private nodeAdmChgPassword As New TreeNode("adminChgPass.asp", 0, 0)
    Private nodeIncFolder As New TreeNode("include", 2, 2)
    Private nodeIncEditTable As New TreeNode("editTable.asp", 0, 0)
    Private nodeIncConfirmTable As New TreeNode("confirTable.asp", 0, 0)
    Private nodeIncFuncs As New TreeNode("regFuncs.asp", 0, 0)
    Private nodeDBTableFileName As New TreeNode("memreg.sql", 0, 0)
    Private nodeXMLFileName As New TreeNode("memreg.xml", 0, 0)

    ' class to process metadata
    'Private clsMD As New clsMetadata()
    'Dim metadata As String


    Private Sub FormGenLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.moduleType = ModuleMain.xcModuleType.xcLogin
        Me.NUM_STEPS = 3
        Me.RichTextBox1.Text = clsMD.getMetadataExample()
        Me.initModuleContent()
    End Sub

    Protected Overrides Sub initModuleContent()
        tmpFileFolder = SetupManager.ROOT_PATH & "tmp\" & Me.projectName & "\" & Me.moduleName & "\"
        If Not IOManager.folderExists(tmpFileFolder) Then IOManager.CreateFolder(tmpFileFolder)
        'MsgBox(Me.moduleName & ", tmpfilefolder=" & tmpFileFolder & ", tmpImagefolder=" & tmpImageFolder)

        Me.moduleOutputPath = Me.projectOutputPath & Me.moduleName & "\"

        Me.LabelOutputPath.Text = Me.moduleOutputPath
        'MsgBox(Me.moduleOutputPath)
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
                ElseIf lcase_line.StartsWith("regfile") Then
                    Me.nodeReg.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("regconfirmfile") Then
                    Me.nodeRegConfirm.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("memregeditfile") Then
                    Me.nodeMemRegEdit.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("memregconfirmfile") Then
                    Me.nodeRegConfirm.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("memchgpassfile") Then
                    Me.nodeMemChgPassword.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminreglistfile") Then
                    Me.nodeAdmRegList.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminregeditfile") Then
                    Me.nodeAdmRegEdit.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminregconfirmfile") Then
                    Me.nodeAdmRegConfirm.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminchgpassfile") Then
                    Me.nodeAdmChgPassword.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("incfolder") Then
                    Me.nodeIncFolder.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("incedittablefile") Then
                    Me.nodeIncEditTable.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("incconfirmtablefile") Then
                    Me.nodeIncConfirmTable.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("incfuncsfile") Then
                    Me.nodeIncFuncs.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("dbtablefilename") Then
                    Me.nodeDBTableFileName.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("xmlfilename") Then
                    Me.nodeXMLFileName.Text = MyUtil.getStrAfterDelimiter(line, "=")
                End If
            Next
        End If

        constructTreeViewFileList()
        Me.processStep()

    End Sub


    Protected Overrides Sub initTreeNodesText()
        nodeRegFolder.Text = Me.moduleOutputPath
        nodeReg.Text = "register.asp"
        nodeRegConfirm.Text = "registerConfirm.asp"
        nodeMemRegEdit.Text = "memRegEdit.asp"
        nodeMemRegConfirm.Text = "memRegConfirm.asp"
        nodeMemChgPassword.Text = "memChgPass.asp"
        nodeAdmRegList.Text = "adminRegList.asp"
        nodeAdmRegEdit.Text = "adminRegEdit.asp"
        nodeAdmRegConfirm.Text = "adminRegConfirm.asp"
        nodeAdmChgPassword.Text = "adminChgPass.asp"
        nodeIncFolder.Text = "include"
        nodeIncEditTable.Text = "editTable.asp"
        nodeIncConfirmTable.Text = "confirmTable.asp"
        nodeIncFuncs.Text = "regFuncs.asp"
        nodeDBTableFileName.Text = Me.moduleName & ".sql"
        nodeXMLFileName.Text = Me.moduleName & ".xml"
    End Sub


    Protected Friend Overrides Sub saveModule()
        Dim str As String = ""
        str += "CreateDate=" & Now() & vbCrLf
        str += "ModuleName=" & Me.moduleName & vbCrLf
        str += "ModuleType=" & Me.moduleType & vbCrLf
        str += vbCrLf
        str += "StepNumber=" & Me.stepNumber & vbCrLf
        str += vbCrLf
        str += "RegFile=" & Me.nodeReg.Text & vbCrLf
        str += "RegConfimFile=" & Me.nodeRegConfirm.Text & vbCrLf
        str += "MemRegEditFile=" & Me.nodeMemRegEdit.Text & vbCrLf
        str += "MemRegConfirmFile=" & Me.nodeMemRegConfirm.Text & vbCrLf
        str += "MemChgPassFile=" & Me.nodeMemChgPassword.Text & vbCrLf
        str += "AdminRegListFile=" & Me.nodeAdmRegList.Text & vbCrLf
        str += "AdminRegEditFile=" & Me.nodeAdmRegEdit.Text & vbCrLf
        str += "AdminRegConfirmFile=" & Me.nodeAdmRegConfirm.Text & vbCrLf
        str += "AdminChgPassFile=" & Me.nodeAdmChgPassword.Text & vbCrLf
        str += "IncFolder=" & Me.nodeIncFolder.Text & vbCrLf
        str += "IncEditTableFile=" & Me.nodeIncEditTable.Text & vbCrLf
        str += "IncConfirmTableFile=" & Me.nodeIncConfirmTable.Text & vbCrLf
        str += "IncFuncsFile=" & Me.nodeIncFuncs.Text & vbCrLf
        str += "DBTableFileName=" & Me.nodeDBTableFileName.Text & vbCrLf
        str += "XMLFileName=" & Me.nodeXMLFileName.Text & vbCrLf
        str += vbCrLf
        ' interface status

        IOManager.SaveTextToFile(str, Me.moduleFullPath)

        Me.infoSaved = True
    End Sub


    Protected Overrides Sub constructTreeViewFileList()
        Dim TreeViewFileList As TreeView = Me.TreeView1
        Try
            nodeRegFolder.Nodes.Add(Me.nodeIncFolder)
            nodeIncFolder.Nodes.Add(Me.nodeIncConfirmTable)
            nodeIncFolder.Nodes.Add(Me.nodeIncEditTable)
            nodeIncFolder.Nodes.Add(Me.nodeIncFuncs)
            nodeRegFolder.Nodes.Add(Me.nodeReg)
            nodeRegFolder.Nodes.Add(Me.nodeRegConfirm)
            nodeRegFolder.Nodes.Add(Me.nodeMemRegEdit)
            nodeRegFolder.Nodes.Add(Me.nodeMemRegConfirm)
            nodeRegFolder.Nodes.Add(Me.nodeMemChgPassword)
            nodeRegFolder.Nodes.Add(Me.nodeAdmRegList)
            nodeRegFolder.Nodes.Add(Me.nodeAdmRegEdit)
            nodeRegFolder.Nodes.Add(Me.nodeAdmRegConfirm)
            nodeRegFolder.Nodes.Add(Me.nodeAdmChgPassword)
            nodeRegFolder.Nodes.Add(Me.nodeDBTableFileName)
            nodeRegFolder.Nodes.Add(Me.nodeXMLFileName)

            nodeRegFolder.ExpandAll()
            TreeViewFileList.Nodes.Clear()
            TreeViewFileList.Nodes.Add(nodeRegFolder)
        Catch ex As Exception
            MsgBox("Cannot create initial node:" & ex.ToString)
            End
        End Try

    End Sub


    Private Sub TreeView1_BeforeLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TreeView1.BeforeLabelEdit
        e.CancelEdit = (e.Node.Text = Me.moduleOutputPath)
    End Sub



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Interface functions
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Protected Overrides Sub processStep()
        'MsgBox(stepNumber)
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
           CType(MyUtil.getControlFromName(Me.TabControlOptions, "TabpageOption" & stepNumber), TabPage)
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


    Private Sub TabControlOptions_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControlOptions.SelectedIndexChanged
        Select Case Me.TabControlOptions.SelectedIndex
            Case 0
                Me.stepNumber = 1
                Me.TabControlDisplay.SelectedTab = Me.TabPageSource
                Me.LabelStepInstr.Text = ""
            Case 1
                Me.stepNumber = 2
                Me.TabControlDisplay.SelectedTab = Me.TabPageFileList
                Me.LabelStepInstr.Text = "You don't really need to change filenames for the module to work. But in case you prefer other names, you can do it."
            Case 2
                Me.stepNumber = 3
                Me.LabelStepInstr.Text = "The shortcut to build project is F5. Only opening modules in a project are built. The shortcut to build only the currently active module is F6."
        End Select
        Me.processStep()
        prevSelectedIndex = Me.TabControlOptions.SelectedIndex
    End Sub


    Private Sub btnMetadata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMetadata.Click
        Me.metadata = Me.RichTextBox1.Text
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' These are functions to get the real contents for each file
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private websiteFuncFilePath As String = "../asp-bin/aspFuncLib.asp"
    Private moduleFuncFilePath As String
    Private moduleEditTableFilePath As String
    Private moduleConfirmTableFilePath As String
    Private loginAdmSecFilePath As String = "../login/admin_security.asp"
    Private loginMemSecFilePath As String = "../login/mem_security.asp"
    Private rootIncHeaderFilePath As String = "../rootInclude/header.asp"
    Private rootIncFooterFilePath As String = "../rootInclude/footer.asp"

    ''''
    ' Generate Web Site Framework
    '
    Protected Friend Overrides Sub buildModule( _
    Optional ByVal frameworkParams As clsWebsiteFrameworkParameters = Nothing)
        ' Init metadata. 

    End Sub


End Class
