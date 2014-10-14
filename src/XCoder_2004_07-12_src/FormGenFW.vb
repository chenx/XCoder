

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
    Friend WithEvents TabPageOption1 As System.Windows.Forms.TabPage
    Friend WithEvents TextBoxPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBoxLogin As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBoxODBCName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LabelODBCInfo As System.Windows.Forms.Label
    Friend WithEvents ButtonCreateDB As System.Windows.Forms.Button
    Friend WithEvents LabelCreateDB As System.Windows.Forms.Label
    Friend WithEvents LabelSetFileNames As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnUpload As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnChgContent As System.Windows.Forms.Button
    Friend WithEvents LabelOutputPath As System.Windows.Forms.Label
    Friend WithEvents LabelAfterBuild As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnOpenOutputFolder As System.Windows.Forms.Button
    Friend WithEvents btnUseDefaultFileNames As System.Windows.Forms.Button
    Friend WithEvents RadioButtonIndex As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonFooter As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonCss As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonHeader As System.Windows.Forms.RadioButton
    Friend WithEvents OpenFileDialogUploadImage As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TabPageOption2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption5 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption3 As System.Windows.Forms.TabPage
    Friend WithEvents LabelStep5 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxRegApprove As System.Windows.Forms.CheckBox
    Friend WithEvents btnRegForm As System.Windows.Forms.Button
    Friend WithEvents TextBoxDBTableName As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxChooseLogin As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxChoosePassword As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxChoosePassRpt As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TextBoxPasswdMinLen As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxRegLoginCaseSensitive As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxRegPassCaseSensitive As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormGenFW))
        Me.LabelStep4 = New System.Windows.Forms.Label()
        Me.LabelStep3 = New System.Windows.Forms.Label()
        Me.LabelStep2 = New System.Windows.Forms.Label()
        Me.LabelStep1 = New System.Windows.Forms.Label()
        Me.TabPageOption1 = New System.Windows.Forms.TabPage()
        Me.TextBoxPassword = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBoxLogin = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBoxODBCName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LabelODBCInfo = New System.Windows.Forms.Label()
        Me.ButtonCreateDB = New System.Windows.Forms.Button()
        Me.LabelCreateDB = New System.Windows.Forms.Label()
        Me.TabPageOption4 = New System.Windows.Forms.TabPage()
        Me.btnUseDefaultFileNames = New System.Windows.Forms.Button()
        Me.LabelSetFileNames = New System.Windows.Forms.Label()
        Me.TabPageOption2 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnUpload = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnChgContent = New System.Windows.Forms.Button()
        Me.RadioButtonIndex = New System.Windows.Forms.RadioButton()
        Me.RadioButtonFooter = New System.Windows.Forms.RadioButton()
        Me.RadioButtonCss = New System.Windows.Forms.RadioButton()
        Me.RadioButtonHeader = New System.Windows.Forms.RadioButton()
        Me.TabPageOption5 = New System.Windows.Forms.TabPage()
        Me.btnOpenOutputFolder = New System.Windows.Forms.Button()
        Me.LabelOutputPath = New System.Windows.Forms.Label()
        Me.LabelAfterBuild = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.OpenFileDialogUploadImage = New System.Windows.Forms.OpenFileDialog()
        Me.TabPageOption3 = New System.Windows.Forms.TabPage()
        Me.ComboBoxChoosePassRpt = New System.Windows.Forms.ComboBox()
        Me.ComboBoxChoosePassword = New System.Windows.Forms.ComboBox()
        Me.ComboBoxChooseLogin = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TextBoxDBTableName = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CheckBoxRegApprove = New System.Windows.Forms.CheckBox()
        Me.CheckBoxRegLoginCaseSensitive = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnRegForm = New System.Windows.Forms.Button()
        Me.LabelStep5 = New System.Windows.Forms.Label()
        Me.TextBoxPasswdMinLen = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.CheckBoxRegPassCaseSensitive = New System.Windows.Forms.CheckBox()
        Me.TabControlOptions.SuspendLayout()
        Me.GroupBoxInstr.SuspendLayout()
        Me.TabPageQuickView.SuspendLayout()
        Me.TabPageSource.SuspendLayout()
        Me.PanelStepTitle.SuspendLayout()
        Me.PanelInstrTitle.SuspendLayout()
        Me.PanelStepDetail.SuspendLayout()
        Me.PanelDisplay.SuspendLayout()
        Me.PanelDispTitle.SuspendLayout()
        Me.TabControlDisplay.SuspendLayout()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelStep.SuspendLayout()
        Me.PanelInstr.SuspendLayout()
        Me.GroupBoxStep.SuspendLayout()
        Me.TabPageFileList.SuspendLayout()
        Me.TabPageOption1.SuspendLayout()
        Me.TabPageOption4.SuspendLayout()
        Me.TabPageOption2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPageOption5.SuspendLayout()
        Me.TabPageOption3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControlOptions
        '
        Me.TabControlOptions.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPageOption1, Me.TabPageOption2, Me.TabPageOption3, Me.TabPageOption4, Me.TabPageOption5})
        Me.TabControlOptions.ItemSize = New System.Drawing.Size(115, 21)
        Me.TabControlOptions.Visible = True
        '
        'GroupBoxInstr
        '
        Me.GroupBoxInstr.Visible = True
        '
        'LabelInstr
        '
        Me.LabelInstr.Visible = True
        '
        'SplitterInstrDetail
        '
        Me.SplitterInstrDetail.Visible = True
        '
        'TabPageQuickView
        '
        Me.TabPageQuickView.Size = New System.Drawing.Size(529, 100)
        '
        'TabPageSource
        '
        Me.TabPageSource.Size = New System.Drawing.Size(1017, 443)
        '
        'SplitterInstr
        '
        Me.SplitterInstr.Location = New System.Drawing.Point(253, 156)
        Me.SplitterInstr.Visible = True
        '
        'PanelStepTitle
        '
        Me.PanelStepTitle.Visible = True
        '
        'LabelStepTitle
        '
        Me.LabelStepTitle.Visible = True
        '
        'PanelInstrTitle
        '
        Me.PanelInstrTitle.Visible = True
        '
        'btnCloseInstr
        '
        Me.btnCloseInstr.Visible = True
        '
        'LabelInstrTitle
        '
        Me.LabelInstrTitle.BackColor = System.Drawing.Color.Blue
        Me.LabelInstrTitle.ForeColor = System.Drawing.Color.White
        Me.LabelInstrTitle.Visible = True
        '
        'PanelStepDetail
        '
        Me.PanelStepDetail.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelStep5, Me.LabelStep4, Me.LabelStep3, Me.LabelStep2, Me.LabelStep1})
        Me.PanelStepDetail.Size = New System.Drawing.Size(248, 532)
        Me.PanelStepDetail.Visible = True
        '
        'ImageListChild
        '
        Me.ImageListChild.ImageStream = CType(resources.GetObject("ImageListChild.ImageStream"), System.Windows.Forms.ImageListStreamer)
        '
        'btnCloseStep
        '
        Me.btnCloseStep.Visible = True
        '
        'PanelDisplay
        '
        Me.PanelDisplay.Size = New System.Drawing.Size(539, 156)
        Me.PanelDisplay.Visible = True
        '
        'PanelDispTitle
        '
        Me.PanelDispTitle.Visible = True
        '
        'LabelDispTitle
        '
        Me.LabelDispTitle.Visible = True
        '
        'btnCloseDisp
        '
        Me.btnCloseDisp.Visible = True
        '
        'TabControlDisplay
        '
        Me.TabControlDisplay.Size = New System.Drawing.Size(537, 129)
        Me.TabControlDisplay.Visible = True
        '
        'AxWebBrowser1
        '
        Me.AxWebBrowser1.Size = New System.Drawing.Size(529, 100)
        Me.AxWebBrowser1.Visible = True
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(25, 457)
        Me.btnBack.Visible = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(125, 457)
        Me.btnNext.Visible = True
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
        'PanelInstr
        '
        Me.PanelInstr.Location = New System.Drawing.Point(253, 159)
        Me.PanelInstr.Size = New System.Drawing.Size(539, 400)
        Me.PanelInstr.Visible = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Size = New System.Drawing.Size(1017, 443)
        Me.RichTextBox1.Visible = True
        '
        'LabelStepInstr
        '
        Me.LabelStepInstr.Visible = True
        '
        'GroupBoxStep
        '
        Me.GroupBoxStep.Visible = True
        '
        'TabPageFileList
        '
        Me.TabPageFileList.Size = New System.Drawing.Size(1017, 443)
        '
        'TreeView1
        '
        Me.TreeView1.FullRowSelect = False
        Me.TreeView1.Size = New System.Drawing.Size(1017, 443)
        Me.TreeView1.Visible = True
        '
        'LabelStep4
        '
        Me.LabelStep4.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep4.Location = New System.Drawing.Point(24, 290)
        Me.LabelStep4.Name = "LabelStep4"
        Me.LabelStep4.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep4.TabIndex = 27
        Me.LabelStep4.Text = "4. Rename Files"
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
        Me.LabelStep3.TabIndex = 26
        Me.LabelStep3.Text = "3. Add Member Register Form"
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
        Me.LabelStep2.TabIndex = 25
        Me.LabelStep2.Text = "2. Add User Interface"
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
        Me.LabelStep1.TabIndex = 24
        Me.LabelStep1.Text = "1. Create Database"
        Me.LabelStep1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TabPageOption1
        '
        Me.TabPageOption1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBoxPassword, Me.Label5, Me.TextBoxLogin, Me.Label4, Me.TextBoxODBCName, Me.Label3, Me.LabelODBCInfo, Me.ButtonCreateDB, Me.LabelCreateDB})
        Me.TabPageOption1.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption1.Name = "TabPageOption1"
        Me.TabPageOption1.Size = New System.Drawing.Size(529, 269)
        Me.TabPageOption1.TabIndex = 0
        Me.TabPageOption1.Text = "Create Database"
        '
        'TextBoxPassword
        '
        Me.TextBoxPassword.Location = New System.Drawing.Point(124, 200)
        Me.TextBoxPassword.Name = "TextBoxPassword"
        Me.TextBoxPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.TextBoxPassword.Size = New System.Drawing.Size(184, 22)
        Me.TextBoxPassword.TabIndex = 42
        Me.TextBoxPassword.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(20, 200)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 23)
        Me.Label5.TabIndex = 41
        Me.Label5.Text = "Password:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBoxLogin
        '
        Me.TextBoxLogin.Location = New System.Drawing.Point(124, 176)
        Me.TextBoxLogin.Name = "TextBoxLogin"
        Me.TextBoxLogin.Size = New System.Drawing.Size(184, 22)
        Me.TextBoxLogin.TabIndex = 40
        Me.TextBoxLogin.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(20, 176)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 23)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "Login:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBoxODBCName
        '
        Me.TextBoxODBCName.Location = New System.Drawing.Point(124, 152)
        Me.TextBoxODBCName.Name = "TextBoxODBCName"
        Me.TextBoxODBCName.Size = New System.Drawing.Size(184, 22)
        Me.TextBoxODBCName.TabIndex = 38
        Me.TextBoxODBCName.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(20, 152)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(136, 23)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "ODBC DSN:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelODBCInfo
        '
        Me.LabelODBCInfo.Location = New System.Drawing.Point(12, 112)
        Me.LabelODBCInfo.Name = "LabelODBCInfo"
        Me.LabelODBCInfo.Size = New System.Drawing.Size(436, 25)
        Me.LabelODBCInfo.TabIndex = 36
        Me.LabelODBCInfo.Text = "Enter the ODBC system DSN name, login and password below. "
        '
        'ButtonCreateDB
        '
        Me.ButtonCreateDB.Location = New System.Drawing.Point(12, 64)
        Me.ButtonCreateDB.Name = "ButtonCreateDB"
        Me.ButtonCreateDB.Size = New System.Drawing.Size(332, 24)
        Me.ButtonCreateDB.TabIndex = 34
        Me.ButtonCreateDB.Text = "Guidance on create Database and ODBC DSN"
        '
        'LabelCreateDB
        '
        Me.LabelCreateDB.Location = New System.Drawing.Point(12, 16)
        Me.LabelCreateDB.Name = "LabelCreateDB"
        Me.LabelCreateDB.Size = New System.Drawing.Size(480, 32)
        Me.LabelCreateDB.TabIndex = 33
        Me.LabelCreateDB.Text = "If you have not yet created the database and ODBC DSN name, please do it yourself" & _
        " or follow the guidance."
        '
        'TabPageOption4
        '
        Me.TabPageOption4.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnUseDefaultFileNames, Me.LabelSetFileNames})
        Me.TabPageOption4.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption4.Name = "TabPageOption4"
        Me.TabPageOption4.Size = New System.Drawing.Size(1020, 269)
        Me.TabPageOption4.TabIndex = 1
        Me.TabPageOption4.Text = "Rename Files"
        '
        'btnUseDefaultFileNames
        '
        Me.btnUseDefaultFileNames.Location = New System.Drawing.Point(24, 104)
        Me.btnUseDefaultFileNames.Name = "btnUseDefaultFileNames"
        Me.btnUseDefaultFileNames.Size = New System.Drawing.Size(160, 23)
        Me.btnUseDefaultFileNames.TabIndex = 2
        Me.btnUseDefaultFileNames.Text = "Use default names"
        '
        'LabelSetFileNames
        '
        Me.LabelSetFileNames.Location = New System.Drawing.Point(16, 24)
        Me.LabelSetFileNames.Name = "LabelSetFileNames"
        Me.LabelSetFileNames.Size = New System.Drawing.Size(464, 64)
        Me.LabelSetFileNames.TabIndex = 1
        Me.LabelSetFileNames.Text = "All the files to be generated are shown in the TreeView list above. If you want t" & _
        "o change the name of a file or folder, double click on a node or right click on " & _
        "a node and choose ""Rename"" from the menu."
        '
        'TabPageOption2
        '
        Me.TabPageOption2.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox2, Me.GroupBox1})
        Me.TabPageOption2.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption2.Name = "TabPageOption2"
        Me.TabPageOption2.Size = New System.Drawing.Size(529, 269)
        Me.TabPageOption2.TabIndex = 2
        Me.TabPageOption2.Text = "User Interface"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnUpload})
        Me.GroupBox2.Location = New System.Drawing.Point(16, 168)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(472, 80)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Upload images to image folder"
        '
        'btnUpload
        '
        Me.btnUpload.Location = New System.Drawing.Point(24, 32)
        Me.btnUpload.Name = "btnUpload"
        Me.btnUpload.Size = New System.Drawing.Size(144, 24)
        Me.btnUpload.TabIndex = 0
        Me.btnUpload.Text = "Upload images..."
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnView, Me.btnChgContent, Me.RadioButtonIndex, Me.RadioButtonFooter, Me.RadioButtonCss, Me.RadioButtonHeader})
        Me.GroupBox1.Location = New System.Drawing.Point(16, 14)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(472, 138)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Please choose the file you want to change content:"
        '
        'btnView
        '
        Me.btnView.Enabled = False
        Me.btnView.Location = New System.Drawing.Point(184, 96)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(256, 23)
        Me.btnView.TabIndex = 7
        Me.btnView.Text = "Preview as HTML"
        '
        'btnChgContent
        '
        Me.btnChgContent.Enabled = False
        Me.btnChgContent.Location = New System.Drawing.Point(24, 96)
        Me.btnChgContent.Name = "btnChgContent"
        Me.btnChgContent.Size = New System.Drawing.Size(144, 23)
        Me.btnChgContent.TabIndex = 6
        Me.btnChgContent.Text = "Change file content"
        '
        'RadioButtonIndex
        '
        Me.RadioButtonIndex.Location = New System.Drawing.Point(24, 32)
        Me.RadioButtonIndex.Name = "RadioButtonIndex"
        Me.RadioButtonIndex.TabIndex = 0
        Me.RadioButtonIndex.Text = "index.asp"
        '
        'RadioButtonFooter
        '
        Me.RadioButtonFooter.Location = New System.Drawing.Point(216, 56)
        Me.RadioButtonFooter.Name = "RadioButtonFooter"
        Me.RadioButtonFooter.Size = New System.Drawing.Size(224, 24)
        Me.RadioButtonFooter.TabIndex = 4
        Me.RadioButtonFooter.Text = "rootInclude/footer.asp"
        '
        'RadioButtonCss
        '
        Me.RadioButtonCss.Location = New System.Drawing.Point(24, 56)
        Me.RadioButtonCss.Name = "RadioButtonCss"
        Me.RadioButtonCss.TabIndex = 3
        Me.RadioButtonCss.Text = "xc.css"
        '
        'RadioButtonHeader
        '
        Me.RadioButtonHeader.Location = New System.Drawing.Point(216, 32)
        Me.RadioButtonHeader.Name = "RadioButtonHeader"
        Me.RadioButtonHeader.Size = New System.Drawing.Size(224, 24)
        Me.RadioButtonHeader.TabIndex = 5
        Me.RadioButtonHeader.Text = "rootInclude/header.asp"
        '
        'TabPageOption5
        '
        Me.TabPageOption5.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnOpenOutputFolder, Me.LabelOutputPath, Me.LabelAfterBuild, Me.Label6})
        Me.TabPageOption5.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption5.Name = "TabPageOption5"
        Me.TabPageOption5.Size = New System.Drawing.Size(1020, 269)
        Me.TabPageOption5.TabIndex = 3
        Me.TabPageOption5.Text = "Build"
        '
        'btnOpenOutputFolder
        '
        Me.btnOpenOutputFolder.Location = New System.Drawing.Point(24, 192)
        Me.btnOpenOutputFolder.Name = "btnOpenOutputFolder"
        Me.btnOpenOutputFolder.Size = New System.Drawing.Size(176, 23)
        Me.btnOpenOutputFolder.TabIndex = 6
        Me.btnOpenOutputFolder.Text = "Open Output Folder..."
        '
        'LabelOutputPath
        '
        Me.LabelOutputPath.Location = New System.Drawing.Point(24, 104)
        Me.LabelOutputPath.Name = "LabelOutputPath"
        Me.LabelOutputPath.Size = New System.Drawing.Size(472, 80)
        Me.LabelOutputPath.TabIndex = 5
        Me.LabelOutputPath.Text = "..."
        '
        'LabelAfterBuild
        '
        Me.LabelAfterBuild.Location = New System.Drawing.Point(24, 56)
        Me.LabelAfterBuild.Name = "LabelAfterBuild"
        Me.LabelAfterBuild.Size = New System.Drawing.Size(480, 32)
        Me.LabelAfterBuild.TabIndex = 4
        Me.LabelAfterBuild.Text = "You can copy the output module to other places you want. The output path of the m" & _
        "odule is:"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(480, 32)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "Use the ""Build"" menu or toolbar triangle button to generate files."
        '
        'OpenFileDialogUploadImage
        '
        Me.OpenFileDialogUploadImage.Filter = "JPEG files (*.jpg)|*.jpg|GIF files (*.gif)|*.gif"
        Me.OpenFileDialogUploadImage.Multiselect = True
        Me.OpenFileDialogUploadImage.Title = "Upload Images"
        '
        'TabPageOption3
        '
        Me.TabPageOption3.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBoxRegPassCaseSensitive, Me.Label10, Me.TextBoxPasswdMinLen, Me.ComboBoxChoosePassRpt, Me.ComboBoxChoosePassword, Me.ComboBoxChooseLogin, Me.Label9, Me.TextBoxDBTableName, Me.Label8, Me.Label1, Me.CheckBoxRegApprove, Me.CheckBoxRegLoginCaseSensitive, Me.Label2, Me.Label7, Me.btnRegForm})
        Me.TabPageOption3.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption3.Name = "TabPageOption3"
        Me.TabPageOption3.Size = New System.Drawing.Size(529, 269)
        Me.TabPageOption3.TabIndex = 4
        Me.TabPageOption3.Text = "Register Form"
        '
        'ComboBoxChoosePassRpt
        '
        Me.ComboBoxChoosePassRpt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxChoosePassRpt.Location = New System.Drawing.Point(224, 128)
        Me.ComboBoxChoosePassRpt.Name = "ComboBoxChoosePassRpt"
        Me.ComboBoxChoosePassRpt.Size = New System.Drawing.Size(136, 24)
        Me.ComboBoxChoosePassRpt.TabIndex = 62
        '
        'ComboBoxChoosePassword
        '
        Me.ComboBoxChoosePassword.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxChoosePassword.Location = New System.Drawing.Point(224, 104)
        Me.ComboBoxChoosePassword.Name = "ComboBoxChoosePassword"
        Me.ComboBoxChoosePassword.Size = New System.Drawing.Size(136, 24)
        Me.ComboBoxChoosePassword.TabIndex = 61
        '
        'ComboBoxChooseLogin
        '
        Me.ComboBoxChooseLogin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxChooseLogin.Location = New System.Drawing.Point(224, 80)
        Me.ComboBoxChooseLogin.Name = "ComboBoxChooseLogin"
        Me.ComboBoxChooseLogin.Size = New System.Drawing.Size(136, 24)
        Me.ComboBoxChooseLogin.TabIndex = 60
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(16, 128)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(200, 23)
        Me.Label9.TabIndex = 58
        Me.Label9.Text = "Password Repeat Field:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBoxDBTableName
        '
        Me.TextBoxDBTableName.Location = New System.Drawing.Point(224, 56)
        Me.TextBoxDBTableName.Name = "TextBoxDBTableName"
        Me.TextBoxDBTableName.Size = New System.Drawing.Size(136, 22)
        Me.TextBoxDBTableName.TabIndex = 57
        Me.TextBoxDBTableName.Text = "members"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(16, 56)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(152, 23)
        Me.Label8.TabIndex = 56
        Me.Label8.Text = "Database Table Name:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(240, 24)
        Me.Label1.TabIndex = 55
        Me.Label1.Text = "Create the registration form metadata:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CheckBoxRegApprove
        '
        Me.CheckBoxRegApprove.Checked = True
        Me.CheckBoxRegApprove.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxRegApprove.Location = New System.Drawing.Point(24, 216)
        Me.CheckBoxRegApprove.Name = "CheckBoxRegApprove"
        Me.CheckBoxRegApprove.Size = New System.Drawing.Size(280, 24)
        Me.CheckBoxRegApprove.TabIndex = 54
        Me.CheckBoxRegApprove.Text = "Registration needs administrator approval"
        '
        'CheckBoxRegLoginCaseSensitive
        '
        Me.CheckBoxRegLoginCaseSensitive.Location = New System.Drawing.Point(24, 168)
        Me.CheckBoxRegLoginCaseSensitive.Name = "CheckBoxRegLoginCaseSensitive"
        Me.CheckBoxRegLoginCaseSensitive.Size = New System.Drawing.Size(176, 24)
        Me.CheckBoxRegLoginCaseSensitive.TabIndex = 53
        Me.CheckBoxRegLoginCaseSensitive.Text = "Login is case sensitive"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(200, 23)
        Me.Label2.TabIndex = 51
        Me.Label2.Text = "Database Table Login Field:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 104)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(200, 23)
        Me.Label7.TabIndex = 49
        Me.Label7.Text = "Database Table Password Field:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnRegForm
        '
        Me.btnRegForm.Location = New System.Drawing.Point(256, 16)
        Me.btnRegForm.Name = "btnRegForm"
        Me.btnRegForm.Size = New System.Drawing.Size(224, 23)
        Me.btnRegForm.TabIndex = 48
        Me.btnRegForm.Text = "Create Registration Form..."
        '
        'LabelStep5
        '
        Me.LabelStep5.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep5.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep5.Location = New System.Drawing.Point(24, 314)
        Me.LabelStep5.Name = "LabelStep5"
        Me.LabelStep5.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep5.TabIndex = 28
        Me.LabelStep5.Text = "5. Build And Generate Files"
        Me.LabelStep5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBoxPasswdMinLen
        '
        Me.TextBoxPasswdMinLen.Location = New System.Drawing.Point(400, 192)
        Me.TextBoxPasswdMinLen.Name = "TextBoxPasswdMinLen"
        Me.TextBoxPasswdMinLen.Size = New System.Drawing.Size(72, 22)
        Me.TextBoxPasswdMinLen.TabIndex = 63
        Me.TextBoxPasswdMinLen.Text = "8"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(224, 192)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(176, 23)
        Me.Label10.TabIndex = 64
        Me.Label10.Text = "Password Minimum Length:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CheckBoxRegPassCaseSensitive
        '
        Me.CheckBoxRegPassCaseSensitive.Checked = True
        Me.CheckBoxRegPassCaseSensitive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxRegPassCaseSensitive.Location = New System.Drawing.Point(24, 192)
        Me.CheckBoxRegPassCaseSensitive.Name = "CheckBoxRegPassCaseSensitive"
        Me.CheckBoxRegPassCaseSensitive.Size = New System.Drawing.Size(192, 24)
        Me.CheckBoxRegPassCaseSensitive.TabIndex = 65
        Me.CheckBoxRegPassCaseSensitive.Text = "Password is case sensitive"
        '
        'FormGenFW
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 559)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDisplay, Me.SplitterInstr, Me.PanelInstr, Me.SplitterStep, Me.PanelStep})
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FormGenFW"
        Me.Text = "Framework"
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.TabControlOptions.ResumeLayout(False)
        Me.GroupBoxInstr.ResumeLayout(False)
        Me.TabPageQuickView.ResumeLayout(False)
        Me.TabPageSource.ResumeLayout(False)
        Me.PanelStepTitle.ResumeLayout(False)
        Me.PanelInstrTitle.ResumeLayout(False)
        Me.PanelStepDetail.ResumeLayout(False)
        Me.PanelDisplay.ResumeLayout(False)
        Me.PanelDispTitle.ResumeLayout(False)
        Me.TabControlDisplay.ResumeLayout(False)
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelStep.ResumeLayout(False)
        Me.PanelInstr.ResumeLayout(False)
        Me.GroupBoxStep.ResumeLayout(False)
        Me.TabPageFileList.ResumeLayout(False)
        Me.TabPageOption1.ResumeLayout(False)
        Me.TabPageOption4.ResumeLayout(False)
        Me.TabPageOption2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.TabPageOption5.ResumeLayout(False)
        Me.TabPageOption3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    ' Treeview nodes

    ' used by Website Framework
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

    Private curIEUrl As String
    Private selectedRadioButton As String

    Private xcLoginForm As String = "<!--xcloginform-->"
    Private xcLogoutForm As String = "<!--xclogoutform-->"

    Private indexFileContent As String
    Private cssFileContent As String
    Private rootIncHeaderFileContent As String
    Private rootIncFooterFileContent As String

    ' Used by Register module.
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

    'Private clsMD As New clsMetadata()
    'Dim metadata As String
    Private regLogin As String
    Private regPassword As String
    Private regIsCaseSensitive As Boolean
    Private regNeedsAdmApproval As Boolean



    Private Sub FormGenFW_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.moduleType = ModuleMain.xcModuleType.xcWebSiteFramework
        Me.NUM_STEPS = 5
        Me.initModuleContent()
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
        str += vbCrLf
        ' Used by Register Module
        str += "' Register Module" & vbCrLf
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
        str += "DBTableName=" & Me.TextBoxDBTableName.Text & vbCrLf
        str += "DBTableFileName=" & Me.nodeDBTableFileName.Text & vbCrLf
        str += "XMLFileName=" & Me.nodeXMLFileName.Text & vbCrLf
        str += vbCrLf
        str += "RegLoginDBField=" & Me.ComboBoxChooseLogin.SelectedItem & vbCrLf
        str += "RegPasswordDBField=" & Me.ComboBoxChoosePassword.SelectedItem & vbCrLf
        str += "RegPassRptDBField=" & Me.ComboBoxChoosePassRpt.SelectedItem & vbCrLf
        str += "RegPasswordMinLength=" & Me.TextBoxPasswdMinLen.Text & vbCrLf
        str += "RegLoginIsCaseSensitive=" & Me.CheckBoxRegLoginCaseSensitive.Checked & vbCrLf
        str += "RegPasswordIsCaseSensitive=" & Me.CheckBoxRegPassCaseSensitive.Checked & vbCrLf
        str += "RegNeedsAdmApproval=" & Me.CheckBoxRegApprove.Checked & vbCrLf
        str += "RegMetadata=" & Me.moduleName & ".metadata.txt" & vbCrLf
        str += vbCrLf
        ' Start - Table attributes of the default layout
        str += "' Table attributes of the default layout" & vbCrLf
        str += "TableBorder=" & Me.tblAttr.border & vbCrLf
        str += "TableWidth=" & Me.tblAttr.width & vbCrLf
        str += "TableBgColor=" & Me.tblAttr.bgColor & vbCrLf
        str += "TableForeColor=" & Me.tblAttr.foreColor & vbCrLf
        str += "TableFirstColWidth=" & Me.tblAttr.firstColWidth & vbCrLf
        str += "TableRows=" & Me.tblAttr.rows & vbCrLf
        str += "TableCols=" & Me.tblAttr.cols & vbCrLf
        str += "TableCellPadding=" & Me.tblAttr.cellPadding & vbCrLf
        str += "TableCellSpacing=" & Me.tblAttr.cellSpacing & vbCrLf
        ' End - Table attributes of the default layout
        str += vbCrLf

        IOManager.SaveTextToFile(str, Me.moduleFullPath)

        IOManager.SaveTextToFile(Me.getIndexFile, IOManager.getFolder(Me.moduleFullPath) & Me.moduleName & ".index.txt")
        IOManager.SaveTextToFile(Me.getCssFile, IOManager.getFolder(Me.moduleFullPath) & Me.moduleName & ".css.txt")
        IOManager.SaveTextToFile(Me.getRootIncHeaderFile, IOManager.getFolder(Me.moduleFullPath) & Me.moduleName & ".header.txt")
        IOManager.SaveTextToFile(Me.getRootIncFooterFile(), IOManager.getFolder(Me.moduleFullPath) & Me.moduleName & ".footer.txt")
        IOManager.SaveTextToFile(Me.metadata, IOManager.getFolder(Me.moduleFullPath) & Me.moduleName & ".metadata.txt")

        Me.infoSaved = True
    End Sub


    Protected Overrides Sub initTreeNodesText()
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

        ' Used by Register Module
        nodeRegFolder.Text = "register"
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
        nodeDBTableFileName.Text = "register.sql"
        nodeXMLFileName.Text = "register.xml"
    End Sub


    ' Use moduleInfoLines to initialized module.
    Protected Overrides Sub initModuleContent()
        Dim loginVal, passwordVal, passRptVal As String

        tmpFileFolder = SetupManager.ROOT_PATH & "tmp\" & Me.projectName & "\" & Me.moduleName & "\"
        tmpFileSubFolder = SetupManager.ROOT_PATH & "tmp\" & Me.projectName & "\" & Me.moduleName & "\sub\"
        tmpImageFolder = SetupManager.ROOT_PATH & "tmp\" & Me.projectName & "\" & Me.moduleName & "\images\"
        IOManager.CreateFolder(tmpFileFolder)
        IOManager.CreateFolder(tmpFileSubFolder)
        IOManager.CreateFolder(tmpImageFolder)
        'MsgBox(Me.moduleName & ", tmpfilefolder=" & tmpFileFolder & ", tmpImagefolder=" & tmpImageFolder)

        Me.LabelOutputPath.Text = Me.projectOutputPath
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
                    ' Start - Used by Register module
                ElseIf lcase_line.StartsWith("registerFolder") Then
                    Me.nodeRegFolder.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("regfile") Then
                    Me.nodeReg.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("regconfirmfile") Then
                    Me.nodeRegConfirm.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("memregeditfile") Then
                    Me.nodeMemRegEdit.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("memregconfirmfile") Then
                    Me.nodeMemRegConfirm.Text = MyUtil.getStrAfterDelimiter(line, "=")
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
                ElseIf lcase_line.StartsWith("dbtablename") Then
                    Me.TextBoxDBTableName.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("dbtablefilename") Then
                    Me.nodeDBTableFileName.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("xmlfilename") Then
                    Me.nodeXMLFileName.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("reglogindbfield") Then
                    loginVal = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("regpassworddbfield") Then
                    passwordVal = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("regpassrptdbfield") Then
                    passRptVal = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("regpasswordminlength") Then
                    Me.TextBoxPasswdMinLen.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("regloginiscasesensitive") Then
                    Me.CheckBoxRegLoginCaseSensitive.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("regpasswordiscasesensitive") Then
                    Me.CheckBoxRegPassCaseSensitive.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("regneedsadmapproval") Then
                    Me.CheckBoxRegApprove.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("regmetadata") Then
                    Me.metadata = (IOManager.GetFileContents(IOManager.getFolder(Me.moduleFullPath) & MyUtil.getStrAfterDelimiter(line, "=")))
                    ' Used by default layout table attributes
                ElseIf lcase_line.StartsWith("tableborder") Then
                    Me.tblAttr.border = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("tablewidth") Then
                    Me.tblAttr.width = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("tablebgcolor") Then
                    Me.tblAttr.bgColor = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("tableforecolor") Then
                    Me.tblAttr.foreColor = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("tablefirstcolwidth") Then
                    Me.tblAttr.firstColWidth = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("tablerows") Then
                    Me.tblAttr.rows = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("tablecols") Then
                    Me.tblAttr.cols = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("tablecellpadding") Then
                    Me.tblAttr.cellPadding = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("tablecellspacing") Then
                    Me.tblAttr.cellSpacing = MyUtil.getStrAfterDelimiter(line, "=")
                    ' End - used by Register Module
                End If
            Next
        End If

        constructTreeViewFileList()
        'MsgBox("initContent - processStep. stepNumber = " & stepnumber)
        Me.processStep()
        'Me.TabControlOptions.SelectedTab = _
        '   CType(MyUtil.getControlFromName(Me.TabControlOptions, "TabpageOption" & stepNumber), TabPage)
        Me.setLoginPassInfo(loginVal, passwordVal, passRptVal)
    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Protected Overrides Sub constructTreeViewFileList()
        Dim TreeViewFileList As TreeView = Me.TreeView1
        If TreeViewFileList Is Nothing Then Exit Sub

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

            Dim rootNode As New TreeNode(projectOutputPath, 2, 2)

            rootNode.Nodes.Add(Me.indexFileNode)
            rootNode.Nodes.Add(Me.cssFileNode)
            rootNode.Nodes.Add(Me.adminFolderNode)
            rootNode.Nodes.Add(Me.aspBinFolderNode)
            rootNode.Nodes.Add(Me.loginFolderNode)
            rootNode.Nodes.Add(Me.memberFolderNode)
            rootNode.Nodes.Add(Me.rootIncFolderNode)
            rootNode.Nodes.Add(Me.imagesFolderNode)

            ' Used by Register Module
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

            rootNode.Nodes.Add(nodeRegFolder)
            ' End - used by Register Module

            rootNode.ExpandAll()
            TreeViewFileList.Nodes.Add(rootNode)

        Catch ex As Exception
            MsgBox("ConstructTreeViewFileList Error:" & vbCrLf & ex.ToString)
            End
        End Try

    End Sub


    Private Sub TreeView1_BeforeLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TreeView1.BeforeLabelEdit
        e.CancelEdit = (e.Node.Text = Me.projectOutputPath)
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' These are functions to get the real contents for each file
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''
    ' Generate Web Site Framework
    '
    Protected Friend Overrides Sub buildModule( _
    Optional ByVal frameworkParams As clsWebsiteFrameworkParameters = Nothing)

        'MsgBox("validate database information")
        If Trim(Me.TextBoxODBCName.Text) = "" Or _
           Trim(Me.TextBoxLogin.Text) = "" Or _
           Trim(Me.TextBoxPassword.Text) = "" Then
            MsgBox("If you leave ODBC information empty, you will need to manually edit file " & Me.aspBinFolderNode.Text & "/" & Me.aspBinDBConfigFileNode.Text & " later.")
        End If

        If Trim(Me.metadata) = "" Then
            MsgBox("metadata cannot be empty")
            Exit Sub
        End If

        If Trim(Me.TextBoxDBTableName.Text) = "" Or Trim(Me.ComboBoxChooseLogin.SelectedItem) = "" Or _
           Trim(Me.ComboBoxChoosePassword.SelectedItem) = "" Or Trim(Me.ComboBoxChoosePassRpt.SelectedItem) = "" Then
            MsgBox("Database Table Name, login, password and repeat password fields cannot be empty.")
            Me.TabControlOptions.SelectedTab = Me.TabPageOption3
            Exit Sub
        End If

        If Trim(Me.ComboBoxChoosePassword.SelectedItem) = Trim(Me.ComboBoxChoosePassRpt.SelectedItem) Then
            MsgBox("Password and Password repetition fields cannot be the same.")
            Exit Sub
        End If

        Try
            Dim passMinLen As Integer = CType(Me.TextBoxPasswdMinLen.Text, Integer)
        Catch ex As Exception
            MsgBox("Password length field must be integer: " & ex.Message)
            Me.TextBoxPasswdMinLen.Select(0, Len(Me.TextBoxPasswdMinLen.Text.Length))
            Exit Sub
        End Try

        ' Clear generated files of last time.
        If IOManager.folderExists(projectOutputPath) Then
            IOManager.DeleteFolder(projectOutputPath, True)
        End If
        IOManager.CreateFolder(projectOutputPath)

        Dim fw As New clsModuleWebsiteFramework()
        fw.init(Me.indexFileNode.Text, Me.cssFileNode.Text, _
                Me.loginFolderNode.Text, Me.rootIncFolderNode.Text, _
                Me.rootIncHeaderFileNode.Text, Me.rootIncFooterFileNode.Text, _
                Me.loginAdminSecFileNode.Text, Me.loginMemSecFileNode.Text, _
                Me.loginLoginFileNode.Text, Me.loginLogoutFileNode.Text, _
                Me.memberFolderNode.Text, Me.memberHomeFileNode.Text, _
                Me.aspBinFolderNode.Text, Me.aspBinAspFuncLibFileNode.Text, _
                Me.aspBinDBConfigFileNode.Text, Me.aspBinAdovbsFileNode.Text, _
                Me.adminFolderNode.Text, Me.adminHomeFileNode.Text, Me.adminApproveFileNode.Text, _
                Me.imagesFolderNode.Text, Me.TextBoxODBCName.Text, _
                Me.TextBoxLogin.Text, Me.TextBoxPassword.Text, _
                Me.getIndexFile(), Me.getCssFile(), _
                Me.getRootIncHeaderFile(), Me.getRootIncFooterFile(), _
                Me.nodeRegFolder.Text & "/" & Me.nodeReg.Text, _
                "../" & Me.nodeRegFolder.Text & "/" & Me.nodeMemRegConfirm.Text, _
                "../" & Me.nodeRegFolder.Text & "/" & Me.nodeAdmRegList.Text)

        Dim progressBar As New FormProgressBar()
        progressBar.setText(Me.moduleName & SetupManager.MODULE_FILE_EXTENSION)
        progressBar.Show()

        If projectOutputPath.EndsWith("\") = False Then
            projectOutputPath &= "\"
        End If

        Dim adminFolder As String = Me.adminFolderNode.Text & "\"
        IOManager.CreateFolder(projectOutputPath & adminFolder)
        IOManager.SaveTextToFile(fw.getAdminHomeFile, _
                  projectOutputPath & adminFolder & Me.adminHomeFileNode.Text)
        IOManager.SaveTextToFile(fw.getAdminApproveFile, _
                  projectOutputPath & adminFolder & Me.adminApproveFileNode.Text)

        progressBar.setProgress(1)

        Dim aspBinFolder As String = Me.aspBinFolderNode.Text & "\"
        IOManager.CreateFolder(projectOutputPath & aspBinFolder)
        IOManager.SaveTextToFile(fw.getAspBinAdovbsFile, _
                  projectOutputPath & aspBinFolder & Me.aspBinAdovbsFileNode.Text)
        IOManager.SaveTextToFile(fw.getAspBinAspFuncLibFile, _
                  projectOutputPath & aspBinFolder & Me.aspBinAspFuncLibFileNode.Text)
        IOManager.SaveTextToFile(fw.getAspBinDBConfigFile(), _
                  projectOutputPath & aspBinFolder & Me.aspBinDBConfigFileNode.Text)

        progressBar.setProgress(1)

        Dim rootIncFolder As String = Me.rootIncFolderNode.Text & "\"
        IOManager.CreateFolder(projectOutputPath & rootIncFolder)
        IOManager.SaveTextToFile(fw.getRootIncHeaderFile, _
                  projectOutputPath & rootIncFolder & Me.rootIncHeaderFileNode.Text)
        IOManager.SaveTextToFile(fw.getRootIncFooterFile, _
                  projectOutputPath & rootIncFolder & Me.rootIncFooterFileNode.Text)

        progressBar.setProgress(1)

        Dim memFolder As String = Me.memberFolderNode.Text & "\"
        IOManager.CreateFolder(projectOutputPath & memFolder)
        IOManager.SaveTextToFile(fw.getMemHomeFile, _
                  projectOutputPath & memFolder & Me.memberHomeFileNode.Text)

        progressBar.setProgress(1)

        Dim loginFolder As String = Me.loginFolderNode.Text & "\"
        IOManager.CreateFolder(projectOutputPath & loginFolder)
        IOManager.SaveTextToFile(fw.getLoginLoginFile, _
                  projectOutputPath & loginFolder & Me.loginLoginFileNode.Text)
        IOManager.SaveTextToFile(fw.getLoginLogoutFile, _
                  projectOutputPath & loginFolder & Me.loginLogoutFileNode.Text)
        IOManager.SaveTextToFile(fw.getLoginAdminSecFile, _
                  projectOutputPath & loginFolder & Me.loginAdminSecFileNode.Text)
        IOManager.SaveTextToFile(fw.getLoginMemSecFile, _
                  projectOutputPath & loginFolder & Me.loginMemSecFileNode.Text)

        progressBar.setProgress(1)

        Dim imageFolder As String = Me.imagesFolderNode.Text & "\"
        IOManager.CreateFolder(projectOutputPath & imageFolder)
        ' If there are uploaded image files, copy to target folder
        If IOManager.folderExists(tmpImageFolder) Then
            Dim images As String() = IO.Directory.GetFiles(tmpImageFolder)
            Dim image As String
            For Each image In images
                IO.File.Copy(image, projectOutputPath & imageFolder & IOManager.getFile(image), True)
            Next
        End If

        IOManager.SaveTextToFile(fw.getCssFile, projectOutputPath & Me.cssFileNode.Text)
        IOManager.SaveTextToFile(fw.getIndexFile, projectOutputPath & Me.indexFileNode.Text)

        progressBar.setProgress(1)

        ' Create Register Module
        Me.generateLoginModule()

        progressBar.setProgress(4)

        progressBar.Close()
    End Sub


    Private Sub generateLoginModule()

        Dim reg As New clsModuleLogin()
        reg.init(Me.metadata, "../" & Me.indexFileNode.Text, _
                 "../" & Me.aspBinFolderNode.Text & "/" & Me.aspBinAspFuncLibFileNode.Text, _
                 "../" & Me.loginFolderNode.Text & "/" & Me.loginAdminSecFileNode.Text, _
                 "../" & Me.loginFolderNode.Text & "/" & Me.loginMemSecFileNode.Text, _
                 "../" & Me.rootIncFolderNode.Text & "/" & Me.rootIncHeaderFileNode.Text, _
                 "../" & Me.rootIncFolderNode.Text & "/" & Me.rootIncFooterFileNode.Text, _
                 "../" & Me.adminFolderNode.Text & "/" & Me.adminHomeFileNode.Text, _
                 "../" & Me.memberFolderNode.Text & "/" & Me.memberHomeFileNode.Text, _
                 "./" & Me.nodeIncFolder.Text & "/" & Me.nodeIncFuncs.Text, _
                 "./" & Me.nodeIncFolder.Text & "/" & Me.nodeIncEditTable.Text, _
                 "./" & Me.nodeIncFolder.Text & "/" & Me.nodeIncConfirmTable.Text, _
                 Me.nodeReg.Text, Me.nodeRegConfirm.Text, _
                 Me.nodeMemRegEdit.Text, Me.nodeMemRegConfirm.Text, _
                 Me.nodeMemChgPassword.Text, _
                 Me.nodeAdmChgPassword.Text, Me.nodeAdmRegList.Text, _
                 Me.nodeAdmRegEdit.Text, Me.nodeAdmRegConfirm.Text, _
                 Me.TextBoxDBTableName.Text, Me.ComboBoxChooseLogin.SelectedItem, _
                 Me.ComboBoxChoosePassword.SelectedItem, Me.ComboBoxChoosePassRpt.SelectedItem, _
                 CType(Me.TextBoxPasswdMinLen.Text, Integer), _
                 Me.CheckBoxRegLoginCaseSensitive.Checked, _
                 Me.CheckBoxRegPassCaseSensitive.Checked, _
                 Me.CheckBoxRegApprove.Checked)

        'nodeRegFolder.Text = "memRegFolder"
        Dim folderPath As String = projectOutputPath & Me.nodeRegFolder.Text & "\"
        IOManager.CreateFolder(folderPath)
        'MsgBox(folderPath)

        IOManager.SaveTextToFile(reg.getRegFile(), folderPath & Me.nodeReg.Text)
        IOManager.SaveTextToFile(reg.getRegConfirmFile(), folderPath & Me.nodeRegConfirm.Text)
        IOManager.SaveTextToFile(reg.getMemRegEditFile, folderPath & Me.nodeMemRegEdit.Text)
        IOManager.SaveTextToFile(reg.getMemRegConfirmFile, folderPath & Me.nodeMemRegConfirm.Text)
        IOManager.SaveTextToFile(reg.getMemChgPasswordFile, folderPath & Me.nodeMemChgPassword.Text)
        IOManager.SaveTextToFile(reg.getAdmRegListFile, folderPath & Me.nodeAdmRegList.Text)
        IOManager.SaveTextToFile(reg.getAdmRegEditFile, folderPath & Me.nodeAdmRegEdit.Text)
        IOManager.SaveTextToFile(reg.getAdmRegConfirmFile, folderPath & Me.nodeAdmRegConfirm.Text)
        IOManager.SaveTextToFile(reg.getAdmChgPasswordFile, folderPath & Me.nodeAdmChgPassword.Text)

        'IOManager.SaveTextToFile(reg.getSQLFile, folderPath & Me.nodeDBTableFileName.Text)
        IOManager.SaveTextToFile(reg.getRegSQLFile, folderPath & Me.nodeDBTableFileName.Text)
        IOManager.SaveTextToFile(reg.getXMLFile, folderPath & Me.nodeXMLFileName.Text)

        Dim subFolderPath As String = folderPath & Me.nodeIncFolder.Text & "\"
        IOManager.CreateFolder(subFolderPath)

        IOManager.SaveTextToFile(reg.getIncEditTableFile(Me.tblAttr), subFolderPath & Me.nodeIncEditTable.Text)
        IOManager.SaveTextToFile(reg.getIncConfirmTableFile(Me.tblAttr), subFolderPath & Me.nodeIncConfirmTable.Text)
        IOManager.SaveTextToFile(reg.getIncFuncsFile, subFolderPath & Me.nodeIncFuncs.Text)

        ' The is the login file that uses database.
        Dim loginFilePath As String = projectOutputPath & _
            Me.loginFolderNode.Text & "\" & Me.loginLoginFileNode.Text
        IOManager.SaveTextToFile(reg.getLoginLoginFile(), loginFilePath)
        Dim approveAdminFilePath As String = projectOutputPath & _
            Me.adminFolderNode.Text & "\" & Me.adminApproveFileNode.Text
        IOManager.SaveTextToFile(reg.getApproveAdminFile(), approveAdminFilePath)
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
        'MsgBox("set step lables color. stepnumber = " & stepnumber)
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


    Private Sub ButtonCreateDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCreateDB.Click
        Me.curIEUrl = SetupManager.WEBSITE_LIB_PATH & "dbsetup.html"
        AxWebBrowser1.Navigate(Me.curIEUrl)
        Me.TabControlDisplay.SelectedTab = Me.TabPageQuickView
    End Sub


    Private Sub TabControlOptions_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControlOptions.SelectedIndexChanged
        Select Case Me.TabControlOptions.SelectedIndex
            Case 0
                Me.stepNumber = 1
                Me.LabelStepInstr.Text = "You should create the database for your website first. Follow the instruction in Instruction Explorer."
                Me.LabelInstr.Text = "If you leave ODBC name, login or password empty for now, you will need to manually edit database configuration file later."
                Me.TabControlDisplay.SelectedTab = Me.TabPageQuickView
            Case 1
                Me.stepNumber = 2
                Me.TabControlDisplay.SelectedTab = Me.TabPageFileList
                Me.LabelStepInstr.Text = "Change the user interface files. These include the website homepage file, and website global header and footer files."
                Me.TabControlDisplay.SelectedTab = Me.TabPageSource
            Case 2
                Me.stepNumber = 3
                Me.LabelStepInstr.Text = "Add register form so a user can register as a new member."
                'Me.RichTextBox1.Text = clsMD.getMetadataExample()
                Me.RichTextBox1.Text = Me.metadata
                Me.LabelInstr.Text = "Enter metadata, or use the wizard to create the metadata (to be constructed)."
                Me.TabControlDisplay.SelectedTab = Me.TabPageSource
            Case 3
                Me.stepNumber = 4
                Me.LabelStepInstr.Text = "You don't really need to change filenames for the module to work. But in case you prefer other names, you can do it."
                Me.LabelInstr.Text = ""
                Me.TabControlDisplay.SelectedTab = Me.TabPageFileList
            Case 4
                Me.stepNumber = 5
                Me.LabelStepInstr.Text = "The shortcut to build project is F5. Only opening modules in a project are built. The shortcut to build only the currently active module is F6."
                Me.LabelInstr.Text = "After building the module, read the readme.txt file in the module folder about what to do."
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

    Private Sub btnOpenOutputFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenOutputFolder.Click
        If IOManager.folderExists(Me.projectOutputPath) Then
            Dim myProcess As New Process()
            myProcess = System.Diagnostics.Process.Start(Me.projectOutputPath)
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

    Private Sub btnUseDefaultFileNames_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUseDefaultFileNames.Click
        Me.initTreeNodesText()
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions to change file contents
    ''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub RadioButtonIndex_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonIndex.CheckedChanged
        If Me.RadioButtonIndex.Checked = True Then
            'Me.RichTextBox1.Text = Me.getIndexFile()
            Me.RichTextBox1.Text = Me.indexFileContent
            Me.TabControlDisplay.SelectedTab = Me.TabPageSource
            Me.btnChgContent.Enabled = True
            Me.btnView.Enabled = True
            Me.LabelInstr.Text = "Insert " & Chr(34) & Me.xcLoginForm & Chr(34) & " at where you want the login form to appear."
        End If
    End Sub

    Private Sub RadioButtonCss_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonCss.CheckedChanged
        If Me.RadioButtonCss.Checked = True Then
            Me.RichTextBox1.Text = Me.getCssFile()
            Me.TabControlDisplay.SelectedTab = Me.TabPageSource
            Me.btnChgContent.Enabled = True
            Me.btnView.Enabled = False
            Me.LabelInstr.Text = ""
        End If
    End Sub

    Private Sub RadioButtonHeader_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonHeader.CheckedChanged
        If Me.RadioButtonHeader.Checked = True Then
            'Me.RichTextBox1.Text = Me.getRootIncHeaderFile()
            Me.RichTextBox1.Text = Me.rootIncHeaderFileContent
            Me.TabControlDisplay.SelectedTab = Me.TabPageSource
            Me.btnChgContent.Enabled = True
            Me.btnView.Enabled = True
            Me.LabelInstr.Text = "Insert " & Chr(34) & Me.xcLogoutForm & Chr(34) & " at where you want the logout form to appear."
        End If
    End Sub

    Private Sub RadioButtonFooter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonFooter.CheckedChanged
        If Me.RadioButtonFooter.Checked = True Then
            'Me.RichTextBox1.Text = Me.getRootIncFooterFile()
            Me.RichTextBox1.Text = Me.rootIncFooterFileContent
            Me.TabControlDisplay.SelectedTab = Me.TabPageSource
            Me.btnChgContent.Enabled = True
            Me.btnView.Enabled = True
            Me.LabelInstr.Text = "Insert " & Chr(34) & Me.xcLogoutForm & Chr(34) & " at where you want the logout form to appear."
        End If
    End Sub

    Private Sub btnChgContent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChgContent.Click
        Me.saveFileContent()
    End Sub


    Private Sub saveFileContent()
        If Me.RadioButtonIndex.Checked = True Then
            Me.setIndexFile(Me.RichTextBox1.Text)
            Me.btnView.Text = "View Homepage As HTML"
            selectedRadioButton = Me.RadioButtonIndex.Name
        ElseIf Me.RadioButtonCss.Checked = True Then
            Me.setCssFile(Me.RichTextBox1.Text)
            Me.btnView.Text = "View File As HTML"
            selectedRadioButton = Me.RadioButtonCss.Name
        ElseIf Me.RadioButtonHeader.Checked = True Then
            Me.setRootIncHeaderFile(Me.RichTextBox1.Text)
            Me.btnView.Text = "View Header And Footer As HTML"
            selectedRadioButton = Me.RadioButtonHeader.Name
        ElseIf Me.RadioButtonFooter.Checked = True Then
            Me.setRootIncFooterFile(Me.RichTextBox1.Text)
            Me.btnView.Text = "View Header And Footer As HTML"
            selectedRadioButton = Me.RadioButtonFooter.Name
        End If
        Me.btnChgContent.Enabled = False
    End Sub


    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
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



    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
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


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''' content of files
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Private Sub btnRegForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRegForm.Click
        Me.metadata = Me.RichTextBox1.Text

        Dim oldLoginVal As String = Me.ComboBoxChooseLogin.SelectedItem
        Dim oldPasswordVal As String = Me.ComboBoxChoosePassword.SelectedItem
        Dim oldPassRptVal As String = Me.ComboBoxChoosePassRpt.SelectedItem

        Dim mdwizard As New FormMetadataWizard()
        mdwizard.init(Me.tmpFileSubFolder, Me.metadata, Me.tblAttr, Me)
        mdwizard.ShowDialog(Me)

        Me.RichTextBox1.Text = Me.metadata
        Me.setLoginPassInfo(oldLoginVal, oldPasswordVal, oldPassRptVal)
    End Sub


    Private Sub setLoginPassInfo( _
        Optional ByVal loginVal As String = "", _
        Optional ByVal passwordVal As String = "", _
        Optional ByVal passRptVal As String = "")

        ' get login/password fields.
        clsmd.init(Me.metadata)
        Dim i As Integer

        Me.ComboBoxChooseLogin.Items.Clear()
        Me.ComboBoxChoosePassword.Items.Clear()
        Me.ComboBoxChoosePassRpt.Items.Clear()

        For i = 0 To Me.clsMD.ArrayListTextFields.Count - 1
            Me.ComboBoxChooseLogin.Items.Add(clsMD.ArrayListTextFields(i))
            If clsmd.ArrayListTextFields(i) = loginVal Then Me.ComboBoxChooseLogin.SelectedIndex = i
        Next
        For i = 0 To Me.clsMD.ArrayListPasswordFields.Count - 1
            Me.ComboBoxChoosePassword.Items.Add(clsmd.ArrayListPasswordFields(i))
            If clsMD.ArrayListPasswordFields(i) = passwordVal Then Me.ComboBoxChoosePassword.SelectedIndex = i
            Me.ComboBoxChoosePassRpt.Items.Add(clsmd.ArrayListPasswordFields(i))
            If clsMD.ArrayListPasswordFields(i) = passRptVal Then Me.ComboBoxChoosePassRpt.SelectedIndex = i
        Next

    End Sub


    'Private rootIncHeaderFileNode As New TreeNode("header.asp", 0, 0)
    Public Function getRootIncHeaderFile() As String
        Dim fileStr As String = Me.rootIncHeaderFileContent
        fileStr = Replace(fileStr, Me.xcLogoutForm, Me.getDefaultLogoutForm(), 1, 1)
        Return Replace(fileStr, Me.xcLoginForm, Me.getDefaultLoginForm(), 1, 1)
    End Function

    Private Sub setRootIncHeaderFile(ByVal content As String)
        Me.rootIncHeaderFileContent = Replace(content, Me.getDefaultLogoutForm, Me.xcLogoutForm, 1, 1)
        Me.rootIncHeaderFileContent = Replace(rootIncHeaderFileContent, Me.getDefaultLoginForm, Me.xcLoginForm, 1, 1)
    End Sub

    'Private rootIncFooterFileNode As New TreeNode("footer.asp", 0, 0)
    Public Function getRootIncFooterFile() As String
        Dim fileStr As String = Me.rootIncFooterFileContent
        fileStr = Replace(fileStr, Me.xcLogoutForm, Me.getDefaultLogoutForm(), 1, 1)
        Return Replace(fileStr, Me.xcLoginForm, Me.getDefaultLoginForm(), 1, 1)
    End Function

    Private Sub setRootIncFooterFile(ByVal content As String)
        Me.rootIncFooterFileContent = Replace(content, Me.getDefaultLogoutForm, Me.xcLogoutForm, 1, 1)
        Me.rootIncFooterFileContent = Replace(rootIncFooterFileContent, Me.getDefaultLoginForm, Me.xcLoginForm, 1, 1)
    End Sub


    Public Function getDefaultLogoutForm()
        Dim str As String
        str = str & "<%if session(" & Chr(34) & "_isMem" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
        str = str & "    <form name=" & Chr(34) & "logoffForm" & Chr(34) & " method=" & Chr(34) & "post" & Chr(34) & " action=" & Chr(34) & "../" & Me.loginFolderNode.Text & "/" & Me.loginLogoutFileNode.Text & Chr(34) & ">" & vbCrLf
        str = str & "    <input type=" & Chr(34) & "submit" & Chr(34) & " name=" & Chr(34) & "btnLogoff" & Chr(34) & " value=" & Chr(34) & "Logoff" & Chr(34) & ">" & vbCrLf
        str = str & "    </form>" & vbCrLf
        str = str & "<%End IF%>" & vbCrLf
        str = str & vbCrLf
        Return str
    End Function

    Public Function getDefaultLoginForm()
        Dim str As String
        str = str & "<%" & vbCrLf
        str = str & "if session(" & Chr(34) & "_isMem" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	Response.Redirect " & Chr(34) & "./" & Me.memberFolderNode.Text & "/" & Me.memberHomeFileNode.Text & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<center>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--xcLoginForm.start-->" & vbCrLf
        str = str & "<form name=" & Chr(34) & "loginForm" & Chr(34) & " method=" & Chr(34) & "post" & Chr(34) & " action=" & Chr(34) & "./" & Me.loginFolderNode.Text & "/" & Me.loginLoginFileNode.Text & Chr(34) & ">" & vbCrLf
        str = str & "	<table border=0>" & vbCrLf
        str = str & "	<tr>" & vbCrLf
        str = str & "	<td align=right>Account</td>" & vbCrLf
        str = str & "	<td align=right><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "usrLoginname" & Chr(34) & " size=15 value=<%=request(" & Chr(34) & "login" & Chr(34) & ")%>></td>" & vbCrLf
        str = str & "	</tr>" & vbCrLf
        str = str & "	<tr>" & vbCrLf
        str = str & "	<td align=right>Password</td>" & vbCrLf
        str = str & "	<td align=right><input type=" & Chr(34) & "password" & Chr(34) & " name=" & Chr(34) & "usrPassword" & Chr(34) & " size=13></td>" & vbCrLf
        str = str & "	</tr>" & vbCrLf
        str = str & "	<tr>" & vbCrLf
        str = str & "	<td colspan=2 align=right>" & vbCrLf
        str = str & "	<input type=" & Chr(34) & "submit" & Chr(34) & " name=" & Chr(34) & "btnLogin" & Chr(34) & " value=" & Chr(34) & "Login" & Chr(34) & "></td>" & vbCrLf
        str = str & "	</tr>" & vbCrLf
        str = str & "	</table>" & vbCrLf
        str = str & "</form>" & vbCrLf
        str = str & "<p><font color=red><%=request(" & Chr(34) & "error" & Chr(34) & ")%></font></p>" & vbCrLf
        str = str & vbCrLf
        str = str & "<script language=" & Chr(34) & "javascript" & Chr(34) & ">" & vbCrLf
        str = str & "<!--" & vbCrLf
        str = str & "document.loginForm.usrLoginname.focus();" & vbCrLf
        str = str & "-->" & vbCrLf
        str = str & "</script>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--xcLoginForm.end-->" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "</center>" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        Return str
    End Function


    'Private indexFileNode As New TreeNode("index.asp", 0, 0)
    Public Function getIndexFile() As String
        Dim fileStr As String = Me.indexFileContent
        Return Replace(fileStr, Me.xcLoginForm, Me.getDefaultLoginForm(), 1, 1)
    End Function

    Private Sub setIndexFile(ByVal content As String)
        Me.indexFileContent = Replace(content, Me.getDefaultLoginForm, Me.xcLoginForm, 1, 1)
    End Sub

    'Private cssFileNode As New TreeNode("xc.asp", 0, 0)
    Public Function getCssFile() As String
        'Dim str As String = ""
        Return Me.cssFileContent
    End Function

    Private Sub setCssFile(ByVal content As String)
        Me.cssFileContent = content
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged
        Me.btnRegForm.Enabled = True
    End Sub


    ' Used by FormMain to read website framework settings.
    Friend Sub getFrameworkParams( _
        ByRef frameworkParams As clsWebsiteFrameworkParameters)
        frameworkParams.homeFile = Me.indexFileNode.Text
        frameworkParams.cssFile = Me.cssFileNode.Text
        frameworkParams.rootIncFolder = Me.rootIncFolderNode.Text
        frameworkParams.rootIncHeaderFile = Me.rootIncHeaderFileNode.Text
        frameworkParams.rootIncFooterFile = Me.rootIncFooterFileNode.Text
        frameworkParams.admFolder = Me.adminFolderNode.Text
        frameworkParams.admHomeFile = Me.adminHomeFileNode.Text
        frameworkParams.memFolder = Me.memberFolderNode.Text
        frameworkParams.memHomeFile = Me.memberHomeFileNode.Text
        frameworkParams.loginFolder = Me.loginFolderNode.Text
        frameworkParams.loginAdmSecurityFile = Me.loginAdminSecFileNode.Text
        frameworkParams.loginMemSecurityFile = Me.loginMemSecFileNode.Text
        frameworkParams.aspBinFolder = Me.aspBinFolderNode.Text
        frameworkParams.aspFuncLibFile = Me.aspBinAspFuncLibFileNode.Text
        frameworkParams.imageFolder = Me.imagesFolderNode.Text
    End Sub

End Class
