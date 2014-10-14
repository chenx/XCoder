

Public Class FormGenUpload
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
    Friend WithEvents TabPageOption1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption4 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBoxRestriction As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBoxFileSize As System.Windows.Forms.CheckBox
    Friend WithEvents TextBoxFileSize As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBoxFileType As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxMultiUpload As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBoxOrder As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton5 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents CheckBoxAllowSearch As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents RadioButtonNonMemAccessable As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton8 As System.Windows.Forms.RadioButton
    Friend WithEvents CheckBoxNonMemView As System.Windows.Forms.CheckBox
    Friend WithEvents btnUseDefaultFileNames As System.Windows.Forms.Button
    Friend WithEvents LabelSetFileNames As System.Windows.Forms.Label
    Friend WithEvents btnOpenOutputFolder As System.Windows.Forms.Button
    Friend WithEvents LabelOutputPath As System.Windows.Forms.Label
    Friend WithEvents LabelAfterBuild As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LabelStep4 As System.Windows.Forms.Label
    Friend WithEvents LabelStep3 As System.Windows.Forms.Label
    Friend WithEvents LabelStep2 As System.Windows.Forms.Label
    Friend WithEvents LabelStep1 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxFileType As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormGenUpload))
        Me.TabPageOption1 = New System.Windows.Forms.TabPage()
        Me.CheckBoxAllowSearch = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.GroupBoxOrder = New System.Windows.Forms.GroupBox()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.GroupBoxRestriction = New System.Windows.Forms.GroupBox()
        Me.CheckBoxMultiUpload = New System.Windows.Forms.CheckBox()
        Me.TextBoxFileType = New System.Windows.Forms.TextBox()
        Me.CheckBoxFileType = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxFileSize = New System.Windows.Forms.TextBox()
        Me.CheckBoxFileSize = New System.Windows.Forms.CheckBox()
        Me.TabPageOption2 = New System.Windows.Forms.TabPage()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonNonMemAccessable = New System.Windows.Forms.RadioButton()
        Me.RadioButton8 = New System.Windows.Forms.RadioButton()
        Me.CheckBoxNonMemView = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.TabPageOption3 = New System.Windows.Forms.TabPage()
        Me.btnUseDefaultFileNames = New System.Windows.Forms.Button()
        Me.LabelSetFileNames = New System.Windows.Forms.Label()
        Me.TabPageOption4 = New System.Windows.Forms.TabPage()
        Me.btnOpenOutputFolder = New System.Windows.Forms.Button()
        Me.LabelOutputPath = New System.Windows.Forms.Label()
        Me.LabelAfterBuild = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.LabelStep4 = New System.Windows.Forms.Label()
        Me.LabelStep3 = New System.Windows.Forms.Label()
        Me.LabelStep2 = New System.Windows.Forms.Label()
        Me.LabelStep1 = New System.Windows.Forms.Label()
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
        Me.GroupBox1.SuspendLayout()
        Me.GroupBoxOrder.SuspendLayout()
        Me.GroupBoxRestriction.SuspendLayout()
        Me.TabPageOption2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPageOption3.SuspendLayout()
        Me.TabPageOption4.SuspendLayout()
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
        Me.RichTextBox1.Size = New System.Drawing.Size(532, 474)
        Me.RichTextBox1.Visible = True
        '
        'ImageListChild
        '
        Me.ImageListChild.ImageStream = CType(resources.GetObject("ImageListChild.ImageStream"), System.Windows.Forms.ImageListStreamer)
        '
        'PanelStepDetail
        '
        Me.PanelStepDetail.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelStep4, Me.LabelStep3, Me.LabelStep2, Me.LabelStep1})
        Me.PanelStepDetail.Size = New System.Drawing.Size(248, 532)
        Me.PanelStepDetail.Visible = True
        '
        'LabelInstrTitle
        '
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
        'TabPageSource
        '
        Me.TabPageSource.Size = New System.Drawing.Size(532, 474)
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
        Me.TabControlOptions.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPageOption1, Me.TabPageOption2, Me.TabPageOption3, Me.TabPageOption4})
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
        'TabPageOption1
        '
        Me.TabPageOption1.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBoxAllowSearch, Me.GroupBox1, Me.GroupBoxOrder, Me.GroupBoxRestriction})
        Me.TabPageOption1.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption1.Name = "TabPageOption1"
        Me.TabPageOption1.Size = New System.Drawing.Size(529, 344)
        Me.TabPageOption1.TabIndex = 0
        Me.TabPageOption1.Text = "General Options"
        '
        'CheckBoxAllowSearch
        '
        Me.CheckBoxAllowSearch.Checked = True
        Me.CheckBoxAllowSearch.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxAllowSearch.Location = New System.Drawing.Point(24, 184)
        Me.CheckBoxAllowSearch.Name = "CheckBoxAllowSearch"
        Me.CheckBoxAllowSearch.Size = New System.Drawing.Size(232, 24)
        Me.CheckBoxAllowSearch.TabIndex = 7
        Me.CheckBoxAllowSearch.Text = "Provide Search Function"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.RadioButton5, Me.RadioButton4})
        Me.GroupBox1.Location = New System.Drawing.Point(328, 168)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(184, 80)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Order Sequence"
        '
        'RadioButton5
        '
        Me.RadioButton5.Checked = True
        Me.RadioButton5.Location = New System.Drawing.Point(88, 28)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(72, 24)
        Me.RadioButton5.TabIndex = 3
        Me.RadioButton5.TabStop = True
        Me.RadioButton5.Text = "DESC"
        '
        'RadioButton4
        '
        Me.RadioButton4.Location = New System.Drawing.Point(16, 28)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(60, 24)
        Me.RadioButton4.TabIndex = 2
        Me.RadioButton4.Text = "ASC"
        '
        'GroupBoxOrder
        '
        Me.GroupBoxOrder.Controls.AddRange(New System.Windows.Forms.Control() {Me.RadioButton3, Me.RadioButton2, Me.RadioButton1})
        Me.GroupBoxOrder.Location = New System.Drawing.Point(328, 16)
        Me.GroupBoxOrder.Name = "GroupBoxOrder"
        Me.GroupBoxOrder.Size = New System.Drawing.Size(184, 144)
        Me.GroupBoxOrder.TabIndex = 1
        Me.GroupBoxOrder.TabStop = False
        Me.GroupBoxOrder.Text = "Default Order By"
        '
        'RadioButton3
        '
        Me.RadioButton3.Location = New System.Drawing.Point(16, 88)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(152, 24)
        Me.RadioButton3.TabIndex = 2
        Me.RadioButton3.Text = "File Title/Description"
        '
        'RadioButton2
        '
        Me.RadioButton2.Location = New System.Drawing.Point(16, 56)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(80, 24)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "Uploader"
        '
        'RadioButton1
        '
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(16, 24)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Upload time"
        '
        'GroupBoxRestriction
        '
        Me.GroupBoxRestriction.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBoxMultiUpload, Me.TextBoxFileType, Me.CheckBoxFileType, Me.Label1, Me.TextBoxFileSize, Me.CheckBoxFileSize})
        Me.GroupBoxRestriction.Location = New System.Drawing.Point(16, 16)
        Me.GroupBoxRestriction.Name = "GroupBoxRestriction"
        Me.GroupBoxRestriction.Size = New System.Drawing.Size(296, 144)
        Me.GroupBoxRestriction.TabIndex = 0
        Me.GroupBoxRestriction.TabStop = False
        Me.GroupBoxRestriction.Text = "Upload Restrictions"
        '
        'CheckBoxMultiUpload
        '
        Me.CheckBoxMultiUpload.Location = New System.Drawing.Point(24, 112)
        Me.CheckBoxMultiUpload.Name = "CheckBoxMultiUpload"
        Me.CheckBoxMultiUpload.Size = New System.Drawing.Size(176, 24)
        Me.CheckBoxMultiUpload.TabIndex = 5
        Me.CheckBoxMultiUpload.Text = "Allow Multiple Upload"
        '
        'TextBoxFileType
        '
        Me.TextBoxFileType.Enabled = False
        Me.TextBoxFileType.Location = New System.Drawing.Point(56, 80)
        Me.TextBoxFileType.Name = "TextBoxFileType"
        Me.TextBoxFileType.Size = New System.Drawing.Size(216, 22)
        Me.TextBoxFileType.TabIndex = 4
        Me.TextBoxFileType.Text = "*.exe|*.com|*.bat"
        '
        'CheckBoxFileType
        '
        Me.CheckBoxFileType.Location = New System.Drawing.Point(24, 56)
        Me.CheckBoxFileType.Name = "CheckBoxFileType"
        Me.CheckBoxFileType.Size = New System.Drawing.Size(176, 24)
        Me.CheckBoxFileType.TabIndex = 3
        Me.CheckBoxFileType.Text = "Forbidden File Types:"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(240, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(24, 23)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "KB"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBoxFileSize
        '
        Me.TextBoxFileSize.Enabled = False
        Me.TextBoxFileSize.Location = New System.Drawing.Point(152, 24)
        Me.TextBoxFileSize.Name = "TextBoxFileSize"
        Me.TextBoxFileSize.Size = New System.Drawing.Size(80, 22)
        Me.TextBoxFileSize.TabIndex = 1
        Me.TextBoxFileSize.Text = ""
        '
        'CheckBoxFileSize
        '
        Me.CheckBoxFileSize.Location = New System.Drawing.Point(24, 24)
        Me.CheckBoxFileSize.Name = "CheckBoxFileSize"
        Me.CheckBoxFileSize.Size = New System.Drawing.Size(120, 24)
        Me.CheckBoxFileSize.TabIndex = 0
        Me.CheckBoxFileSize.Text = "Limit file size to:"
        '
        'TabPageOption2
        '
        Me.TabPageOption2.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox3, Me.GroupBox2})
        Me.TabPageOption2.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption2.Name = "TabPageOption2"
        Me.TabPageOption2.Size = New System.Drawing.Size(532, 344)
        Me.TabPageOption2.TabIndex = 1
        Me.TabPageOption2.Text = "User Options"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.RadioButtonNonMemAccessable, Me.RadioButton8, Me.CheckBoxNonMemView})
        Me.GroupBox3.Location = New System.Drawing.Point(32, 120)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(272, 112)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Non-member Permissions"
        '
        'RadioButtonNonMemAccessable
        '
        Me.RadioButtonNonMemAccessable.Location = New System.Drawing.Point(20, 48)
        Me.RadioButtonNonMemAccessable.Name = "RadioButtonNonMemAccessable"
        Me.RadioButtonNonMemAccessable.Size = New System.Drawing.Size(232, 24)
        Me.RadioButtonNonMemAccessable.TabIndex = 7
        Me.RadioButtonNonMemAccessable.Text = "Accessable to non-member"
        '
        'RadioButton8
        '
        Me.RadioButton8.Checked = True
        Me.RadioButton8.Location = New System.Drawing.Point(20, 24)
        Me.RadioButton8.Name = "RadioButton8"
        Me.RadioButton8.Size = New System.Drawing.Size(232, 24)
        Me.RadioButton8.TabIndex = 6
        Me.RadioButton8.TabStop = True
        Me.RadioButton8.Text = "Not accessable to non-member"
        '
        'CheckBoxNonMemView
        '
        Me.CheckBoxNonMemView.Enabled = False
        Me.CheckBoxNonMemView.Location = New System.Drawing.Point(68, 72)
        Me.CheckBoxNonMemView.Name = "CheckBoxNonMemView"
        Me.CheckBoxNonMemView.Size = New System.Drawing.Size(116, 24)
        Me.CheckBoxNonMemView.TabIndex = 4
        Me.CheckBoxNonMemView.Text = "Upload"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBox3, Me.CheckBox2})
        Me.GroupBox2.Location = New System.Drawing.Point(32, 24)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(272, 80)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Member Permissions"
        '
        'CheckBox3
        '
        Me.CheckBox3.Location = New System.Drawing.Point(128, 32)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.TabIndex = 1
        Me.CheckBox3.Text = "Delete "
        '
        'CheckBox2
        '
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Enabled = False
        Me.CheckBox2.Location = New System.Drawing.Point(24, 32)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.TabIndex = 0
        Me.CheckBox2.Text = "Upload"
        '
        'TabPageOption3
        '
        Me.TabPageOption3.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnUseDefaultFileNames, Me.LabelSetFileNames})
        Me.TabPageOption3.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption3.Name = "TabPageOption3"
        Me.TabPageOption3.Size = New System.Drawing.Size(532, 344)
        Me.TabPageOption3.TabIndex = 2
        Me.TabPageOption3.Text = "Rename Files"
        '
        'btnUseDefaultFileNames
        '
        Me.btnUseDefaultFileNames.Location = New System.Drawing.Point(32, 104)
        Me.btnUseDefaultFileNames.Name = "btnUseDefaultFileNames"
        Me.btnUseDefaultFileNames.Size = New System.Drawing.Size(160, 23)
        Me.btnUseDefaultFileNames.TabIndex = 6
        Me.btnUseDefaultFileNames.Text = "Use default names"
        '
        'LabelSetFileNames
        '
        Me.LabelSetFileNames.Location = New System.Drawing.Point(24, 24)
        Me.LabelSetFileNames.Name = "LabelSetFileNames"
        Me.LabelSetFileNames.Size = New System.Drawing.Size(464, 64)
        Me.LabelSetFileNames.TabIndex = 5
        Me.LabelSetFileNames.Text = "All the files to be generated are shown in the TreeView list above. If you want t" & _
        "o change the name of a file or folder, double click on a node or right click on " & _
        "a node and choose ""Rename"" from the menu."
        '
        'TabPageOption4
        '
        Me.TabPageOption4.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnOpenOutputFolder, Me.LabelOutputPath, Me.LabelAfterBuild, Me.Label6})
        Me.TabPageOption4.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption4.Name = "TabPageOption4"
        Me.TabPageOption4.Size = New System.Drawing.Size(532, 344)
        Me.TabPageOption4.TabIndex = 3
        Me.TabPageOption4.Text = "Create Files"
        '
        'btnOpenOutputFolder
        '
        Me.btnOpenOutputFolder.Location = New System.Drawing.Point(24, 192)
        Me.btnOpenOutputFolder.Name = "btnOpenOutputFolder"
        Me.btnOpenOutputFolder.Size = New System.Drawing.Size(176, 23)
        Me.btnOpenOutputFolder.TabIndex = 14
        Me.btnOpenOutputFolder.Text = "Open Output Folder..."
        '
        'LabelOutputPath
        '
        Me.LabelOutputPath.Location = New System.Drawing.Point(24, 104)
        Me.LabelOutputPath.Name = "LabelOutputPath"
        Me.LabelOutputPath.Size = New System.Drawing.Size(472, 80)
        Me.LabelOutputPath.TabIndex = 13
        Me.LabelOutputPath.Text = "..."
        '
        'LabelAfterBuild
        '
        Me.LabelAfterBuild.Location = New System.Drawing.Point(24, 56)
        Me.LabelAfterBuild.Name = "LabelAfterBuild"
        Me.LabelAfterBuild.Size = New System.Drawing.Size(480, 32)
        Me.LabelAfterBuild.TabIndex = 12
        Me.LabelAfterBuild.Text = "You can copy the output module to other places you want. The output path of the m" & _
        "odule is:"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(480, 32)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Use the ""Build"" menu or toolbar triangle button to generate files."
        '
        'LabelStep4
        '
        Me.LabelStep4.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep4.Location = New System.Drawing.Point(24, 290)
        Me.LabelStep4.Name = "LabelStep4"
        Me.LabelStep4.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep4.TabIndex = 43
        Me.LabelStep4.Text = "4. Create Files"
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
        Me.LabelStep3.TabIndex = 42
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
        Me.LabelStep2.TabIndex = 41
        Me.LabelStep2.Text = "2. User Options"
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
        Me.LabelStep1.TabIndex = 40
        Me.LabelStep1.Text = "1. General Options"
        Me.LabelStep1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FormGenUpload
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 559)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDisplay, Me.SplitterInstr, Me.PanelInstr, Me.SplitterStep, Me.PanelStep})
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FormGenUpload"
        Me.Text = "Upload"
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
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBoxOrder.ResumeLayout(False)
        Me.GroupBoxRestriction.ResumeLayout(False)
        Me.TabPageOption2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.TabPageOption3.ResumeLayout(False)
        Me.TabPageOption4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FormGenUpload_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.moduleType = ModuleMain.xcModuleType.xcUpload
        Me.NUM_STEPS = 4
        Me.initModuleContent()
    End Sub


    Private Sub FormGenUpload_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize

    End Sub



    Protected Overrides Sub initTreeNodesText()

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
                End If
            Next
        End If

        constructTreeViewFileList()
        Me.processStep()

    End Sub


    Protected Overrides Sub constructTreeViewFileList()

    End Sub


    Protected Friend Overrides Sub saveModule()
        Dim str As String = ""
        str += "CreateDate=" & Now() & vbCrLf
        str += "ModuleName=" & Me.moduleName & vbCrLf
        str += "ModuleType=" & Me.moduleType & vbCrLf
        str += vbCrLf
        str += "StepNumber=" & Me.stepNumber & vbCrLf
        str += vbCrLf
        'str += "indexFileName=" & Me.indexFileNode.Text & vbCrLf
        str += vbCrLf
        ' interface status

        IOManager.SaveTextToFile(str, Me.moduleFullPath)

        Me.infoSaved = True
    End Sub


    Protected Friend Overrides Sub buildModule( _
    Optional ByVal frameworkParams As clsWebsiteFrameworkParameters = Nothing)

    End Sub


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

    Private Sub CheckBoxFileType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxFileType.CheckedChanged
        If Me.CheckBoxFileType.Checked = True Then
            Me.TextBoxFileType.Enabled = True
        Else
            Me.TextBoxFileType.Enabled = False
        End If
    End Sub

    Private Sub TabControlOptions_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControlOptions.SelectedIndexChanged
        Select Case Me.TabControlOptions.SelectedIndex
            Case 0
                Me.stepNumber = 1
                Me.TabControlDisplay.SelectedTab = Me.TabPageSource
                Me.LabelStepInstr.Text = "General options of the Upload module."
            Case 1
                Me.stepNumber = 2
                Me.TabControlDisplay.SelectedTab = Me.TabPageFileList
                Me.LabelStepInstr.Text = "Options about what each type of user can do."
            Case 2
                Me.stepNumber = 3
                Me.LabelStepInstr.Text = "You don't really need to change filenames for the module to work. But in case you prefer other names, you can do it."
            Case 3
                Me.stepNumber = 4
                Me.LabelStepInstr.Text = "The shortcut to build project is F5. Only opening modules in a project are built. The shortcut to build only the currently active module is F6."
        End Select
        Me.processStep()
        prevSelectedIndex = Me.TabControlOptions.SelectedIndex
    End Sub

    Private Sub CheckBoxFileSize_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxFileSize.CheckedChanged
        If Me.CheckBoxFileSize.Checked = True Then
            Me.TextBoxFileSize.Enabled = True
        Else
            Me.TextBoxFileSize.Enabled = False
        End If
    End Sub
End Class
