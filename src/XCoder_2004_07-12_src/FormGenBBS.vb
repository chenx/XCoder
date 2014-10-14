

Public Class FormGenBBS
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton6 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton7 As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents TabPageOption2 As System.Windows.Forms.TabPage
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox7 As System.Windows.Forms.CheckBox
    Friend WithEvents RadioButton8 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonNonMemAccessable As System.Windows.Forms.RadioButton
    Friend WithEvents CheckBoxNonMemPost As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxNonMemView As System.Windows.Forms.CheckBox
    Friend WithEvents TabPageOption3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption4 As System.Windows.Forms.TabPage
    Friend WithEvents LabelStep4 As System.Windows.Forms.Label
    Friend WithEvents LabelStep3 As System.Windows.Forms.Label
    Friend WithEvents LabelStep2 As System.Windows.Forms.Label
    Friend WithEvents LabelStep1 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxAllowSearch As System.Windows.Forms.CheckBox
    Friend WithEvents TextBoxMsgPerPage As System.Windows.Forms.TextBox
    Friend WithEvents btnUseDefaultFileNames As System.Windows.Forms.Button
    Friend WithEvents LabelSetFileNames As System.Windows.Forms.Label
    Friend WithEvents btnOpenOutputFolder As System.Windows.Forms.Button
    Friend WithEvents LabelOutputPath As System.Windows.Forms.Label
    Friend WithEvents LabelAfterBuild As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents RadioButtonDESC As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonASC As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonThreadOrTime As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonTime As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonThread As System.Windows.Forms.RadioButton
    Friend WithEvents ComboBoxBBSStyle As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormGenBBS))
        Me.TabPageOption1 = New System.Windows.Forms.TabPage()
        Me.CheckBoxAllowSearch = New System.Windows.Forms.CheckBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.ComboBoxBBSStyle = New System.Windows.Forms.ComboBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.RadioButton6 = New System.Windows.Forms.RadioButton()
        Me.RadioButton7 = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.TextBoxMsgPerPage = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonDESC = New System.Windows.Forms.RadioButton()
        Me.RadioButtonASC = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonThreadOrTime = New System.Windows.Forms.RadioButton()
        Me.RadioButtonTime = New System.Windows.Forms.RadioButton()
        Me.RadioButtonThread = New System.Windows.Forms.RadioButton()
        Me.TabPageOption2 = New System.Windows.Forms.TabPage()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonNonMemAccessable = New System.Windows.Forms.RadioButton()
        Me.RadioButton8 = New System.Windows.Forms.RadioButton()
        Me.CheckBoxNonMemPost = New System.Windows.Forms.CheckBox()
        Me.CheckBoxNonMemView = New System.Windows.Forms.CheckBox()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.CheckBox7 = New System.Windows.Forms.CheckBox()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.CheckBox6 = New System.Windows.Forms.CheckBox()
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
        Me.PanelInstr.SuspendLayout()
        Me.PanelStep.SuspendLayout()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControlDisplay.SuspendLayout()
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
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPageOption2.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.TabPageOption3.SuspendLayout()
        Me.TabPageOption4.SuspendLayout()
        Me.SuspendLayout()
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Size = New System.Drawing.Size(529, 74)
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
        Me.LabelDispTitle.BackColor = System.Drawing.Color.Blue
        Me.LabelDispTitle.ForeColor = System.Drawing.Color.White
        Me.LabelDispTitle.Visible = True
        '
        'PanelDispTitle
        '
        Me.PanelDispTitle.Visible = True
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
        Me.TabPageSource.Size = New System.Drawing.Size(529, 74)
        '
        'TabPageQuickView
        '
        Me.TabPageQuickView.Size = New System.Drawing.Size(529, 100)
        '
        'SplitterInstrDetail
        '
        Me.SplitterInstrDetail.Visible = True
        '
        'LabelInstr
        '
        Me.LabelInstr.Visible = True
        '
        'GroupBoxInstr
        '
        Me.GroupBoxInstr.Visible = True
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
        'TabControlOptions
        '
        Me.TabControlOptions.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPageOption1, Me.TabPageOption2, Me.TabPageOption3, Me.TabPageOption4})
        Me.TabControlOptions.ItemSize = New System.Drawing.Size(0, 21)
        Me.TabControlOptions.Visible = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(125, 457)
        Me.btnNext.Visible = True
        '
        'GroupBoxStep
        '
        Me.GroupBoxStep.Visible = True
        '
        'LabelStepInstr
        '
        Me.LabelStepInstr.Visible = True
        '
        'TabPageOption1
        '
        Me.TabPageOption1.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBoxAllowSearch, Me.GroupBox5, Me.GroupBox4, Me.GroupBox3, Me.GroupBox2, Me.GroupBox1})
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
        Me.CheckBoxAllowSearch.Location = New System.Drawing.Point(272, 216)
        Me.CheckBoxAllowSearch.Name = "CheckBoxAllowSearch"
        Me.CheckBoxAllowSearch.Size = New System.Drawing.Size(232, 24)
        Me.CheckBoxAllowSearch.TabIndex = 6
        Me.CheckBoxAllowSearch.Text = "Provide Search Function"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.AddRange(New System.Windows.Forms.Control() {Me.ComboBoxBBSStyle})
        Me.GroupBox5.Location = New System.Drawing.Point(264, 96)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(240, 96)
        Me.GroupBox5.TabIndex = 5
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Choose style"
        '
        'ComboBoxBBSStyle
        '
        Me.ComboBoxBBSStyle.Items.AddRange(New Object() {"Style 1 (MIT BBS)", "Style 2 (Wenxuechen)"})
        Me.ComboBoxBBSStyle.Location = New System.Drawing.Point(24, 32)
        Me.ComboBoxBBSStyle.Name = "ComboBoxBBSStyle"
        Me.ComboBoxBBSStyle.Size = New System.Drawing.Size(200, 24)
        Me.ComboBoxBBSStyle.TabIndex = 0
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.AddRange(New System.Windows.Forms.Control() {Me.RadioButton6, Me.RadioButton7})
        Me.GroupBox4.Location = New System.Drawing.Point(264, 16)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(240, 64)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "At the beginning go to"
        '
        'RadioButton6
        '
        Me.RadioButton6.Checked = True
        Me.RadioButton6.Location = New System.Drawing.Point(120, 24)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.TabIndex = 3
        Me.RadioButton6.TabStop = True
        Me.RadioButton6.Text = "Last Page"
        '
        'RadioButton7
        '
        Me.RadioButton7.Location = New System.Drawing.Point(16, 24)
        Me.RadioButton7.Name = "RadioButton7"
        Me.RadioButton7.Size = New System.Drawing.Size(96, 24)
        Me.RadioButton7.TabIndex = 2
        Me.RadioButton7.Text = "First Page"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBoxMsgPerPage})
        Me.GroupBox3.Location = New System.Drawing.Point(16, 200)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(240, 56)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Number of articles per page"
        '
        'TextBoxMsgPerPage
        '
        Me.TextBoxMsgPerPage.Location = New System.Drawing.Point(16, 24)
        Me.TextBoxMsgPerPage.Name = "TextBoxMsgPerPage"
        Me.TextBoxMsgPerPage.Size = New System.Drawing.Size(184, 22)
        Me.TextBoxMsgPerPage.TabIndex = 0
        Me.TextBoxMsgPerPage.Text = "20"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.RadioButtonDESC, Me.RadioButtonASC})
        Me.GroupBox2.Location = New System.Drawing.Point(16, 136)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(240, 56)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Order articles by time"
        '
        'RadioButtonDESC
        '
        Me.RadioButtonDESC.Location = New System.Drawing.Point(104, 24)
        Me.RadioButtonDESC.Name = "RadioButtonDESC"
        Me.RadioButtonDESC.Size = New System.Drawing.Size(72, 24)
        Me.RadioButtonDESC.TabIndex = 1
        Me.RadioButtonDESC.Text = "DESC"
        '
        'RadioButtonASC
        '
        Me.RadioButtonASC.Checked = True
        Me.RadioButtonASC.Location = New System.Drawing.Point(16, 24)
        Me.RadioButtonASC.Name = "RadioButtonASC"
        Me.RadioButtonASC.Size = New System.Drawing.Size(72, 24)
        Me.RadioButtonASC.TabIndex = 0
        Me.RadioButtonASC.TabStop = True
        Me.RadioButtonASC.Text = "ASC"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.RadioButtonThreadOrTime, Me.RadioButtonTime, Me.RadioButtonThread})
        Me.GroupBox1.Location = New System.Drawing.Point(16, 16)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(240, 112)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "List articles by"
        '
        'RadioButtonThreadOrTime
        '
        Me.RadioButtonThreadOrTime.Location = New System.Drawing.Point(16, 72)
        Me.RadioButtonThreadOrTime.Name = "RadioButtonThreadOrTime"
        Me.RadioButtonThreadOrTime.Size = New System.Drawing.Size(176, 24)
        Me.RadioButtonThreadOrTime.TabIndex = 2
        Me.RadioButtonThreadOrTime.Text = "Thread (default) or Time"
        '
        'RadioButtonTime
        '
        Me.RadioButtonTime.Location = New System.Drawing.Point(16, 48)
        Me.RadioButtonTime.Name = "RadioButtonTime"
        Me.RadioButtonTime.Size = New System.Drawing.Size(112, 24)
        Me.RadioButtonTime.TabIndex = 1
        Me.RadioButtonTime.Text = "Time"
        '
        'RadioButtonThread
        '
        Me.RadioButtonThread.Checked = True
        Me.RadioButtonThread.Location = New System.Drawing.Point(16, 24)
        Me.RadioButtonThread.Name = "RadioButtonThread"
        Me.RadioButtonThread.Size = New System.Drawing.Size(120, 24)
        Me.RadioButtonThread.TabIndex = 0
        Me.RadioButtonThread.TabStop = True
        Me.RadioButtonThread.Text = "Thread"
        '
        'TabPageOption2
        '
        Me.TabPageOption2.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox7, Me.GroupBox6})
        Me.TabPageOption2.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption2.Name = "TabPageOption2"
        Me.TabPageOption2.Size = New System.Drawing.Size(529, 344)
        Me.TabPageOption2.TabIndex = 1
        Me.TabPageOption2.Text = "User Options"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.AddRange(New System.Windows.Forms.Control() {Me.RadioButtonNonMemAccessable, Me.RadioButton8, Me.CheckBoxNonMemPost, Me.CheckBoxNonMemView})
        Me.GroupBox7.Location = New System.Drawing.Point(24, 144)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(280, 112)
        Me.GroupBox7.TabIndex = 7
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Non-member user permissions"
        '
        'RadioButtonNonMemAccessable
        '
        Me.RadioButtonNonMemAccessable.Location = New System.Drawing.Point(32, 48)
        Me.RadioButtonNonMemAccessable.Name = "RadioButtonNonMemAccessable"
        Me.RadioButtonNonMemAccessable.Size = New System.Drawing.Size(232, 24)
        Me.RadioButtonNonMemAccessable.TabIndex = 3
        Me.RadioButtonNonMemAccessable.Text = "Accessable to non-member"
        '
        'RadioButton8
        '
        Me.RadioButton8.Checked = True
        Me.RadioButton8.Location = New System.Drawing.Point(32, 24)
        Me.RadioButton8.Name = "RadioButton8"
        Me.RadioButton8.Size = New System.Drawing.Size(232, 24)
        Me.RadioButton8.TabIndex = 2
        Me.RadioButton8.TabStop = True
        Me.RadioButton8.Text = "Not accessable to non-member"
        '
        'CheckBoxNonMemPost
        '
        Me.CheckBoxNonMemPost.Enabled = False
        Me.CheckBoxNonMemPost.Location = New System.Drawing.Point(144, 72)
        Me.CheckBoxNonMemPost.Name = "CheckBoxNonMemPost"
        Me.CheckBoxNonMemPost.Size = New System.Drawing.Size(72, 24)
        Me.CheckBoxNonMemPost.TabIndex = 1
        Me.CheckBoxNonMemPost.Text = "Post"
        '
        'CheckBoxNonMemView
        '
        Me.CheckBoxNonMemView.Enabled = False
        Me.CheckBoxNonMemView.Location = New System.Drawing.Point(80, 72)
        Me.CheckBoxNonMemView.Name = "CheckBoxNonMemView"
        Me.CheckBoxNonMemView.Size = New System.Drawing.Size(64, 24)
        Me.CheckBoxNonMemView.TabIndex = 0
        Me.CheckBoxNonMemView.Text = "View"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBox7, Me.CheckBox5, Me.CheckBox6, Me.CheckBox2})
        Me.GroupBox6.Location = New System.Drawing.Point(24, 24)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(280, 96)
        Me.GroupBox6.TabIndex = 6
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Member permissions"
        '
        'CheckBox7
        '
        Me.CheckBox7.Location = New System.Drawing.Point(160, 24)
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.TabIndex = 8
        Me.CheckBox7.Text = "Delete"
        '
        'CheckBox5
        '
        Me.CheckBox5.Checked = True
        Me.CheckBox5.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox5.Enabled = False
        Me.CheckBox5.Location = New System.Drawing.Point(88, 24)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(64, 24)
        Me.CheckBox5.TabIndex = 7
        Me.CheckBox5.Text = "Post"
        '
        'CheckBox6
        '
        Me.CheckBox6.Checked = True
        Me.CheckBox6.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox6.Enabled = False
        Me.CheckBox6.Location = New System.Drawing.Point(16, 24)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(56, 24)
        Me.CheckBox6.TabIndex = 6
        Me.CheckBox6.Text = "View"
        '
        'CheckBox2
        '
        Me.CheckBox2.Location = New System.Drawing.Point(16, 56)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(248, 24)
        Me.CheckBox2.TabIndex = 5
        Me.CheckBox2.Text = "Choose View preference"
        '
        'TabPageOption3
        '
        Me.TabPageOption3.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnUseDefaultFileNames, Me.LabelSetFileNames})
        Me.TabPageOption3.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption3.Name = "TabPageOption3"
        Me.TabPageOption3.Size = New System.Drawing.Size(529, 344)
        Me.TabPageOption3.TabIndex = 2
        Me.TabPageOption3.Text = "Rename files"
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
        Me.TabPageOption4.Size = New System.Drawing.Size(529, 344)
        Me.TabPageOption4.TabIndex = 3
        Me.TabPageOption4.Text = "Create Files"
        '
        'btnOpenOutputFolder
        '
        Me.btnOpenOutputFolder.Location = New System.Drawing.Point(16, 192)
        Me.btnOpenOutputFolder.Name = "btnOpenOutputFolder"
        Me.btnOpenOutputFolder.Size = New System.Drawing.Size(176, 23)
        Me.btnOpenOutputFolder.TabIndex = 14
        Me.btnOpenOutputFolder.Text = "Open Output Folder..."
        '
        'LabelOutputPath
        '
        Me.LabelOutputPath.Location = New System.Drawing.Point(16, 104)
        Me.LabelOutputPath.Name = "LabelOutputPath"
        Me.LabelOutputPath.Size = New System.Drawing.Size(472, 80)
        Me.LabelOutputPath.TabIndex = 13
        Me.LabelOutputPath.Text = "..."
        '
        'LabelAfterBuild
        '
        Me.LabelAfterBuild.Location = New System.Drawing.Point(16, 56)
        Me.LabelAfterBuild.Name = "LabelAfterBuild"
        Me.LabelAfterBuild.Size = New System.Drawing.Size(480, 32)
        Me.LabelAfterBuild.TabIndex = 12
        Me.LabelAfterBuild.Text = "You can copy the output module to other places you want. The output path of the m" & _
        "odule is:"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 24)
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
        Me.LabelStep4.TabIndex = 39
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
        Me.LabelStep3.TabIndex = 38
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
        Me.LabelStep2.TabIndex = 37
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
        Me.LabelStep1.TabIndex = 36
        Me.LabelStep1.Text = "1. General Options"
        Me.LabelStep1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FormGenBBS
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 559)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDisplay, Me.SplitterInstr, Me.PanelInstr, Me.SplitterStep, Me.PanelStep})
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FormGenBBS"
        Me.Text = "Discussion Board"
        Me.PanelInstr.ResumeLayout(False)
        Me.PanelStep.ResumeLayout(False)
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControlDisplay.ResumeLayout(False)
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
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.TabPageOption2.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.TabPageOption3.ResumeLayout(False)
        Me.TabPageOption4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub FormGenBBS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.moduleType = ModuleMain.xcModuleType.xcDiscBoard
        Me.NUM_STEPS = 4
        Me.ComboBoxBBSStyle.SelectedIndex = 0
        Me.initModuleContent()
    End Sub


    Private Sub FormGenBBS_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize

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

    Private Sub RadioButtonNonMemAccessable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonNonMemAccessable.CheckedChanged
        If Me.RadioButtonNonMemAccessable.Checked = True Then
            Me.CheckBoxNonMemView.Enabled = True
            Me.CheckBoxNonMemPost.Enabled = True
        Else
            Me.CheckBoxNonMemView.Enabled = False
            Me.CheckBoxNonMemPost.Enabled = False
        End If
    End Sub

    Private Sub TabControlOptions_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControlOptions.SelectedIndexChanged
        Select Case Me.TabControlOptions.SelectedIndex
            Case 0
                Me.stepNumber = 1
                Me.TabControlDisplay.SelectedTab = Me.TabPageSource
                Me.LabelStepInstr.Text = "General options of the discussion board."
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
End Class
