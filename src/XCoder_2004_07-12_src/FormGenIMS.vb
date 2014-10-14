

Public Class FormGenIMS
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
    Friend WithEvents GroupBoxIMSType As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnUseDefaultFileNames As System.Windows.Forms.Button
    Friend WithEvents LabelSetFileNames As System.Windows.Forms.Label
    Friend WithEvents btnOpenOutputFolder As System.Windows.Forms.Button
    Friend WithEvents LabelOutputPath As System.Windows.Forms.Label
    Friend WithEvents LabelAfterBuild As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TabPageOption3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption5 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption6 As System.Windows.Forms.TabPage
    Friend WithEvents btnMetadata As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents LabelStep3 As System.Windows.Forms.Label
    Friend WithEvents LabelStep2 As System.Windows.Forms.Label
    Friend WithEvents LabelStep1 As System.Windows.Forms.Label
    Friend WithEvents LabelStep6 As System.Windows.Forms.Label
    Friend WithEvents LabelStep5 As System.Windows.Forms.Label
    Friend WithEvents LabelStep4 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxIMSType As System.Windows.Forms.ComboBox
    Friend WithEvents LabelIMSTypeDescription As System.Windows.Forms.Label
    Friend WithEvents CheckBoxAccessoryWord As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxAccessoryAccess As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxAccessoryCalendar As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxAccessorySearch As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxAdmDelete As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxAdmApprove As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxAdmEdit As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxAdmView As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxNonMemView As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxMemDisapprove As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxMemDelete As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxMemApprove As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxMemEdit As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxMemView As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxAdmDisapprove As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TabPageOption1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPageOption2 As System.Windows.Forms.TabPage
    Friend WithEvents TextBoxDBTable1Name As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxDBTable2Name As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormGenIMS))
        Me.TabPageOption1 = New System.Windows.Forms.TabPage()
        Me.GroupBoxIMSType = New System.Windows.Forms.GroupBox()
        Me.LabelIMSTypeDescription = New System.Windows.Forms.Label()
        Me.ComboBoxIMSType = New System.Windows.Forms.ComboBox()
        Me.TabPageOption3 = New System.Windows.Forms.TabPage()
        Me.CheckBoxAccessoryWord = New System.Windows.Forms.CheckBox()
        Me.CheckBoxAccessoryAccess = New System.Windows.Forms.CheckBox()
        Me.CheckBoxAccessoryCalendar = New System.Windows.Forms.CheckBox()
        Me.CheckBoxAccessorySearch = New System.Windows.Forms.CheckBox()
        Me.TabPageOption4 = New System.Windows.Forms.TabPage()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxAdmDisapprove = New System.Windows.Forms.CheckBox()
        Me.CheckBoxAdmDelete = New System.Windows.Forms.CheckBox()
        Me.CheckBoxAdmApprove = New System.Windows.Forms.CheckBox()
        Me.CheckBoxAdmEdit = New System.Windows.Forms.CheckBox()
        Me.CheckBoxAdmView = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxNonMemView = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxMemDisapprove = New System.Windows.Forms.CheckBox()
        Me.CheckBoxMemDelete = New System.Windows.Forms.CheckBox()
        Me.CheckBoxMemApprove = New System.Windows.Forms.CheckBox()
        Me.CheckBoxMemEdit = New System.Windows.Forms.CheckBox()
        Me.CheckBoxMemView = New System.Windows.Forms.CheckBox()
        Me.TabPageOption5 = New System.Windows.Forms.TabPage()
        Me.btnUseDefaultFileNames = New System.Windows.Forms.Button()
        Me.LabelSetFileNames = New System.Windows.Forms.Label()
        Me.TabPageOption6 = New System.Windows.Forms.TabPage()
        Me.btnOpenOutputFolder = New System.Windows.Forms.Button()
        Me.LabelOutputPath = New System.Windows.Forms.Label()
        Me.LabelAfterBuild = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TabPageOption2 = New System.Windows.Forms.TabPage()
        Me.TextBoxDBTable2Name = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBoxDBTable1Name = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnMetadata = New System.Windows.Forms.Button()
        Me.LabelStep3 = New System.Windows.Forms.Label()
        Me.LabelStep2 = New System.Windows.Forms.Label()
        Me.LabelStep1 = New System.Windows.Forms.Label()
        Me.LabelStep6 = New System.Windows.Forms.Label()
        Me.LabelStep5 = New System.Windows.Forms.Label()
        Me.LabelStep4 = New System.Windows.Forms.Label()
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
        Me.GroupBoxIMSType.SuspendLayout()
        Me.TabPageOption3.SuspendLayout()
        Me.TabPageOption4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPageOption5.SuspendLayout()
        Me.TabPageOption6.SuspendLayout()
        Me.TabPageOption2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControlOptions
        '
        Me.TabControlOptions.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPageOption1, Me.TabPageOption2, Me.TabPageOption3, Me.TabPageOption4, Me.TabPageOption5, Me.TabPageOption6})
        Me.TabControlOptions.ItemSize = New System.Drawing.Size(0, 21)
        Me.TabControlOptions.Size = New System.Drawing.Size(537, 299)
        Me.TabControlOptions.Visible = True
        '
        'GroupBoxInstr
        '
        Me.GroupBoxInstr.Location = New System.Drawing.Point(1, 328)
        Me.GroupBoxInstr.Visible = True
        '
        'LabelInstr
        '
        Me.LabelInstr.Visible = True
        '
        'SplitterInstrDetail
        '
        Me.SplitterInstrDetail.Location = New System.Drawing.Point(1, 325)
        Me.SplitterInstrDetail.Visible = True
        '
        'TabPageQuickView
        '
        Me.TabPageQuickView.Size = New System.Drawing.Size(529, 99)
        '
        'TabPageSource
        '
        Me.TabPageSource.Size = New System.Drawing.Size(0, 0)
        '
        'SplitterInstr
        '
        Me.SplitterInstr.Location = New System.Drawing.Point(253, 155)
        Me.SplitterInstr.Size = New System.Drawing.Size(539, 4)
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
        Me.PanelStepDetail.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelStep6, Me.LabelStep5, Me.LabelStep4, Me.LabelStep3, Me.LabelStep2, Me.LabelStep1})
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
        Me.PanelDisplay.Size = New System.Drawing.Size(539, 155)
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
        Me.TabControlDisplay.Size = New System.Drawing.Size(537, 128)
        Me.TabControlDisplay.Visible = True
        '
        'AxWebBrowser1
        '
        Me.AxWebBrowser1.Size = New System.Drawing.Size(529, 99)
        Me.AxWebBrowser1.Visible = True
        '
        'btnBack
        '
        Me.btnBack.Visible = True
        '
        'btnNext
        '
        Me.btnNext.Visible = True
        '
        'PanelStep
        '
        Me.PanelStep.Visible = True
        '
        'SplitterStep
        '
        Me.SplitterStep.Size = New System.Drawing.Size(3, 560)
        Me.SplitterStep.Visible = True
        '
        'PanelInstr
        '
        Me.PanelInstr.Location = New System.Drawing.Point(253, 159)
        Me.PanelInstr.Size = New System.Drawing.Size(539, 401)
        Me.PanelInstr.Visible = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Size = New System.Drawing.Size(0, 0)
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
        Me.TabPageFileList.Size = New System.Drawing.Size(1017, 442)
        '
        'TreeView1
        '
        Me.TreeView1.Size = New System.Drawing.Size(1017, 442)
        Me.TreeView1.Visible = True
        '
        'TabPageOption1
        '
        Me.TabPageOption1.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBoxIMSType})
        Me.TabPageOption1.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption1.Name = "TabPageOption1"
        Me.TabPageOption1.Size = New System.Drawing.Size(529, 270)
        Me.TabPageOption1.TabIndex = 0
        Me.TabPageOption1.Text = "IMS Type"
        '
        'GroupBoxIMSType
        '
        Me.GroupBoxIMSType.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelIMSTypeDescription, Me.ComboBoxIMSType})
        Me.GroupBoxIMSType.Location = New System.Drawing.Point(16, 24)
        Me.GroupBoxIMSType.Name = "GroupBoxIMSType"
        Me.GroupBoxIMSType.Size = New System.Drawing.Size(488, 144)
        Me.GroupBoxIMSType.TabIndex = 0
        Me.GroupBoxIMSType.TabStop = False
        Me.GroupBoxIMSType.Text = "Choose Information Management Module type"
        '
        'LabelIMSTypeDescription
        '
        Me.LabelIMSTypeDescription.Location = New System.Drawing.Point(24, 32)
        Me.LabelIMSTypeDescription.Name = "LabelIMSTypeDescription"
        Me.LabelIMSTypeDescription.Size = New System.Drawing.Size(440, 48)
        Me.LabelIMSTypeDescription.TabIndex = 4
        '
        'ComboBoxIMSType
        '
        Me.ComboBoxIMSType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxIMSType.Items.AddRange(New Object() {"Type 0 - Simple IMS", "Type 1 - Member Self-manage IMS", "Type 2 - Member And Admin manage IMS"})
        Me.ComboBoxIMSType.Location = New System.Drawing.Point(24, 96)
        Me.ComboBoxIMSType.Name = "ComboBoxIMSType"
        Me.ComboBoxIMSType.Size = New System.Drawing.Size(440, 24)
        Me.ComboBoxIMSType.TabIndex = 3
        '
        'TabPageOption3
        '
        Me.TabPageOption3.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBoxAccessoryWord, Me.CheckBoxAccessoryAccess, Me.CheckBoxAccessoryCalendar, Me.CheckBoxAccessorySearch})
        Me.TabPageOption3.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption3.Name = "TabPageOption3"
        Me.TabPageOption3.Size = New System.Drawing.Size(1020, 269)
        Me.TabPageOption3.TabIndex = 1
        Me.TabPageOption3.Text = "Accessories"
        '
        'CheckBoxAccessoryWord
        '
        Me.CheckBoxAccessoryWord.Location = New System.Drawing.Point(24, 144)
        Me.CheckBoxAccessoryWord.Name = "CheckBoxAccessoryWord"
        Me.CheckBoxAccessoryWord.Size = New System.Drawing.Size(456, 24)
        Me.CheckBoxAccessoryWord.TabIndex = 3
        Me.CheckBoxAccessoryWord.Text = "Add Report as MS Word file function"
        '
        'CheckBoxAccessoryAccess
        '
        Me.CheckBoxAccessoryAccess.Location = New System.Drawing.Point(24, 104)
        Me.CheckBoxAccessoryAccess.Name = "CheckBoxAccessoryAccess"
        Me.CheckBoxAccessoryAccess.Size = New System.Drawing.Size(456, 24)
        Me.CheckBoxAccessoryAccess.TabIndex = 2
        Me.CheckBoxAccessoryAccess.Text = "Add Export to MS Access/Excel function"
        '
        'CheckBoxAccessoryCalendar
        '
        Me.CheckBoxAccessoryCalendar.Location = New System.Drawing.Point(24, 56)
        Me.CheckBoxAccessoryCalendar.Name = "CheckBoxAccessoryCalendar"
        Me.CheckBoxAccessoryCalendar.Size = New System.Drawing.Size(456, 40)
        Me.CheckBoxAccessoryCalendar.TabIndex = 1
        Me.CheckBoxAccessoryCalendar.Text = "Add Calendar function (for this your module must have date fields)"
        '
        'CheckBoxAccessorySearch
        '
        Me.CheckBoxAccessorySearch.Location = New System.Drawing.Point(24, 24)
        Me.CheckBoxAccessorySearch.Name = "CheckBoxAccessorySearch"
        Me.CheckBoxAccessorySearch.Size = New System.Drawing.Size(456, 24)
        Me.CheckBoxAccessorySearch.TabIndex = 0
        Me.CheckBoxAccessorySearch.Text = "Add Search function "
        '
        'TabPageOption4
        '
        Me.TabPageOption4.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox3, Me.GroupBox2, Me.GroupBox1})
        Me.TabPageOption4.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption4.Name = "TabPageOption4"
        Me.TabPageOption4.Size = New System.Drawing.Size(1020, 269)
        Me.TabPageOption4.TabIndex = 2
        Me.TabPageOption4.Text = "User Actions"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBoxAdmDisapprove, Me.CheckBoxAdmDelete, Me.CheckBoxAdmApprove, Me.CheckBoxAdmEdit, Me.CheckBoxAdmView})
        Me.GroupBox3.Location = New System.Drawing.Point(264, 96)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(232, 160)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Administrator functions"
        '
        'CheckBoxAdmDisapprove
        '
        Me.CheckBoxAdmDisapprove.Checked = True
        Me.CheckBoxAdmDisapprove.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxAdmDisapprove.Location = New System.Drawing.Point(40, 96)
        Me.CheckBoxAdmDisapprove.Name = "CheckBoxAdmDisapprove"
        Me.CheckBoxAdmDisapprove.Size = New System.Drawing.Size(168, 24)
        Me.CheckBoxAdmDisapprove.TabIndex = 4
        Me.CheckBoxAdmDisapprove.Text = "Disapprove"
        '
        'CheckBoxAdmDelete
        '
        Me.CheckBoxAdmDelete.Checked = True
        Me.CheckBoxAdmDelete.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxAdmDelete.Location = New System.Drawing.Point(40, 120)
        Me.CheckBoxAdmDelete.Name = "CheckBoxAdmDelete"
        Me.CheckBoxAdmDelete.Size = New System.Drawing.Size(136, 24)
        Me.CheckBoxAdmDelete.TabIndex = 3
        Me.CheckBoxAdmDelete.Text = "Delete"
        '
        'CheckBoxAdmApprove
        '
        Me.CheckBoxAdmApprove.Checked = True
        Me.CheckBoxAdmApprove.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxAdmApprove.Location = New System.Drawing.Point(40, 72)
        Me.CheckBoxAdmApprove.Name = "CheckBoxAdmApprove"
        Me.CheckBoxAdmApprove.Size = New System.Drawing.Size(168, 24)
        Me.CheckBoxAdmApprove.TabIndex = 2
        Me.CheckBoxAdmApprove.Text = "Approve"
        '
        'CheckBoxAdmEdit
        '
        Me.CheckBoxAdmEdit.Checked = True
        Me.CheckBoxAdmEdit.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxAdmEdit.Location = New System.Drawing.Point(40, 48)
        Me.CheckBoxAdmEdit.Name = "CheckBoxAdmEdit"
        Me.CheckBoxAdmEdit.Size = New System.Drawing.Size(120, 24)
        Me.CheckBoxAdmEdit.TabIndex = 1
        Me.CheckBoxAdmEdit.Text = "Edit"
        '
        'CheckBoxAdmView
        '
        Me.CheckBoxAdmView.Checked = True
        Me.CheckBoxAdmView.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxAdmView.Enabled = False
        Me.CheckBoxAdmView.Location = New System.Drawing.Point(40, 24)
        Me.CheckBoxAdmView.Name = "CheckBoxAdmView"
        Me.CheckBoxAdmView.Size = New System.Drawing.Size(128, 24)
        Me.CheckBoxAdmView.TabIndex = 0
        Me.CheckBoxAdmView.Text = "View"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBoxNonMemView})
        Me.GroupBox2.Location = New System.Drawing.Point(16, 16)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(480, 64)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Non-member user access permission"
        '
        'CheckBoxNonMemView
        '
        Me.CheckBoxNonMemView.Location = New System.Drawing.Point(40, 24)
        Me.CheckBoxNonMemView.Name = "CheckBoxNonMemView"
        Me.CheckBoxNonMemView.Size = New System.Drawing.Size(344, 24)
        Me.CheckBoxNonMemView.TabIndex = 0
        Me.CheckBoxNonMemView.Text = "View data in a public page"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBoxMemDisapprove, Me.CheckBoxMemDelete, Me.CheckBoxMemApprove, Me.CheckBoxMemEdit, Me.CheckBoxMemView})
        Me.GroupBox1.Location = New System.Drawing.Point(16, 96)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(232, 160)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Member access permission"
        '
        'CheckBoxMemDisapprove
        '
        Me.CheckBoxMemDisapprove.Location = New System.Drawing.Point(40, 96)
        Me.CheckBoxMemDisapprove.Name = "CheckBoxMemDisapprove"
        Me.CheckBoxMemDisapprove.Size = New System.Drawing.Size(128, 24)
        Me.CheckBoxMemDisapprove.TabIndex = 4
        Me.CheckBoxMemDisapprove.Text = "Disapprove"
        '
        'CheckBoxMemDelete
        '
        Me.CheckBoxMemDelete.Location = New System.Drawing.Point(40, 120)
        Me.CheckBoxMemDelete.Name = "CheckBoxMemDelete"
        Me.CheckBoxMemDelete.Size = New System.Drawing.Size(136, 24)
        Me.CheckBoxMemDelete.TabIndex = 3
        Me.CheckBoxMemDelete.Text = "Delete"
        '
        'CheckBoxMemApprove
        '
        Me.CheckBoxMemApprove.Location = New System.Drawing.Point(40, 72)
        Me.CheckBoxMemApprove.Name = "CheckBoxMemApprove"
        Me.CheckBoxMemApprove.Size = New System.Drawing.Size(128, 24)
        Me.CheckBoxMemApprove.TabIndex = 2
        Me.CheckBoxMemApprove.Text = "Approve"
        '
        'CheckBoxMemEdit
        '
        Me.CheckBoxMemEdit.Checked = True
        Me.CheckBoxMemEdit.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxMemEdit.Location = New System.Drawing.Point(40, 48)
        Me.CheckBoxMemEdit.Name = "CheckBoxMemEdit"
        Me.CheckBoxMemEdit.Size = New System.Drawing.Size(120, 24)
        Me.CheckBoxMemEdit.TabIndex = 1
        Me.CheckBoxMemEdit.Text = "Edit"
        '
        'CheckBoxMemView
        '
        Me.CheckBoxMemView.Checked = True
        Me.CheckBoxMemView.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxMemView.Location = New System.Drawing.Point(40, 24)
        Me.CheckBoxMemView.Name = "CheckBoxMemView"
        Me.CheckBoxMemView.Size = New System.Drawing.Size(128, 24)
        Me.CheckBoxMemView.TabIndex = 0
        Me.CheckBoxMemView.Text = "View"
        '
        'TabPageOption5
        '
        Me.TabPageOption5.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnUseDefaultFileNames, Me.LabelSetFileNames})
        Me.TabPageOption5.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption5.Name = "TabPageOption5"
        Me.TabPageOption5.Size = New System.Drawing.Size(1020, 269)
        Me.TabPageOption5.TabIndex = 3
        Me.TabPageOption5.Text = "Rename Files"
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
        'TabPageOption6
        '
        Me.TabPageOption6.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnOpenOutputFolder, Me.LabelOutputPath, Me.LabelAfterBuild, Me.Label6})
        Me.TabPageOption6.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption6.Name = "TabPageOption6"
        Me.TabPageOption6.Size = New System.Drawing.Size(1020, 269)
        Me.TabPageOption6.TabIndex = 4
        Me.TabPageOption6.Text = "Create Files"
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
        'TabPageOption2
        '
        Me.TabPageOption2.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBoxDBTable2Name, Me.Label2, Me.TextBoxDBTable1Name, Me.Label1, Me.btnMetadata})
        Me.TabPageOption2.Location = New System.Drawing.Point(4, 25)
        Me.TabPageOption2.Name = "TabPageOption2"
        Me.TabPageOption2.Size = New System.Drawing.Size(529, 270)
        Me.TabPageOption2.TabIndex = 5
        Me.TabPageOption2.Text = "Form Data"
        '
        'TextBoxDBTable2Name
        '
        Me.TextBoxDBTable2Name.Location = New System.Drawing.Point(184, 96)
        Me.TextBoxDBTable2Name.Name = "TextBoxDBTable2Name"
        Me.TextBoxDBTable2Name.Size = New System.Drawing.Size(144, 22)
        Me.TextBoxDBTable2Name.TabIndex = 53
        Me.TextBoxDBTable2Name.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(168, 23)
        Me.Label2.TabIndex = 52
        Me.Label2.Text = "Database Table 2 Name:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBoxDBTable1Name
        '
        Me.TextBoxDBTable1Name.Location = New System.Drawing.Point(184, 72)
        Me.TextBoxDBTable1Name.Name = "TextBoxDBTable1Name"
        Me.TextBoxDBTable1Name.Size = New System.Drawing.Size(144, 22)
        Me.TextBoxDBTable1Name.TabIndex = 51
        Me.TextBoxDBTable1Name.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(168, 23)
        Me.Label1.TabIndex = 50
        Me.Label1.Text = "Database Table 1 Name:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnMetadata
        '
        Me.btnMetadata.Location = New System.Drawing.Point(16, 24)
        Me.btnMetadata.Name = "btnMetadata"
        Me.btnMetadata.Size = New System.Drawing.Size(240, 23)
        Me.btnMetadata.TabIndex = 48
        Me.btnMetadata.Text = "Create Registration Form..."
        '
        'LabelStep3
        '
        Me.LabelStep3.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep3.Location = New System.Drawing.Point(24, 264)
        Me.LabelStep3.Name = "LabelStep3"
        Me.LabelStep3.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep3.TabIndex = 34
        Me.LabelStep3.Text = "3. Choose Accessories"
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
        Me.LabelStep2.TabIndex = 33
        Me.LabelStep2.Text = "2. Create Register Form"
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
        Me.LabelStep1.TabIndex = 32
        Me.LabelStep1.Text = "1. Choose IMS Type"
        Me.LabelStep1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep6
        '
        Me.LabelStep6.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep6.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep6.Location = New System.Drawing.Point(24, 336)
        Me.LabelStep6.Name = "LabelStep6"
        Me.LabelStep6.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep6.TabIndex = 37
        Me.LabelStep6.Text = "6. Create Files"
        Me.LabelStep6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep5
        '
        Me.LabelStep5.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep5.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep5.Location = New System.Drawing.Point(24, 312)
        Me.LabelStep5.Name = "LabelStep5"
        Me.LabelStep5.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep5.TabIndex = 36
        Me.LabelStep5.Text = "5. Rename Files"
        Me.LabelStep5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep4
        '
        Me.LabelStep4.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep4.Location = New System.Drawing.Point(24, 288)
        Me.LabelStep4.Name = "LabelStep4"
        Me.LabelStep4.Size = New System.Drawing.Size(200, 24)
        Me.LabelStep4.TabIndex = 35
        Me.LabelStep4.Text = "4. Choose User Actions"
        Me.LabelStep4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FormGenIMS
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 560)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDisplay, Me.SplitterInstr, Me.PanelInstr, Me.SplitterStep, Me.PanelStep})
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FormGenIMS"
        Me.Text = "Information Management System"
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
        Me.GroupBoxIMSType.ResumeLayout(False)
        Me.TabPageOption3.ResumeLayout(False)
        Me.TabPageOption4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.TabPageOption5.ResumeLayout(False)
        Me.TabPageOption6.ResumeLayout(False)
        Me.TabPageOption2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

 
    ' Nodes
    ' Used by type 1 - non-member reg with admin approval
    Private nodeModuleRoot As New TreeNode("ModuleRoot", 2, 2)

    Private nodeSQLFile As New TreeNode("DBTableFile.sql", 0, 0)
    Private nodeSQLFinalFile As New TreeNode("DBTableFinalFile.sql", 0, 0)
    Private nodeXMLFile As New TreeNode("XMLFile.sql", 0, 0)

    Private nodeIncFolder As New TreeNode("include", 2, 2)
    Private nodeIncFunc As New TreeNode("regFuncs.asp", 0, 0)
    Private nodeIncEditTable As New TreeNode("editTable.asp", 0, 0)
    Private nodeIncConfirmTable As New TreeNode("confirmTable.asp", 0, 0)

    Private nodeRegEdit As New TreeNode("register.asp", 0, 0)
    Private nodeRegConfirm As New TreeNode("registerConfirm.asp", 0, 0)
    Private nodeAdminRegList As New TreeNode("adminRegList.asp", 0, 0)
    Private nodeAdminRegEdit As New TreeNode("adminRegEdit.asp", 0, 0)
    Private nodeAdminRegConfirm As New TreeNode("adminRegConfirm.asp", 0, 0)
    ' Extra files used by type 2
    Private nodeMemRegList As New TreeNode("memRegList.asp", 0, 0)
    Private nodeMemRegEdit As New TreeNode("memRegEdit.asp", 0, 0)
    Private nodeMemRegConfirm As New TreeNode("memRegConfirm.asp", 0, 0)
    ' Extra files used by type 3
    Private nodeAdminFinalEdit As New TreeNode("adminFinalEdit.asp", 0, 0)
    Private nodeAdminFinalConfirm As New TreeNode("adminFinalConfirm.asp", 0, 0)


    Private Sub FormGenIMS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.moduleType = ModuleMain.xcModuleType.xcIMS
        Me.NUM_STEPS = 6
        Me.initModuleContent()
    End Sub


    Private Sub FormGenIMS_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize

    End Sub



    Protected Overrides Sub initTreeNodesText()
        ' Used by type 1 - non-member reg with admin approval
        nodeModuleRoot.Text = Me.moduleName
        nodeSQLFile.Text = Me.moduleName & ".sql"
        nodeSQLFinalFile.Text = Me.moduleName & "Final.sql"
        nodeXMLFile.Text = Me.moduleName & ".xml"

        nodeIncFolder.Text = "include"
        nodeIncFunc.Text = "regFuncs.asp"
        nodeIncEditTable.Text = "editTable.asp"
        nodeIncConfirmTable.Text = "confirmTable.asp"
        nodeRegEdit.Text = "register.asp"
        nodeRegConfirm.Text = "registerConfirm.asp"
        nodeAdminRegList.Text = "adminRegList.asp"
        nodeAdminRegEdit.Text = "adminRegEdit.asp"
        nodeAdminRegConfirm.Text = "adminRegConfirm.asp"
        ' Extra files used by type 2
        nodeMemRegList.Text = "memRegList.asp"
        nodeMemRegEdit.Text = "memRegEdit.asp"
        nodeMemRegConfirm.Text = "memRegConfirm.asp"
        ' Extra files used by type 3
        nodeAdminFinalEdit.Text = "adminFinalEdit.asp"
        nodeAdminFinalConfirm.Text = "adminFinalConfirm.asp"
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
        Me.ComboBoxIMSType.SelectedIndex = 0
        Me.TextBoxDBTable1Name.Text = Me.moduleName
        Me.TextBoxDBTable2Name.Text = Me.moduleName & "Final"

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
                    ' Interface options
                ElseIf lcase_line.StartsWith("databasetable1name") Then
                    Me.TextBoxDBTable1Name.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("databasetable2name") Then
                    Me.TextBoxDBTable2Name.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("imstype") Then
                    Me.ComboBoxIMSType.SelectedIndex = CType(MyUtil.getStrAfterDelimiter(line, "="), Integer)
                ElseIf lcase_line.StartsWith("useaccessorysearch") Then
                    Me.CheckBoxAccessorySearch.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("useaccessorycalendar") Then
                    Me.CheckBoxAccessoryCalendar.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("useaccessoryaccess") Then
                    Me.CheckBoxAccessoryAccess.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("useaccessoryword") Then
                    Me.CheckBoxAccessoryWord.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("userrightnonMemView") Then
                    Me.CheckBoxNonMemView.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("userrightmemview") Then
                    Me.CheckBoxMemView.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("userrightmemedit") Then
                    Me.CheckBoxMemEdit.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("userrightmemapprove") Then
                    Me.CheckBoxMemApprove.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("userrightmemdisapprove") Then
                    Me.CheckBoxMemDisapprove.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("userrightmemdelete") Then
                    Me.CheckBoxMemDelete.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("userrightadmview") Then
                    Me.CheckBoxAdmView.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("userrightadmedit") Then
                    Me.CheckBoxAdmView.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("userrightadmapprove") Then
                    Me.CheckBoxAdmView.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("userrightadmdisapprove") Then
                    Me.CheckBoxAdmView.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                ElseIf lcase_line.StartsWith("Userrightadmdelete") Then
                    Me.CheckBoxAdmView.Checked = CType(MyUtil.getStrAfterDelimiter(line, "="), Boolean)
                    ' Start - Node Text
                ElseIf lcase_line.StartsWith("incfolder") Then
                    Me.nodeIncFolder.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("incfuncfile") Then
                    Me.nodeIncFunc.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("incedittablefile") Then
                    Me.nodeIncEditTable.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("incconfirmtablefile") Then
                    Me.nodeIncConfirmTable.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("regeditfile") Then
                    Me.nodeRegEdit.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("regconfirmfile") Then
                    Me.nodeRegConfirm.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminreglistfile") Then
                    Me.nodeAdminRegList.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminregeditfile") Then
                    Me.nodeAdminRegEdit.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminregconfirmfile") Then
                    Me.nodeAdminRegConfirm.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("memreglistfile") Then
                    Me.nodeMemRegList.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("memregeditfile") Then
                    Me.nodeMemRegEdit.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("memregconfirmfile") Then
                    Me.nodeMemRegConfirm.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminfinaleditfile") Then
                    Me.nodeAdminFinalEdit.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("adminfinalconfirmfile") Then
                    Me.nodeAdminFinalConfirm.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("sqlfile") Then
                    Me.nodeSQLFile.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("sqlfinalfile") Then
                    Me.nodeSQLFinalFile.Text = MyUtil.getStrAfterDelimiter(line, "=")
                ElseIf lcase_line.StartsWith("xmlfile") Then
                    Me.nodeXMLFile.Text = MyUtil.getStrAfterDelimiter(line, "=")
                    ' End - Node Text
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
                    ' End - used by default layout table attributes               
                End If
            Next
        End If

        constructTreeViewFileList()
        Me.processStep()

    End Sub


    Protected Overrides Sub constructTreeViewFileList()
        Dim tv As TreeView = Me.TreeView1
        If tv Is Nothing Then Exit Sub

        Try
            tv.Nodes.Clear()
            Me.nodeModuleRoot.Nodes.Clear()
            Me.nodeIncFolder.Nodes.Clear()

            tv.Nodes.Add(Me.nodeModuleRoot)
            Me.nodeModuleRoot.Nodes.Add(Me.nodeIncFolder)

            Me.nodeIncFolder.Nodes.Add(Me.nodeIncEditTable)
            Me.nodeIncFolder.Nodes.Add(Me.nodeIncConfirmTable)
            Me.nodeIncFolder.Nodes.Add(Me.nodeIncFunc)

            Me.nodeModuleRoot.Nodes.Add(Me.nodeRegEdit)
            Me.nodeModuleRoot.Nodes.Add(Me.nodeRegConfirm)

            If Me.ComboBoxIMSType.SelectedIndex = 1 Or _
               Me.ComboBoxIMSType.SelectedIndex = 2 Then
                Me.nodeModuleRoot.Nodes.Add(Me.nodeMemRegList)
                Me.nodeModuleRoot.Nodes.Add(Me.nodeMemRegEdit)
                Me.nodeModuleRoot.Nodes.Add(Me.nodeMemRegConfirm)
            End If

            Me.nodeModuleRoot.Nodes.Add(Me.nodeAdminRegList)
            Me.nodeModuleRoot.Nodes.Add(Me.nodeAdminRegEdit)
            Me.nodeModuleRoot.Nodes.Add(Me.nodeAdminRegConfirm)

            If Me.ComboBoxIMSType.SelectedIndex = 2 Then
                Me.nodeModuleRoot.Nodes.Add(Me.nodeAdminFinalEdit)
                Me.nodeModuleRoot.Nodes.Add(Me.nodeAdminFinalConfirm)
            End If

            Me.nodeModuleRoot.Nodes.Add(Me.nodeSQLFile)
            If Me.ComboBoxIMSType.SelectedIndex = 2 Then
                Me.nodeModuleRoot.Nodes.Add(Me.nodeSQLFinalFile)
            End If

            Me.nodeModuleRoot.Nodes.Add(Me.nodeXMLFile)

            tv.ExpandAll()
        Catch ex As Exception
            MsgBox("ConstructTreeViewFileList Error:" & vbCrLf & ex.ToString)
            End
        End Try
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
        str += "DatabaseTable1Name=" & Me.TextBoxDBTable1Name.Text & vbCrLf
        str += "DatabaseTable2Name=" & Me.TextBoxDBTable2Name.Text & vbCrLf

        str += "IMSType=" & Me.ComboBoxIMSType.SelectedIndex & vbCrLf
        str += vbCrLf
        str += "UseAccessorySearch=" & Me.CheckBoxAccessorySearch.Checked & vbCrLf
        str += "UseAccessoryCalendar=" & Me.CheckBoxAccessoryCalendar.Checked & vbCrLf
        str += "UseAccessoryAccess=" & Me.CheckBoxAccessoryAccess.Checked & vbCrLf
        str += "UseAccessoryWord=" & Me.CheckBoxAccessoryWord.Checked & vbCrLf
        str += vbCrLf
        str += "UserRightNonMemView=" & Me.CheckBoxNonMemView.Checked & vbCrLf
        str += "UserRightMemView=" & Me.CheckBoxMemView.Checked & vbCrLf
        str += "UserRightMemEdit=" & Me.CheckBoxMemEdit.Checked & vbCrLf
        str += "UserRightMemApprove=" & Me.CheckBoxMemApprove.Checked & vbCrLf
        str += "UserRightMemDisapprove=" & Me.CheckBoxMemDisapprove.Checked & vbCrLf
        str += "UserRightMemDelete=" & Me.CheckBoxMemDelete.Checked & vbCrLf
        str += "UserRightAdmView=" & Me.CheckBoxAdmView.Checked & vbCrLf
        str += "UserRightAdmEdit=" & Me.CheckBoxAdmEdit.Checked & vbCrLf
        str += "UserRightAdmApprove=" & Me.CheckBoxAdmApprove.Checked & vbCrLf
        str += "UserRightAdmDisapprove=" & Me.CheckBoxAdmDisapprove.Checked & vbCrLf
        str += "UserRightAdmDelete=" & Me.CheckBoxAdmDelete.Checked & vbCrLf
        str += vbCrLf
        ' Start - Nodes text
        str += "IncFolder=" & Me.nodeIncFolder.Text & vbCrLf
        str += "IncFuncFile=" & Me.nodeIncFunc.Text & vbCrLf
        str += "IncEditTableFile=" & Me.nodeIncEditTable.Text & vbCrLf
        str += "IncConfirmTableFile=" & Me.nodeIncConfirmTable.Text & vbCrLf
        str += "RegEditFile=" & Me.nodeRegEdit.Text & vbCrLf
        str += "RegConfirmFile=" & Me.nodeRegConfirm.Text & vbCrLf
        str += "AdminRegListFile=" & Me.nodeAdminRegList.Text & vbCrLf
        str += "AdminRegEditFile=" & Me.nodeAdminRegEdit.Text & vbCrLf
        str += "AdminRegConfirmFile=" & Me.nodeAdminRegConfirm.Text & vbCrLf
        str += "' Extra files used by Type 1 IMS" & vbCrLf
        str += "MemRegListFile=" & Me.nodeMemRegList.Text & vbCrLf
        str += "MemRegEditFile=" & Me.nodeMemRegEdit.Text & vbCrLf
        str += "MemRegConfirmFile=" & Me.nodeMemRegConfirm.Text & vbCrLf
        str += "' Extra files used by Type 2 IMS" & vbCrLf
        str += "AdminFinalEditFile=" & Me.nodeAdminFinalEdit.Text & vbCrLf
        str += "AdminFinalConfirmFile=" & Me.nodeAdminFinalConfirm.Text & vbCrLf
        str += "SQLFile=" & Me.nodeSQLFile.Text & vbCrLf
        str += "SQLFinalFile=" & Me.nodeSQLFinalFile.Text & vbCrLf
        str += "XMLFile=" & Me.nodeXMLFile.Text & vbCrLf
        str += vbCrLf
        ' End - Nodes text
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
        ' Save metadata
        IOManager.SaveTextToFile(Me.metadata, IOManager.getFolder(Me.moduleFullPath) & Me.moduleName & ".metadata.txt")

        Me.infoSaved = True
    End Sub


    Private Sub constructFrameworkParams(ByRef frameworkParams As clsWebsiteFrameworkParameters)
        Dim f As clsWebsiteFrameworkParameters = frameworkParams
        Me.globalHomePath = "../" & f.homeFile
        Me.globalCssFilePath = "../" & f.cssFile
        Me.headerFilePath = "../" & f.rootIncFolder & "/" & f.rootIncHeaderFile
        Me.footerFilePath = "../" & f.rootIncFolder & "/" & f.rootIncFooterFile
        Me.adminSecurityFilePath = "../" & f.loginFolder & "/" & f.loginAdmSecurityFile
        Me.memSecurityFilePath = "../" & f.loginFolder & "/" & f.loginMemSecurityFile
        Me.globleFuncFilePath = "../../" & f.aspBinFolder & "/" & f.aspFuncLibFile
        Me.adminHomeFilePath = "../" & f.admFolder & "/" & f.admHomeFile
        Me.memHomeFilePath = "../" & f.memFolder & "/" & f.memHomeFile
        Me.imageFolderPath = "../" & f.imageFolder
    End Sub


    Protected Friend Overrides Sub buildModule( _
    Optional ByVal frameworkParams As clsWebsiteFrameworkParameters = Nothing)
        Dim outputFolder, outputIncFolder As String

        If Trim(Me.metadata) = "" Then
            MsgBox("Registration Form metadata is empty. Please create the registration form.")
            Exit Sub
        End If

        ' Clear generated files of last time.
        'MsgBox(Me.moduleOutputPath)
        If IOManager.folderExists(moduleOutputPath) Then
            IOManager.DeleteFolder(moduleOutputPath, True)
        End If
        IOManager.CreateFolder(moduleOutputPath)

        ' Contruct website framework settings.
        Me.constructFrameworkParams(frameworkParams)

        Dim clsIMS As New clsModuleIMS()
        clsIMS.init(Me.metadata, Me.moduleName, _
               Me.TextBoxDBTable1Name.Text, Me.TextBoxDBTable2Name.Text, _
               Me.globalHomePath, Me.globleFuncFilePath, _
               Me.adminSecurityFilePath, Me.memSecurityFilePath, _
               Me.headerFilePath, Me.footerFilePath, _
               Me.adminHomeFilePath, Me.memHomeFilePath, _
               Me.nodeRegEdit.Text, Me.nodeRegConfirm.Text, _
               Me.nodeMemRegList.Text, Me.nodeMemRegEdit.Text, Me.nodeMemRegConfirm.Text, _
               Me.nodeAdminRegList.Text, Me.nodeAdminRegEdit.Text, Me.nodeAdminRegConfirm.Text, _
               Me.nodeAdminFinalEdit.Text, Me.nodeAdminFinalConfirm.Text, _
               Me.nodeIncFolder.Text, Me.nodeIncFunc.Text, _
               Me.nodeIncEditTable.Text, Me.nodeIncConfirmTable.Text, _
               Me.ComboBoxIMSType.SelectedIndex, _
               Me.CheckBoxAccessorySearch.Checked, Me.CheckBoxAccessoryCalendar.Checked, _
               Me.CheckBoxAccessoryAccess.Checked, Me.CheckBoxAccessoryWord.Checked, _
               Me.CheckBoxNonMemView.Checked, Me.CheckBoxMemView.Checked, _
               Me.CheckBoxMemEdit.Checked, Me.CheckBoxMemApprove.Checked, _
               Me.CheckBoxMemDisapprove.Checked, Me.CheckBoxMemDelete.Checked, _
               Me.CheckBoxAdmView.Checked, Me.CheckBoxAdmEdit.Checked, _
               Me.CheckBoxAdmApprove.Checked, Me.CheckBoxAdmDisapprove.Checked, _
               Me.CheckBoxAdmDelete.Checked)

        Dim progressBar As New FormProgressBar()
        progressBar.setText(Me.moduleName & SetupManager.MODULE_FILE_EXTENSION)
        progressBar.Show()
        progressBar.setProgress(1)

        If Me.ComboBoxIMSType.SelectedIndex = 0 Or _
           Me.ComboBoxIMSType.SelectedIndex = 1 Or _
           Me.ComboBoxIMSType.SelectedIndex = 2 Then
            IOManager.SaveTextToFile(clsIMS.generateDBTableSQL(Me.TextBoxDBTable1Name.Text, Me.ComboBoxIMSType.SelectedIndex), Me.moduleOutputPath & Me.nodeSQLFile.Text)
            IOManager.SaveTextToFile(clsIMS.generateXMLFile(), Me.moduleOutputPath & Me.nodeXMLFile.Text)

            IOManager.SaveTextToFile(clsIMS.generateRegEditFile(), Me.moduleOutputPath & Me.nodeRegEdit.Text)
            progressBar.setProgress(1)
            IOManager.SaveTextToFile(clsIMS.generateRegConfirmFile(), Me.moduleOutputPath & Me.nodeRegConfirm.Text)
            progressBar.setProgress(1)
            IOManager.SaveTextToFile(clsIMS.generateAdmListFile(), Me.moduleOutputPath & Me.nodeAdminRegList.Text)
            progressBar.setProgress(1)
            IOManager.SaveTextToFile(clsIMS.generateAdmEditFile(), Me.moduleOutputPath & Me.nodeAdminRegEdit.Text)
            progressBar.setProgress(1)
            IOManager.SaveTextToFile(clsIMS.generateAdmConfirmFile(), Me.moduleOutputPath & Me.nodeAdminRegConfirm.Text)
            progressBar.setProgress(1)

            outputIncFolder = Me.moduleOutputPath & Me.nodeIncFolder.Text & "/"
            IOManager.CreateFolder(outputIncFolder)
            progressBar.setProgress(1)
            IOManager.SaveTextToFile(clsIMS.generateIncEditTableFile(Me.tblAttr), outputIncFolder & Me.nodeIncEditTable.Text)
            progressBar.setProgress(1)
            IOManager.SaveTextToFile(clsIMS.generateIncConfirmTableFile(Me.tblAttr), outputIncFolder & Me.nodeIncConfirmTable.Text)
            progressBar.setProgress(1)
            IOManager.SaveTextToFile(clsIMS.generateIncFuncFile(), outputIncFolder & Me.nodeIncFunc.Text)
            progressBar.setProgress(1)

            clsIMS.insertLinksToUI(Me.projectOutputPath & "\admin\index.asp", _
                   xcUtil.admTaskListStart, xcUtil.admTaskListEnd, _
                   "../" & Me.moduleName & "/" & Me.nodeAdminRegList.Text)
        End If

        If Me.ComboBoxIMSType.SelectedIndex = 1 Or _
           Me.ComboBoxIMSType.SelectedIndex = 2 Then
            IOManager.SaveTextToFile(clsIMS.generateMemListFile(), Me.moduleOutputPath & Me.nodeMemRegList.Text)
            'progressBar.setProgress(1)
            IOManager.SaveTextToFile(clsIMS.generateMemEditFile(), Me.moduleOutputPath & Me.nodeMemRegEdit.Text)
            'progressBar.setProgress(1)
            IOManager.SaveTextToFile(clsIMS.generateMemConfirmFile(), Me.moduleOutputPath & Me.nodeMemRegConfirm.Text)

            clsIMS.insertLinksToUI(Me.projectOutputPath & "\member\index.asp", _
                   xcUtil.memTaskListStart, xcUtil.memTaskListEnd, _
                   "../" & Me.moduleName & "/" & Me.nodeMemRegList.Text)
        End If

        If Me.ComboBoxIMSType.SelectedIndex = 2 Then
            IOManager.SaveTextToFile(clsIMS.generateDBTableSQL(Me.TextBoxDBTable2Name.Text, Me.ComboBoxIMSType.SelectedIndex, "final"), Me.moduleOutputPath & Me.nodeSQLFinalFile.Text)
            IOManager.SaveTextToFile(clsIMS.generateAdmEditFinalFile(), Me.moduleOutputPath & Me.nodeAdminFinalEdit.Text)
            IOManager.SaveTextToFile(clsIMS.generateAdmConfirmFinalFile(), Me.moduleOutputPath & Me.nodeAdminFinalConfirm.Text)
        End If

        progressBar.Close()

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Interface functions
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
                Me.TabControlDisplay.SelectedTab = Me.TabPageQuickView
                Me.LabelStepInstr.Text = "Type of IMS to be generated. Currently 3 types are available."
            Case 1
                Me.stepNumber = 2
                Me.TabControlDisplay.SelectedTab = Me.TabPageFileList
                Me.LabelStepInstr.Text = "Create IMS module registration form"
            Case 2
                Me.stepNumber = 3
                Me.LabelStepInstr.Text = "Available additional functions besides the basic IMS function: search, calendar etc."
            Case 3
                Me.stepNumber = 4
                Me.LabelStepInstr.Text = "What each type of user (non-member, member, admin) can do."
            Case 4
                Me.stepNumber = 5
                Me.LabelStepInstr.Text = "You don't really need to change filenames for the module to work. But in case you prefer other names, you can do it."
            Case 5
                Me.stepNumber = 6
                Me.LabelStepInstr.Text = "The shortcut to build project is F5. Only opening modules in a project are built. The shortcut to build only the currently active module is F6."
        End Select
        Me.processStep()
        prevSelectedIndex = Me.TabControlOptions.SelectedIndex
    End Sub

    Private Sub ComboBoxIMSType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxIMSType.SelectedIndexChanged
        Select Case Me.ComboBoxIMSType.SelectedIndex
            Case 0
                Me.LabelIMSTypeDescription.Text = "Non-member registration with admin approval. "
                Me.TextBoxDBTable2Name.Enabled = False
                Me.AxWebBrowser1.Navigate(SetupManager.IMS_LIB_PATH & "flowchart1.jpg")
            Case 1
                Me.LabelIMSTypeDescription.Text = "Non-member or Member registration with admin approval, and member update without admin approval. "
                Me.TextBoxDBTable2Name.Enabled = False
                Me.AxWebBrowser1.Navigate(SetupManager.IMS_LIB_PATH & "flowchart2.jpg")
            Case 2
                Me.LabelIMSTypeDescription.Text = "Non-member or Member registration with admin approval, and member update with admin approval. "
                Me.TextBoxDBTable2Name.Enabled = True
                Me.AxWebBrowser1.Navigate(SetupManager.IMS_LIB_PATH & "flowchart3.jpg")
        End Select
        ' Reconstruct TreeView after selecting a different IMS type
        Me.constructTreeViewFileList()
    End Sub

    Private Sub btnMetadata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMetadata.Click

        Dim mdWizard As New FormMetadataWizard()
        'MsgBox("tmpfolder: " & Me.tmpFileFolder)
        mdWizard.init(Me.tmpFileFolder, Me.metadata, Me.tblAttr, Me)
        mdWizard.ShowDialog(Me)

        Me.RichTextBox1.Text = Me.metadata
        'MsgBox(Me.metadata)
        Me.TabControlDisplay.SelectedTab = Me.TabPageSource

        Me.RichTextBox1.Text = Me.metadata
    End Sub


End Class
