Public Class FormChildBase
    Inherits System.Windows.Forms.Form


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
    Public WithEvents GroupBoxInstr As System.Windows.Forms.GroupBox
    Public WithEvents LabelInstr As System.Windows.Forms.Label
    Public WithEvents SplitterInstrDetail As System.Windows.Forms.Splitter
    Public WithEvents TabPageQuickView As System.Windows.Forms.TabPage
    Public WithEvents TabPageSource As System.Windows.Forms.TabPage
    Public WithEvents SplitterInstr As System.Windows.Forms.Splitter
    Public WithEvents PanelStepTitle As System.Windows.Forms.Panel
    Public WithEvents LabelStepTitle As System.Windows.Forms.Label
    Public WithEvents PanelInstrTitle As System.Windows.Forms.Panel
    Public WithEvents btnCloseInstr As System.Windows.Forms.Button
    Public WithEvents LabelInstrTitle As System.Windows.Forms.Label
    Public WithEvents PanelStepDetail As System.Windows.Forms.Panel
    Public WithEvents MainMenuChild As System.Windows.Forms.MainMenu
    Public WithEvents ImageListChild As System.Windows.Forms.ImageList
    Public WithEvents btnCloseStep As System.Windows.Forms.Button
    Public WithEvents PanelDisplay As System.Windows.Forms.Panel
    Public WithEvents PanelDispTitle As System.Windows.Forms.Panel
    Public WithEvents LabelDispTitle As System.Windows.Forms.Label
    Public WithEvents btnCloseDisp As System.Windows.Forms.Button
    Public WithEvents TabControlDisplay As System.Windows.Forms.TabControl
    Public WithEvents AxWebBrowser1 As AxSHDocVw.AxWebBrowser
    Public WithEvents ContextMenuStep As System.Windows.Forms.ContextMenu
    Public WithEvents mnuStepFloat As System.Windows.Forms.MenuItem
    Public WithEvents mnuStepHide As System.Windows.Forms.MenuItem
    Public WithEvents PanelStep As System.Windows.Forms.Panel
    Public WithEvents SplitterStep As System.Windows.Forms.Splitter
    Public WithEvents PanelInstr As System.Windows.Forms.Panel
    Public WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Public WithEvents ContextMenuDisp As System.Windows.Forms.ContextMenu
    Public WithEvents mnuDispFloat As System.Windows.Forms.MenuItem
    Public WithEvents mnuDispHide As System.Windows.Forms.MenuItem
    Public WithEvents ContextMenuInstr As System.Windows.Forms.ContextMenu
    Public WithEvents mnuInstrFloat As System.Windows.Forms.MenuItem
    Public WithEvents mnuInstrHide As System.Windows.Forms.MenuItem
    Public WithEvents ContextMenuFileList As System.Windows.Forms.ContextMenu
    Public WithEvents mnuFileListFloat As System.Windows.Forms.MenuItem
    Public WithEvents mnuFileListHide As System.Windows.Forms.MenuItem
    Public WithEvents ContextMenuFileDesc As System.Windows.Forms.ContextMenu
    Public WithEvents mnuFileDescFloat As System.Windows.Forms.MenuItem
    Public WithEvents mnuFileDescHide As System.Windows.Forms.MenuItem
    Public WithEvents TreeView1 As System.Windows.Forms.TreeView
    Public WithEvents TabPageFileList As System.Windows.Forms.TabPage
    Public WithEvents TabControlOptions As System.Windows.Forms.TabControl
    Public WithEvents GroupBoxStep As System.Windows.Forms.GroupBox
    Public WithEvents LabelStepInstr As System.Windows.Forms.Label
    Public WithEvents ContextMenuFileNode As System.Windows.Forms.ContextMenu
    Public WithEvents mnuFileNodeRename As System.Windows.Forms.MenuItem
    Protected WithEvents btnNext As System.Windows.Forms.Button
    Protected WithEvents btnBack As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormChildBase))
        Me.PanelStep = New System.Windows.Forms.Panel()
        Me.ContextMenuStep = New System.Windows.Forms.ContextMenu()
        Me.mnuStepFloat = New System.Windows.Forms.MenuItem()
        Me.mnuStepHide = New System.Windows.Forms.MenuItem()
        Me.PanelStepDetail = New System.Windows.Forms.Panel()
        Me.GroupBoxStep = New System.Windows.Forms.GroupBox()
        Me.LabelStepInstr = New System.Windows.Forms.Label()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.PanelStepTitle = New System.Windows.Forms.Panel()
        Me.btnCloseStep = New System.Windows.Forms.Button()
        Me.LabelStepTitle = New System.Windows.Forms.Label()
        Me.SplitterStep = New System.Windows.Forms.Splitter()
        Me.ContextMenuFileList = New System.Windows.Forms.ContextMenu()
        Me.mnuFileListFloat = New System.Windows.Forms.MenuItem()
        Me.mnuFileListHide = New System.Windows.Forms.MenuItem()
        Me.ContextMenuFileDesc = New System.Windows.Forms.ContextMenu()
        Me.mnuFileDescFloat = New System.Windows.Forms.MenuItem()
        Me.mnuFileDescHide = New System.Windows.Forms.MenuItem()
        Me.PanelInstr = New System.Windows.Forms.Panel()
        Me.ContextMenuInstr = New System.Windows.Forms.ContextMenu()
        Me.mnuInstrFloat = New System.Windows.Forms.MenuItem()
        Me.mnuInstrHide = New System.Windows.Forms.MenuItem()
        Me.TabControlOptions = New System.Windows.Forms.TabControl()
        Me.SplitterInstrDetail = New System.Windows.Forms.Splitter()
        Me.GroupBoxInstr = New System.Windows.Forms.GroupBox()
        Me.LabelInstr = New System.Windows.Forms.Label()
        Me.PanelInstrTitle = New System.Windows.Forms.Panel()
        Me.LabelInstrTitle = New System.Windows.Forms.Label()
        Me.btnCloseInstr = New System.Windows.Forms.Button()
        Me.SplitterInstr = New System.Windows.Forms.Splitter()
        Me.MainMenuChild = New System.Windows.Forms.MainMenu()
        Me.ImageListChild = New System.Windows.Forms.ImageList(Me.components)
        Me.PanelDisplay = New System.Windows.Forms.Panel()
        Me.TabControlDisplay = New System.Windows.Forms.TabControl()
        Me.TabPageQuickView = New System.Windows.Forms.TabPage()
        Me.AxWebBrowser1 = New AxSHDocVw.AxWebBrowser()
        Me.TabPageSource = New System.Windows.Forms.TabPage()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.TabPageFileList = New System.Windows.Forms.TabPage()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.PanelDispTitle = New System.Windows.Forms.Panel()
        Me.LabelDispTitle = New System.Windows.Forms.Label()
        Me.ContextMenuDisp = New System.Windows.Forms.ContextMenu()
        Me.mnuDispFloat = New System.Windows.Forms.MenuItem()
        Me.mnuDispHide = New System.Windows.Forms.MenuItem()
        Me.btnCloseDisp = New System.Windows.Forms.Button()
        Me.ContextMenuFileNode = New System.Windows.Forms.ContextMenu()
        Me.mnuFileNodeRename = New System.Windows.Forms.MenuItem()
        Me.PanelStep.SuspendLayout()
        Me.PanelStepDetail.SuspendLayout()
        Me.GroupBoxStep.SuspendLayout()
        Me.PanelStepTitle.SuspendLayout()
        Me.PanelInstr.SuspendLayout()
        Me.GroupBoxInstr.SuspendLayout()
        Me.PanelInstrTitle.SuspendLayout()
        Me.PanelDisplay.SuspendLayout()
        Me.TabControlDisplay.SuspendLayout()
        Me.TabPageQuickView.SuspendLayout()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPageSource.SuspendLayout()
        Me.TabPageFileList.SuspendLayout()
        Me.PanelDispTitle.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelStep
        '
        Me.PanelStep.BackColor = System.Drawing.SystemColors.ControlLight
        Me.PanelStep.ContextMenu = Me.ContextMenuStep
        Me.PanelStep.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelStepDetail, Me.PanelStepTitle})
        Me.PanelStep.Dock = System.Windows.Forms.DockStyle.Left
        Me.PanelStep.DockPadding.All = 1
        Me.PanelStep.Name = "PanelStep"
        Me.PanelStep.Size = New System.Drawing.Size(250, 560)
        Me.PanelStep.TabIndex = 1
        '
        'ContextMenuStep
        '
        Me.ContextMenuStep.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuStepFloat, Me.mnuStepHide})
        '
        'mnuStepFloat
        '
        Me.mnuStepFloat.Index = 0
        Me.mnuStepFloat.Text = "Float"
        '
        'mnuStepHide
        '
        Me.mnuStepHide.Index = 1
        Me.mnuStepHide.Text = "Hide"
        '
        'PanelStepDetail
        '
        Me.PanelStepDetail.BackColor = System.Drawing.SystemColors.ControlLight
        Me.PanelStepDetail.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBoxStep, Me.btnNext, Me.btnBack})
        Me.PanelStepDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelStepDetail.Location = New System.Drawing.Point(1, 26)
        Me.PanelStepDetail.Name = "PanelStepDetail"
        Me.PanelStepDetail.Size = New System.Drawing.Size(248, 533)
        Me.PanelStepDetail.TabIndex = 6
        '
        'GroupBoxStep
        '
        Me.GroupBoxStep.BackColor = System.Drawing.SystemColors.ControlLight
        Me.GroupBoxStep.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelStepInstr})
        Me.GroupBoxStep.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GroupBoxStep.Location = New System.Drawing.Point(24, 16)
        Me.GroupBoxStep.Name = "GroupBoxStep"
        Me.GroupBoxStep.Size = New System.Drawing.Size(200, 160)
        Me.GroupBoxStep.TabIndex = 11
        Me.GroupBoxStep.TabStop = False
        Me.GroupBoxStep.Text = "Step Instruction"
        '
        'LabelStepInstr
        '
        Me.LabelStepInstr.BackColor = System.Drawing.SystemColors.ControlLight
        Me.LabelStepInstr.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelStepInstr.Location = New System.Drawing.Point(16, 40)
        Me.LabelStepInstr.Name = "LabelStepInstr"
        Me.LabelStepInstr.Size = New System.Drawing.Size(168, 104)
        Me.LabelStepInstr.TabIndex = 0
        '
        'btnNext
        '
        Me.btnNext.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.btnNext.BackColor = System.Drawing.SystemColors.Control
        Me.btnNext.ForeColor = System.Drawing.SystemColors.WindowText
        Me.btnNext.Location = New System.Drawing.Point(125, 472)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(100, 32)
        Me.btnNext.TabIndex = 4
        Me.btnNext.Text = "Next"
        '
        'btnBack
        '
        Me.btnBack.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.btnBack.BackColor = System.Drawing.SystemColors.Control
        Me.btnBack.ForeColor = System.Drawing.SystemColors.WindowText
        Me.btnBack.Location = New System.Drawing.Point(25, 472)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(100, 32)
        Me.btnBack.TabIndex = 3
        Me.btnBack.Text = "Back"
        '
        'PanelStepTitle
        '
        Me.PanelStepTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCloseStep, Me.LabelStepTitle})
        Me.PanelStepTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelStepTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelStepTitle.Name = "PanelStepTitle"
        Me.PanelStepTitle.Size = New System.Drawing.Size(248, 25)
        Me.PanelStepTitle.TabIndex = 5
        '
        'btnCloseStep
        '
        Me.btnCloseStep.BackColor = System.Drawing.SystemColors.Control
        Me.btnCloseStep.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseStep.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseStep.Location = New System.Drawing.Point(223, 0)
        Me.btnCloseStep.Name = "btnCloseStep"
        Me.btnCloseStep.Size = New System.Drawing.Size(25, 25)
        Me.btnCloseStep.TabIndex = 2
        Me.btnCloseStep.Text = "X"
        '
        'LabelStepTitle
        '
        Me.LabelStepTitle.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.LabelStepTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelStepTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStepTitle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LabelStepTitle.Name = "LabelStepTitle"
        Me.LabelStepTitle.Size = New System.Drawing.Size(248, 25)
        Me.LabelStepTitle.TabIndex = 1
        Me.LabelStepTitle.Text = " Step Explorer"
        Me.LabelStepTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SplitterStep
        '
        Me.SplitterStep.Location = New System.Drawing.Point(250, 0)
        Me.SplitterStep.Name = "SplitterStep"
        Me.SplitterStep.Size = New System.Drawing.Size(3, 560)
        Me.SplitterStep.TabIndex = 2
        Me.SplitterStep.TabStop = False
        '
        'ContextMenuFileList
        '
        Me.ContextMenuFileList.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileListFloat, Me.mnuFileListHide})
        '
        'mnuFileListFloat
        '
        Me.mnuFileListFloat.Index = 0
        Me.mnuFileListFloat.Text = "Float"
        '
        'mnuFileListHide
        '
        Me.mnuFileListHide.Index = 1
        Me.mnuFileListHide.Text = "Hide"
        '
        'ContextMenuFileDesc
        '
        Me.ContextMenuFileDesc.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileDescFloat, Me.mnuFileDescHide})
        '
        'mnuFileDescFloat
        '
        Me.mnuFileDescFloat.Index = 0
        Me.mnuFileDescFloat.Text = "Float"
        '
        'mnuFileDescHide
        '
        Me.mnuFileDescHide.Index = 1
        Me.mnuFileDescHide.Text = "Hide"
        '
        'PanelInstr
        '
        Me.PanelInstr.ContextMenu = Me.ContextMenuInstr
        Me.PanelInstr.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControlOptions, Me.SplitterInstrDetail, Me.GroupBoxInstr, Me.PanelInstrTitle})
        Me.PanelInstr.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelInstr.DockPadding.All = 1
        Me.PanelInstr.Location = New System.Drawing.Point(253, 160)
        Me.PanelInstr.Name = "PanelInstr"
        Me.PanelInstr.Size = New System.Drawing.Size(539, 400)
        Me.PanelInstr.TabIndex = 9
        '
        'ContextMenuInstr
        '
        Me.ContextMenuInstr.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuInstrFloat, Me.mnuInstrHide})
        '
        'mnuInstrFloat
        '
        Me.mnuInstrFloat.Index = 0
        Me.mnuInstrFloat.Text = "Float"
        '
        'mnuInstrHide
        '
        Me.mnuInstrHide.Index = 1
        Me.mnuInstrHide.Text = "Hide"
        '
        'TabControlOptions
        '
        Me.TabControlOptions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControlOptions.Location = New System.Drawing.Point(1, 26)
        Me.TabControlOptions.Multiline = True
        Me.TabControlOptions.Name = "TabControlOptions"
        Me.TabControlOptions.SelectedIndex = 0
        Me.TabControlOptions.ShowToolTips = True
        Me.TabControlOptions.Size = New System.Drawing.Size(537, 298)
        Me.TabControlOptions.TabIndex = 10
        '
        'SplitterInstrDetail
        '
        Me.SplitterInstrDetail.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.SplitterInstrDetail.Location = New System.Drawing.Point(1, 324)
        Me.SplitterInstrDetail.Name = "SplitterInstrDetail"
        Me.SplitterInstrDetail.Size = New System.Drawing.Size(537, 3)
        Me.SplitterInstrDetail.TabIndex = 9
        Me.SplitterInstrDetail.TabStop = False
        '
        'GroupBoxInstr
        '
        Me.GroupBoxInstr.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelInstr})
        Me.GroupBoxInstr.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBoxInstr.Location = New System.Drawing.Point(1, 327)
        Me.GroupBoxInstr.Name = "GroupBoxInstr"
        Me.GroupBoxInstr.Size = New System.Drawing.Size(537, 72)
        Me.GroupBoxInstr.TabIndex = 8
        Me.GroupBoxInstr.TabStop = False
        Me.GroupBoxInstr.Text = "Option Information"
        '
        'LabelInstr
        '
        Me.LabelInstr.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelInstr.Location = New System.Drawing.Point(3, 18)
        Me.LabelInstr.Name = "LabelInstr"
        Me.LabelInstr.Size = New System.Drawing.Size(531, 51)
        Me.LabelInstr.TabIndex = 0
        '
        'PanelInstrTitle
        '
        Me.PanelInstrTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelInstrTitle, Me.btnCloseInstr})
        Me.PanelInstrTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelInstrTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelInstrTitle.Name = "PanelInstrTitle"
        Me.PanelInstrTitle.Size = New System.Drawing.Size(537, 25)
        Me.PanelInstrTitle.TabIndex = 6
        '
        'LabelInstrTitle
        '
        Me.LabelInstrTitle.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.LabelInstrTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelInstrTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelInstrTitle.Name = "LabelInstrTitle"
        Me.LabelInstrTitle.Size = New System.Drawing.Size(512, 25)
        Me.LabelInstrTitle.TabIndex = 1
        Me.LabelInstrTitle.Text = " Option Explorer"
        Me.LabelInstrTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCloseInstr
        '
        Me.btnCloseInstr.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseInstr.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseInstr.Location = New System.Drawing.Point(512, 0)
        Me.btnCloseInstr.Name = "btnCloseInstr"
        Me.btnCloseInstr.Size = New System.Drawing.Size(25, 25)
        Me.btnCloseInstr.TabIndex = 0
        Me.btnCloseInstr.Text = "X"
        '
        'SplitterInstr
        '
        Me.SplitterInstr.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.SplitterInstr.Location = New System.Drawing.Point(253, 157)
        Me.SplitterInstr.Name = "SplitterInstr"
        Me.SplitterInstr.Size = New System.Drawing.Size(539, 3)
        Me.SplitterInstr.TabIndex = 10
        Me.SplitterInstr.TabStop = False
        '
        'ImageListChild
        '
        Me.ImageListChild.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageListChild.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageListChild.ImageStream = CType(resources.GetObject("ImageListChild.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListChild.TransparentColor = System.Drawing.Color.Transparent
        '
        'PanelDisplay
        '
        Me.PanelDisplay.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControlDisplay, Me.PanelDispTitle})
        Me.PanelDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelDisplay.DockPadding.All = 1
        Me.PanelDisplay.Location = New System.Drawing.Point(253, 0)
        Me.PanelDisplay.Name = "PanelDisplay"
        Me.PanelDisplay.Size = New System.Drawing.Size(539, 157)
        Me.PanelDisplay.TabIndex = 11
        '
        'TabControlDisplay
        '
        Me.TabControlDisplay.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.TabControlDisplay.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPageQuickView, Me.TabPageSource, Me.TabPageFileList})
        Me.TabControlDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControlDisplay.ItemSize = New System.Drawing.Size(78, 21)
        Me.TabControlDisplay.Location = New System.Drawing.Point(1, 26)
        Me.TabControlDisplay.Multiline = True
        Me.TabControlDisplay.Name = "TabControlDisplay"
        Me.TabControlDisplay.SelectedIndex = 0
        Me.TabControlDisplay.ShowToolTips = True
        Me.TabControlDisplay.Size = New System.Drawing.Size(537, 130)
        Me.TabControlDisplay.TabIndex = 55
        '
        'TabPageQuickView
        '
        Me.TabPageQuickView.Controls.AddRange(New System.Windows.Forms.Control() {Me.AxWebBrowser1})
        Me.TabPageQuickView.Location = New System.Drawing.Point(4, 4)
        Me.TabPageQuickView.Name = "TabPageQuickView"
        Me.TabPageQuickView.Size = New System.Drawing.Size(529, 101)
        Me.TabPageQuickView.TabIndex = 0
        Me.TabPageQuickView.Text = "Quick View"
        '
        'AxWebBrowser1
        '
        Me.AxWebBrowser1.ContainingControl = Me
        Me.AxWebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AxWebBrowser1.Enabled = True
        Me.AxWebBrowser1.OcxState = CType(resources.GetObject("AxWebBrowser1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxWebBrowser1.Size = New System.Drawing.Size(529, 101)
        Me.AxWebBrowser1.TabIndex = 47
        '
        'TabPageSource
        '
        Me.TabPageSource.Controls.AddRange(New System.Windows.Forms.Control() {Me.RichTextBox1})
        Me.TabPageSource.Location = New System.Drawing.Point(4, 4)
        Me.TabPageSource.Name = "TabPageSource"
        Me.TabPageSource.Size = New System.Drawing.Size(529, 101)
        Me.TabPageSource.TabIndex = 1
        Me.TabPageSource.Text = "Source"
        Me.TabPageSource.Visible = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(529, 101)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.WordWrap = False
        '
        'TabPageFileList
        '
        Me.TabPageFileList.Controls.AddRange(New System.Windows.Forms.Control() {Me.TreeView1})
        Me.TabPageFileList.Location = New System.Drawing.Point(4, 4)
        Me.TabPageFileList.Name = "TabPageFileList"
        Me.TabPageFileList.Size = New System.Drawing.Size(529, 101)
        Me.TabPageFileList.TabIndex = 2
        Me.TabPageFileList.Text = "File List"
        '
        'TreeView1
        '
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeView1.FullRowSelect = True
        Me.TreeView1.ImageList = Me.ImageListChild
        Me.TreeView1.LabelEdit = True
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(529, 101)
        Me.TreeView1.TabIndex = 49
        '
        'PanelDispTitle
        '
        Me.PanelDispTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelDispTitle, Me.btnCloseDisp})
        Me.PanelDispTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelDispTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelDispTitle.Name = "PanelDispTitle"
        Me.PanelDispTitle.Size = New System.Drawing.Size(537, 25)
        Me.PanelDispTitle.TabIndex = 54
        '
        'LabelDispTitle
        '
        Me.LabelDispTitle.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.LabelDispTitle.ContextMenu = Me.ContextMenuDisp
        Me.LabelDispTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelDispTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDispTitle.Name = "LabelDispTitle"
        Me.LabelDispTitle.Size = New System.Drawing.Size(512, 25)
        Me.LabelDispTitle.TabIndex = 1
        Me.LabelDispTitle.Text = " Display Explorer"
        Me.LabelDispTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ContextMenuDisp
        '
        Me.ContextMenuDisp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuDispFloat, Me.mnuDispHide})
        '
        'mnuDispFloat
        '
        Me.mnuDispFloat.Index = 0
        Me.mnuDispFloat.Text = "Float"
        '
        'mnuDispHide
        '
        Me.mnuDispHide.Index = 1
        Me.mnuDispHide.Text = "Hide"
        '
        'btnCloseDisp
        '
        Me.btnCloseDisp.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseDisp.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseDisp.Location = New System.Drawing.Point(512, 0)
        Me.btnCloseDisp.Name = "btnCloseDisp"
        Me.btnCloseDisp.Size = New System.Drawing.Size(25, 25)
        Me.btnCloseDisp.TabIndex = 0
        Me.btnCloseDisp.Text = "X"
        '
        'ContextMenuFileNode
        '
        Me.ContextMenuFileNode.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileNodeRename})
        '
        'mnuFileNodeRename
        '
        Me.mnuFileNodeRename.Index = 0
        Me.mnuFileNodeRename.Text = "&Rename"
        '
        'FormChildBase
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 560)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDisplay, Me.SplitterInstr, Me.PanelInstr, Me.SplitterStep, Me.PanelStep})
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Menu = Me.MainMenuChild
        Me.Name = "FormChildBase"
        Me.Text = "XinCoder"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.PanelStep.ResumeLayout(False)
        Me.PanelStepDetail.ResumeLayout(False)
        Me.GroupBoxStep.ResumeLayout(False)
        Me.PanelStepTitle.ResumeLayout(False)
        Me.PanelInstr.ResumeLayout(False)
        Me.GroupBoxInstr.ResumeLayout(False)
        Me.PanelInstrTitle.ResumeLayout(False)
        Me.PanelDisplay.ResumeLayout(False)
        Me.TabControlDisplay.ResumeLayout(False)
        Me.TabPageQuickView.ResumeLayout(False)
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPageSource.ResumeLayout(False)
        Me.TabPageFileList.ResumeLayout(False)
        Me.PanelDispTitle.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    ' Note: AND - always evaluate both conditions, 
    ' AndAlso - evaluate second only if first condition is true
    ' Hide explorer toolbox(es) if any
    ' disable Alt+F4 and close button "X" on upper right corner
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Dim SC_CLOSE As Integer = 61536
        Dim SC_MINIMIZE As Integer = 61472
        Dim WM_SYSCOMMAND As Integer = 274
        Dim explorer As FormChildExplorerBase

        ' determine whether it's system information, 
        ' and what to do with close window and disable alt + f4
        If m.Msg = WM_SYSCOMMAND And m.WParam.ToInt32 = SC_MINIMIZE Then
            For Each explorer In Me.explorers
                If (Not explorer Is Nothing) AndAlso explorer.isFloating Then
                    explorer.Hide()
                End If
            Next
        ElseIf m.Msg = WM_SYSCOMMAND Then
            For Each explorer In Me.explorers
                If (Not explorer Is Nothing) AndAlso explorer.isFloating Then
                    explorer.Show()
                End If
            Next
        End If

        ' close me. There are some menu cleaning up work to do in FormMain.vb
        If m.Msg = WM_SYSCOMMAND And m.WParam.ToInt32 = SC_CLOSE Then
            CType(Me.MdiParent, FormMain).closeMdiChild(Me)
            Exit Sub
        End If

        MyBase.WndProc(m)

    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Important path used in building the files
    Protected globalHomePath As String = "../index.asp"
    Protected globalCssFilePath As String = "../xc.css"
    Protected headerFilePath As String = "../rootInclude/header.asp"
    Protected footerFilePath As String = "../rootInclude/footer.asp"
    Protected adminSecurityFilePath As String = "../login/admin_security.asp"
    Protected memSecurityFilePath As String = "../login/mem_security.asp"
    Protected globleFuncFilePath As String = "../../asp-bin/aspFuncLib.asp"
    Protected adminHomeFilePath As String = "../admin/index.asp"
    Protected memHomeFilePath As String = "../member/index.asp"
    Protected imageFolderPath As String = "../images/"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    ' Used by the default layout table.
    Protected tblAttr As New clsHTMLTableAttributes()
    Protected clsMD As New clsMetadata()
    Protected metadata As String
    Protected Friend Sub setMetadata(ByVal metadata As String, Optional ByVal tblAttr As clsHTMLTableAttributes = Nothing)
        Me.metadata = metadata
        If Not tblAttr Is Nothing Then Me.tblAttr = tblAttr
    End Sub

    Protected LabelStepTitleText As String = " Step Explorer"
    Protected LabelDispTitleText As String = " Display Explorer"
    Protected LabelInstrTitleText As String = " Instruction Explorer"

    Private labelDefaultForeColor As Color
    Private labelDefaultBackColor As Color
    Private labelHighlightForeColor As Color
    Private labelHighlightBackColor As Color

    Friend explorers As New Collection()
    Friend stepExplorer As FormChildExplorerStep
    Friend dispExplorer As FormChildExplorerDisp
    Friend instrExplorer As FormChildExplorerInstr

    Friend windowMenuItem As MenuItem 'refers to the entry on FormMain.vb's window menu

    Friend projectName As String
    Friend moduleName As String
    Friend moduleFullPath As String
    Friend moduleType As ModuleMain.xcModuleType
    Friend infoSaved As Boolean
    Protected Friend projectOutputPath As String
    Protected Friend moduleOutputPath As String

    ' Used by Preview - to tmporarily view the pages.
    Protected tmpFileFolder As String
    Protected tmpFileSubFolder As String
    Protected tmpImageFolder As String

    ' Store step labels
    Protected arrayListStepLabels As New ArrayList()
    Protected arrayListOptionTabs As New ArrayList()

    Protected stepNumber As Integer
    Protected NUM_STEPS As Integer
    Protected prevSelectedIndex As Integer ' for TabControlOptions



    Private Sub FormChildBase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initLabelColors()
        Me.infoSaved = True

        ' This will be replaced by initParam function.
        Me.moduleType = ModuleMain.xcModuleType.xcTemplateBase
        Me.prevSelectedIndex = 0
        Me.stepNumber = 1

        'Me.initModuleContent()
        ' Control tabControl tabpages.
        'Me.TabControlDisplay.SelectedTab = Me.TabPageQuickView
        'Me.TabControlDisplay.Controls.Remove(Me.TabPageSource)
    End Sub


    Private Sub initLabelColors()
        labelDefaultBackColor = Me.LabelStepTitle.BackColor
        labelDefaultForeColor = Me.LabelStepTitle.ForeColor
        labelHighlightBackColor = Color.Blue
        labelHighlightForeColor = Color.White
    End Sub


    Friend Sub initProjParam(ByRef mdiParentForm As FormMain, _
                  ByVal modType As ModuleMain.xcModuleType, _
                  ByVal modPath As String, _
                  ByVal modName As String)
        Me.MdiParent = mdiParentForm
        Me.projectName = CType(mdiParentForm, FormMain).curProjectName
        Me.moduleName = modName
        Me.moduleFullPath = modPath ' Is this needed? Can I remove this?
        Me.moduleType = modType
        Me.projectOutputPath = CType(mdiParentForm, FormMain).getProjectOutputPath()

        Me.Text = Me.Text & " - " & modName & SetupManager.MODULE_FILE_EXTENSION

        'Me.windowMenuItem = new MenuItem(me.Text )
        Me.windowMenuItem = _
             New MenuItem(mdiParentForm.MdiChildren.GetLength(0) & ". " & Me.Text)

        mdiParentForm.mnuWindow.MenuItems.Add(Me.windowMenuItem)
        AddHandler Me.windowMenuItem.Click, _
                   New System.EventHandler(AddressOf mdiParentForm.mnuWindowDynamicMenuItems_Click)
    End Sub


    Friend Sub rename(ByVal newProjName As String, ByVal newFullPath As String)
        Me.windowMenuItem.Text = MyUtil.removeEndString(Me.windowMenuItem.Text, Me.Text)
        Me.Text = MyUtil.removeEndString(Me.Text, IOManager.formatFileName(Me.moduleName, _
                                         SetupManager.MODULE_FILE_EXTENSION, True))

        Me.moduleName = newProjName
        Me.moduleFullPath = newFullPath

        Me.Text = Me.Text & Me.moduleName & SetupManager.MODULE_FILE_EXTENSION
        Me.windowMenuItem.Text = Me.windowMenuItem.Text & Me.Text
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions to set highlight color
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub setOffFocusColor(ByRef lbl As Label)
        lbl.ForeColor = Me.labelDefaultForeColor
        lbl.BackColor = Me.labelDefaultBackColor
    End Sub

    Friend Sub setOnFocusColor(ByRef lbl As Label)
        setAllOffFocusColor()
        lbl.ForeColor = Me.labelHighlightForeColor
        lbl.BackColor = Me.labelHighlightBackColor
    End Sub

    Private Sub setAllOffFocusColor()
        setOffFocusColor(Me.LabelStepTitle)
        setOffFocusColor(Me.LabelDispTitle)
        setOffFocusColor(Me.LabelInstrTitle)
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Event handlers that highlight title labels
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub LabelStepTitle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelStepTitle.Click
        setOnFocusColor(LabelStepTitle)
    End Sub

    Private Sub LabelDispTitle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelDispTitle.Click
        setOnFocusColor(LabelDispTitle)
    End Sub

    Private Sub LabelInstrTitle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelInstrTitle.Click
        setOnFocusColor(LabelInstrTitle)
    End Sub

    Private Sub TabControlDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControlDisplay.Click
        setOnFocusColor(LabelDispTitle)
    End Sub

    Private Sub PanelStepDetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanelStepDetail.Click
        setOnFocusColor(LabelStepTitle)
    End Sub

    Private Sub GroupBoxOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        setOnFocusColor(LabelInstrTitle)
    End Sub

    Private Sub GroupBoxInstr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxInstr.Click
        setOnFocusColor(LabelInstrTitle)
    End Sub

    Private Sub LabelInstr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelInstr.Click
        setOnFocusColor(LabelInstrTitle)
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions to hide panels.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Friend Sub hidePanelStep()
        Me.PanelStep.Hide()
    End Sub

    Friend Sub hidePanelDisp()
        Me.PanelDisplay.Hide()
        Me.PanelInstr.Dock = DockStyle.Fill
    End Sub

    Friend Sub hidePanelInstr()
        Me.PanelInstr.Hide()
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Interface Functions to hide panels.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub btnCloseStep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseStep.Click
        Me.hidePanelStep()
    End Sub

    Private Sub btnCloseDisp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseDisp.Click
        Me.hidePanelDisp()
    End Sub

    Private Sub btnCloseInstr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseInstr.Click
        Me.hidePanelInstr()
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions to show panels or floating windows.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Friend Sub viewStepExplorer()
        If (Not Me.stepExplorer Is Nothing) AndAlso Me.stepExplorer.isFloating Then
            Me.setAllOffFocusColor()
            Me.stepExplorer.Focus()
            Exit Sub
        End If
        Me.setOnFocusColor(Me.LabelStepTitle)
        Me.PanelStep.Show()
    End Sub

    Friend Sub viewDispExplorer()
        If (Not Me.dispExplorer Is Nothing) AndAlso Me.dispExplorer.isFloating Then
            Me.setAllOffFocusColor()
            Me.dispExplorer.Focus()
            Exit Sub
        End If
        Me.PanelInstr.Dock = DockStyle.Bottom
        Me.setOnFocusColor(Me.LabelDispTitle)
        Me.PanelDisplay.Show()
    End Sub

    Friend Sub viewInstrExplorer()
        If (Not Me.instrExplorer Is Nothing) AndAlso Me.instrExplorer.isFloating Then
            Me.setAllOffFocusColor()
            Me.instrExplorer.Focus()
            Exit Sub
        End If
        Me.setOnFocusColor(Me.LabelInstrTitle)
        Me.PanelInstr.Show()
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions to adjust explorer labels when resizing panels
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub FormChildBase_Resize(ByVal sender As Object, ByVal e As _
           System.EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then Exit Sub
        adjustLabelText(LabelStepTitle)
        adjustLabelText(LabelDispTitle)
        adjustLabelText(LabelInstrTitle)
        Me.btnBack.Top = Me.PanelStepDetail.Height - 75
        Me.btnNext.Top = Me.PanelStepDetail.Height - 75
    End Sub

    Private Sub SplitterStep_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitterStep.SplitterMoved
        adjustLabelText(LabelStepTitle)
        adjustLabelText(LabelDispTitle)
        adjustLabelText(LabelInstrTitle)
    End Sub

    Private Sub SplitterFileList_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
        adjustLabelText(LabelDispTitle)
        adjustLabelText(LabelInstrTitle)
    End Sub


    Private Sub adjustLabelText(ByRef lbl As Label)
        Dim lblWidth As Integer = lbl.Width
        Dim lblFontSize As Integer = lbl.Font.SizeInPoints
        Dim lblTextLength As Integer
        Dim tmpStr As String
        Dim adjustLen As Integer = 25   ' For minor adjustments to lblTextWidth, empirical.
        Dim tmpDots As String = "..."   ' To append to the end of shortened text.

        If lbl Is LabelStepTitle Then
            lblTextLength = LabelStepTitleText.Length
            tmpStr = LabelStepTitleText
            adjustLen = 30
        ElseIf lbl Is LabelDispTitle Then
            lblTextLength = LabelDispTitleText.Length
            tmpStr = LabelDispTitleText
        ElseIf lbl Is LabelInstrTitle Then
            lblTextLength = LabelInstrTitleText.Length
            tmpStr = LabelInstrTitleText
        End If

        Dim lblTextWidth = lblFontSize * lblTextLength

        lblTextWidth += adjustLen
        If lblWidth < lblTextWidth Then
            Dim tmpLen As Integer = (lblTextWidth - lblWidth) / lblFontSize
            If tmpLen < 0 Then
                tmpLen = 0
            ElseIf tmpLen > lblTextLength Then
                tmpLen = lblTextLength
            End If
            If lblWidth < 20 Then
                tmpDots = "."
            ElseIf lblWidth < 10 Then
                tmpDots = ""
            End If
            lbl.Text = StrReverse(StrReverse(tmpStr).Substring(tmpLen)) & tmpDots
        Else
            lbl.Text = tmpStr
        End If

    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions for float/dock explorers
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' StepExplorer
    Private Sub LabelStepTitle_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelStepTitle.DoubleClick
        floatPanelStep()
    End Sub

    Private Sub MenuStepFloat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStepFloat.Click
        floatPanelStep()
    End Sub

    Private Sub MenuStepHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStepHide.Click
        hidePanelStep()
    End Sub

    Private Sub floatPanelStep()
        Me.PanelStep.Hide()
        If stepExplorer Is Nothing Then
            stepExplorer = New FormChildExplorerStep()
            stepExplorer.init(Me)
            Me.explorers.Add(stepExplorer)
        End If
        stepExplorer.floatMe()
    End Sub

    ' DispExplorer
    Private Sub floatPanelDisp()
        Me.PanelDisplay.Hide()
        Me.PanelInstr.Dock = DockStyle.Fill
        If dispExplorer Is Nothing Then
            dispExplorer = New FormChildExplorerDisp()
            dispExplorer.init(Me)
            Me.explorers.Add(dispExplorer)
        End If
        dispExplorer.floatMe()
    End Sub

    Private Sub LabelDispTitle_doubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelDispTitle.DoubleClick
        floatPanelDisp()
    End Sub

    Private Sub mnuDispFloat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDispFloat.Click
        floatPanelDisp()
    End Sub

    Private Sub mnuDispHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDispHide.Click
        hidePanelDisp()
    End Sub


    ' InstructionExplorer
    Private Sub floatPanelInstr()
        Me.PanelInstr.Hide()
        If instrExplorer Is Nothing Then
            instrExplorer = New FormChildExplorerInstr()
            instrExplorer.init(Me)
            Me.explorers.Add(instrExplorer)
        End If
        instrExplorer.floatMe()
    End Sub

    Private Sub LabelInstrTitle_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelInstrTitle.DoubleClick
        floatPanelInstr()
    End Sub

    Private Sub mnuInstrFloat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuInstrFloat.Click
        floatPanelInstr()
    End Sub

    Private Sub mnuInstrHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuInstrHide.Click
        hidePanelInstr()
    End Sub


    ' Function to show context menu right/left upon clicking on Project Explorer
    Private Sub TreeView1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TreeView1.MouseUp
        Me.setOnFocusColor(Me.LabelDispTitle)

        ' Show menu only if Right Mouse button is clicked
        If e.Button = MouseButtons.Right Then
            'MsgBox("right click")
            ' Point where mouse is clicked
            Dim p As Point = New Point(e.X, e.Y)

            ' Go to the node that the user clicked
            Dim node As TreeNode = Me.TreeView1.GetNodeAt(p)
            If Not node Is Nothing Then
                ' Highlight the node that the user clicked.
                ' The node is highlighted until the Menu is displayed on the screen
                TreeView1.SelectedNode = node

                ' The root node is not editable, its text contains the 
                ' path information of the output files.
                If InStr(node.Text, "\") = 0 Then
                    ' only the root node is path and contains "\"
                    Me.ContextMenuFileNode.Show(TreeView1, p)
                End If
            Else
                Me.ContextMenuDisp.Show(TreeView1, p)
            End If
        End If

    End Sub

    Private Sub mnuFileNodeRename_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileNodeRename.Click
        Me.TreeView1.SelectedNode.BeginEdit()
    End Sub

    Private Sub TreeView1_BeforeExpand(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeExpand
        If Not Me.TreeView1.SelectedNode Is Nothing Then
            ' Don't change image icon if it's a file node
            If Me.TreeView1.SelectedNode.ImageIndex = 0 Then Exit Sub
            Me.TreeView1.SelectedNode.ImageIndex = 2
            Me.TreeView1.SelectedNode.SelectedImageIndex = 2
        End If
    End Sub

    Private Sub TreeView1_BeforeCollapse(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeCollapse
        If Not Me.TreeView1.SelectedNode Is Nothing Then
            ' Don't change image icon if it's a file node
            If Me.TreeView1.SelectedNode.ImageIndex = 0 Then Exit Sub
            Me.TreeView1.SelectedNode.ImageIndex = 1
            Me.TreeView1.SelectedNode.SelectedImageIndex = 1
        End If
    End Sub

    Private Sub RichTextBox1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RichTextBox1.MouseUp
        Me.setOnFocusColor(Me.LabelDispTitle)
        If e.Button = MouseButtons.Right Then
            Me.ContextMenuDisp.Show(RichTextBox1, New Point(e.X, e.Y))
        End If
    End Sub

    Private Sub TabControlOptions_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TabControlOptions.MouseUp
        Me.setOnFocusColor(Me.LabelInstrTitle)
    End Sub


    Private Sub TreeView1_AfterLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TreeView1.AfterLabelEdit
        If e.Label = "" Or e.Label = e.Node.Text Then
            e.CancelEdit = True ' Name is not changed, do nothing 
        Else
            If Not IOManager.isValidFileName(e.Label, True) Then
                e.CancelEdit = True
                e.Node.BeginEdit()
            Else    ' Only change project name and project file name, not path.
                e.CancelEdit = False
            End If
        End If
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Interface functions
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        stepNumber = stepNumber - 1
        Me.processStep()
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        stepNumber = stepNumber + 1
        Me.processStep()
    End Sub

    Protected Overridable Sub processStep()
    End Sub

    Protected Sub setStepLabelTrue(ByRef lbl As Label)
        lbl.BackColor = Color.Green
        lbl.ForeColor = Color.White
    End Sub

    Protected Sub setStepLabelFalse(ByRef lbl As Label)
        lbl.BackColor = Color.LightGray
        lbl.ForeColor = Color.Black
    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions for the tasks of the module
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Read module file
    Protected Function readModuleInfo(ByRef lines As String()) As Boolean
        Dim errInfo As String
        ' This happens the first time the module is created.
        ' The module file is not created yet so cannot read from it.
        If Not IOManager.fileExists(moduleFullPath) Then Return False

        If IOManager.getFileAsLines(moduleFullPath, lines, errInfo) = False Then
            'MsgBox("Error: cannot open module " & FilePath & vbCrLf & errInfo, MsgBoxStyle.Critical)
            MsgBox("Error: " & vbCrLf & errInfo, MsgBoxStyle.Critical)
            Return False
        End If

        Return True
    End Function


    Protected Overridable Sub initModuleContent()
    End Sub

    Protected Overridable Sub initTreeNodesText()
    End Sub

    Protected Overridable Sub constructTreeViewFileList()
    End Sub

    Protected Friend Overridable Sub buildModule( _
        Optional ByVal frameworkParams As clsWebsiteFrameworkParameters = Nothing)
        ' to be overriden by child to write module output files.
    End Sub

    Protected Friend Overridable Sub saveModule()
        Dim str As String = ""
        str = str & "CreateDate=" & Now() & vbCrLf
        str = str & "ModuleName=" & Me.moduleName & vbCrLf
        str = str & "ModuleType=" & Me.moduleType & vbCrLf
        str = str & vbCrLf
        str = str & "StepNumber=" & Me.stepNumber & vbCrLf

        IOManager.SaveTextToFile(str, Me.moduleFullPath)

        Me.infoSaved = True
    End Sub

End Class
