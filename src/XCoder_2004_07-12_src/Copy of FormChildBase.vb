Public Class FormChildBase
    Inherits System.Windows.Forms.Form

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

    Protected LabelStepTitleText As String = " Step Explorer"
    Protected LabelDispTitleText As String = " Display Explorer"
    Protected LabelInstrTitleText As String = " Instruction Explorer"
    Protected LabelFileListTitleText As String = " File List Explorer"
    Protected LabelFileDescTitleText As String = " File Description"

    Private labelDefaultForeColor As Color
    Private labelDefaultBackColor As Color
    Private labelHighlightForeColor As Color
    Private labelHighlightBackColor As Color

    Friend explorers As New Collection()
    Friend stepExplorer As FormChildExplorerStep
    Friend dispExplorer As FormChildExplorerDisp
    Friend instrExplorer As FormChildExplorerInstr
    Friend fileListExplorer As FormChildExplorerFileList
    Friend fileDescExplorer As FormChildExplorerFileDesc

    Friend windowMenuItem As MenuItem 'refers to the entry on FormMain.vb's window menu

    Friend moduleName As String
    Friend moduleFullPath As String
    Friend moduleType As ModuleMain.xcModuleType
    Friend infoSaved As Boolean

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
    Public WithEvents GroupBoxOptions As System.Windows.Forms.GroupBox
    Public WithEvents SplitterInstr As System.Windows.Forms.Splitter
    Public WithEvents PanelFileDesc As System.Windows.Forms.Panel
    Public WithEvents SplitterFileDesc As System.Windows.Forms.Splitter
    Public WithEvents PanelStepTitle As System.Windows.Forms.Panel
    Public WithEvents LabelStepTitle As System.Windows.Forms.Label
    Public WithEvents PanelInstrTitle As System.Windows.Forms.Panel
    Public WithEvents btnCloseInstr As System.Windows.Forms.Button
    Public WithEvents LabelInstrTitle As System.Windows.Forms.Label
    Public WithEvents PanelFileListTitle As System.Windows.Forms.Panel
    Public WithEvents btnCloseFileList As System.Windows.Forms.Button
    Public WithEvents LabelFileListTitle As System.Windows.Forms.Label
    Public WithEvents PanelFileDescTitle As System.Windows.Forms.Panel
    Public WithEvents btnCloseFileDesc As System.Windows.Forms.Button
    Public WithEvents LabelFileDescTitle As System.Windows.Forms.Label
    Public WithEvents PanelStepDetail As System.Windows.Forms.Panel
    Public WithEvents MainMenuChild As System.Windows.Forms.MainMenu
    Public WithEvents ImageListChild As System.Windows.Forms.ImageList
    Public WithEvents btnCloseStep As System.Windows.Forms.Button
    Public WithEvents TreeViewFileList As System.Windows.Forms.TreeView
    Public WithEvents PanelFileListDetail As System.Windows.Forms.Panel
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
    Public WithEvents btnNext As System.Windows.Forms.Button
    Public WithEvents btnBack As System.Windows.Forms.Button
    Public WithEvents SplitterFileList As System.Windows.Forms.Splitter
    Public WithEvents PanelInstr As System.Windows.Forms.Panel
    Public WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Public WithEvents PanelFileList As System.Windows.Forms.Panel
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
    Public WithEvents GroupBoxFileDesc As System.Windows.Forms.GroupBox
    Public WithEvents LabelFileDesc As System.Windows.Forms.Label
    Public WithEvents LabelStep1 As System.Windows.Forms.Label
    Public WithEvents LabelStep2 As System.Windows.Forms.Label
    Public WithEvents LabelStep3 As System.Windows.Forms.Label
    Public WithEvents PictureBoxStep As System.Windows.Forms.PictureBox
    Public WithEvents LabelStep4 As System.Windows.Forms.Label
    Friend WithEvents TabPageFileList As System.Windows.Forms.TabPage
    Public WithEvents TreeView1 As System.Windows.Forms.TreeView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormChildBase))
        Me.PanelStep = New System.Windows.Forms.Panel()
        Me.ContextMenuStep = New System.Windows.Forms.ContextMenu()
        Me.mnuStepFloat = New System.Windows.Forms.MenuItem()
        Me.mnuStepHide = New System.Windows.Forms.MenuItem()
        Me.PanelStepDetail = New System.Windows.Forms.Panel()
        Me.PictureBoxStep = New System.Windows.Forms.PictureBox()
        Me.LabelStep4 = New System.Windows.Forms.Label()
        Me.LabelStep3 = New System.Windows.Forms.Label()
        Me.LabelStep2 = New System.Windows.Forms.Label()
        Me.LabelStep1 = New System.Windows.Forms.Label()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.PanelStepTitle = New System.Windows.Forms.Panel()
        Me.btnCloseStep = New System.Windows.Forms.Button()
        Me.LabelStepTitle = New System.Windows.Forms.Label()
        Me.SplitterStep = New System.Windows.Forms.Splitter()
        Me.PanelFileList = New System.Windows.Forms.Panel()
        Me.PanelFileListDetail = New System.Windows.Forms.Panel()
        Me.ContextMenuFileList = New System.Windows.Forms.ContextMenu()
        Me.mnuFileListFloat = New System.Windows.Forms.MenuItem()
        Me.mnuFileListHide = New System.Windows.Forms.MenuItem()
        Me.TreeViewFileList = New System.Windows.Forms.TreeView()
        Me.PanelFileListTitle = New System.Windows.Forms.Panel()
        Me.LabelFileListTitle = New System.Windows.Forms.Label()
        Me.btnCloseFileList = New System.Windows.Forms.Button()
        Me.SplitterFileDesc = New System.Windows.Forms.Splitter()
        Me.PanelFileDesc = New System.Windows.Forms.Panel()
        Me.ContextMenuFileDesc = New System.Windows.Forms.ContextMenu()
        Me.mnuFileDescFloat = New System.Windows.Forms.MenuItem()
        Me.mnuFileDescHide = New System.Windows.Forms.MenuItem()
        Me.GroupBoxFileDesc = New System.Windows.Forms.GroupBox()
        Me.LabelFileDesc = New System.Windows.Forms.Label()
        Me.PanelFileDescTitle = New System.Windows.Forms.Panel()
        Me.LabelFileDescTitle = New System.Windows.Forms.Label()
        Me.btnCloseFileDesc = New System.Windows.Forms.Button()
        Me.SplitterFileList = New System.Windows.Forms.Splitter()
        Me.PanelInstr = New System.Windows.Forms.Panel()
        Me.ContextMenuInstr = New System.Windows.Forms.ContextMenu()
        Me.mnuInstrFloat = New System.Windows.Forms.MenuItem()
        Me.mnuInstrHide = New System.Windows.Forms.MenuItem()
        Me.GroupBoxOptions = New System.Windows.Forms.GroupBox()
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
        Me.ContextMenuDisp = New System.Windows.Forms.ContextMenu()
        Me.mnuDispFloat = New System.Windows.Forms.MenuItem()
        Me.mnuDispHide = New System.Windows.Forms.MenuItem()
        Me.TabControlDisplay = New System.Windows.Forms.TabControl()
        Me.TabPageSource = New System.Windows.Forms.TabPage()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.TabPageQuickView = New System.Windows.Forms.TabPage()
        Me.AxWebBrowser1 = New AxSHDocVw.AxWebBrowser()
        Me.PanelDispTitle = New System.Windows.Forms.Panel()
        Me.LabelDispTitle = New System.Windows.Forms.Label()
        Me.btnCloseDisp = New System.Windows.Forms.Button()
        Me.TabPageFileList = New System.Windows.Forms.TabPage()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.PanelStep.SuspendLayout()
        Me.PanelStepDetail.SuspendLayout()
        Me.PanelStepTitle.SuspendLayout()
        Me.PanelFileList.SuspendLayout()
        Me.PanelFileListDetail.SuspendLayout()
        Me.PanelFileListTitle.SuspendLayout()
        Me.PanelFileDesc.SuspendLayout()
        Me.GroupBoxFileDesc.SuspendLayout()
        Me.PanelFileDescTitle.SuspendLayout()
        Me.PanelInstr.SuspendLayout()
        Me.GroupBoxInstr.SuspendLayout()
        Me.PanelInstrTitle.SuspendLayout()
        Me.PanelDisplay.SuspendLayout()
        Me.TabControlDisplay.SuspendLayout()
        Me.TabPageSource.SuspendLayout()
        Me.TabPageQuickView.SuspendLayout()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelDispTitle.SuspendLayout()
        Me.TabPageFileList.SuspendLayout()
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
        Me.PanelStepDetail.Controls.AddRange(New System.Windows.Forms.Control() {Me.PictureBoxStep, Me.LabelStep4, Me.LabelStep3, Me.LabelStep2, Me.LabelStep1, Me.btnNext, Me.btnBack})
        Me.PanelStepDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelStepDetail.Location = New System.Drawing.Point(1, 26)
        Me.PanelStepDetail.Name = "PanelStepDetail"
        Me.PanelStepDetail.Size = New System.Drawing.Size(248, 533)
        Me.PanelStepDetail.TabIndex = 6
        '
        'PictureBoxStep
        '
        Me.PictureBoxStep.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBoxStep.Location = New System.Drawing.Point(24, 24)
        Me.PictureBoxStep.Name = "PictureBoxStep"
        Me.PictureBoxStep.Size = New System.Drawing.Size(200, 130)
        Me.PictureBoxStep.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBoxStep.TabIndex = 9
        Me.PictureBoxStep.TabStop = False
        '
        'LabelStep4
        '
        Me.LabelStep4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep4.Location = New System.Drawing.Point(25, 296)
        Me.LabelStep4.Name = "LabelStep4"
        Me.LabelStep4.Size = New System.Drawing.Size(200, 32)
        Me.LabelStep4.TabIndex = 8
        Me.LabelStep4.Text = "Step 4. Upload files"
        Me.LabelStep4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep3
        '
        Me.LabelStep3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep3.Location = New System.Drawing.Point(25, 264)
        Me.LabelStep3.Name = "LabelStep3"
        Me.LabelStep3.Size = New System.Drawing.Size(200, 32)
        Me.LabelStep3.TabIndex = 7
        Me.LabelStep3.Text = "Step 3. Generate files"
        Me.LabelStep3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep2
        '
        Me.LabelStep2.BackColor = System.Drawing.SystemColors.ControlLight
        Me.LabelStep2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep2.Location = New System.Drawing.Point(25, 232)
        Me.LabelStep2.Name = "LabelStep2"
        Me.LabelStep2.Size = New System.Drawing.Size(200, 32)
        Me.LabelStep2.TabIndex = 6
        Me.LabelStep2.Text = "Step 2. Set file names"
        Me.LabelStep2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelStep1
        '
        Me.LabelStep1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.LabelStep1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelStep1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStep1.Location = New System.Drawing.Point(25, 200)
        Me.LabelStep1.Name = "LabelStep1"
        Me.LabelStep1.Size = New System.Drawing.Size(200, 32)
        Me.LabelStep1.TabIndex = 5
        Me.LabelStep1.Text = "Step 1. Set options"
        Me.LabelStep1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnNext
        '
        Me.btnNext.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.btnNext.BackColor = System.Drawing.SystemColors.Control
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
        'PanelFileList
        '
        Me.PanelFileList.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelFileListDetail, Me.SplitterFileDesc, Me.PanelFileDesc})
        Me.PanelFileList.Dock = System.Windows.Forms.DockStyle.Right
        Me.PanelFileList.DockPadding.All = 1
        Me.PanelFileList.Location = New System.Drawing.Point(542, 0)
        Me.PanelFileList.Name = "PanelFileList"
        Me.PanelFileList.Size = New System.Drawing.Size(250, 560)
        Me.PanelFileList.TabIndex = 7
        '
        'PanelFileListDetail
        '
        Me.PanelFileListDetail.ContextMenu = Me.ContextMenuFileList
        Me.PanelFileListDetail.Controls.AddRange(New System.Windows.Forms.Control() {Me.TreeViewFileList, Me.PanelFileListTitle})
        Me.PanelFileListDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelFileListDetail.DockPadding.All = 1
        Me.PanelFileListDetail.Location = New System.Drawing.Point(1, 1)
        Me.PanelFileListDetail.Name = "PanelFileListDetail"
        Me.PanelFileListDetail.Size = New System.Drawing.Size(248, 155)
        Me.PanelFileListDetail.TabIndex = 2
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
        'TreeViewFileList
        '
        Me.TreeViewFileList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeViewFileList.ImageIndex = -1
        Me.TreeViewFileList.Location = New System.Drawing.Point(1, 26)
        Me.TreeViewFileList.Name = "TreeViewFileList"
        Me.TreeViewFileList.SelectedImageIndex = -1
        Me.TreeViewFileList.Size = New System.Drawing.Size(246, 128)
        Me.TreeViewFileList.TabIndex = 48
        '
        'PanelFileListTitle
        '
        Me.PanelFileListTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelFileListTitle, Me.btnCloseFileList})
        Me.PanelFileListTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelFileListTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelFileListTitle.Name = "PanelFileListTitle"
        Me.PanelFileListTitle.Size = New System.Drawing.Size(246, 25)
        Me.PanelFileListTitle.TabIndex = 47
        '
        'LabelFileListTitle
        '
        Me.LabelFileListTitle.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.LabelFileListTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelFileListTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFileListTitle.Name = "LabelFileListTitle"
        Me.LabelFileListTitle.Size = New System.Drawing.Size(221, 25)
        Me.LabelFileListTitle.TabIndex = 1
        Me.LabelFileListTitle.Text = " File List Explorer"
        Me.LabelFileListTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCloseFileList
        '
        Me.btnCloseFileList.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseFileList.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseFileList.Location = New System.Drawing.Point(221, 0)
        Me.btnCloseFileList.Name = "btnCloseFileList"
        Me.btnCloseFileList.Size = New System.Drawing.Size(25, 25)
        Me.btnCloseFileList.TabIndex = 0
        Me.btnCloseFileList.Text = "X"
        '
        'SplitterFileDesc
        '
        Me.SplitterFileDesc.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.SplitterFileDesc.Location = New System.Drawing.Point(1, 156)
        Me.SplitterFileDesc.Name = "SplitterFileDesc"
        Me.SplitterFileDesc.Size = New System.Drawing.Size(248, 3)
        Me.SplitterFileDesc.TabIndex = 1
        Me.SplitterFileDesc.TabStop = False
        '
        'PanelFileDesc
        '
        Me.PanelFileDesc.ContextMenu = Me.ContextMenuFileDesc
        Me.PanelFileDesc.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBoxFileDesc, Me.PanelFileDescTitle})
        Me.PanelFileDesc.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelFileDesc.DockPadding.All = 1
        Me.PanelFileDesc.Location = New System.Drawing.Point(1, 159)
        Me.PanelFileDesc.Name = "PanelFileDesc"
        Me.PanelFileDesc.Size = New System.Drawing.Size(248, 400)
        Me.PanelFileDesc.TabIndex = 0
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
        'GroupBoxFileDesc
        '
        Me.GroupBoxFileDesc.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelFileDesc})
        Me.GroupBoxFileDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBoxFileDesc.Location = New System.Drawing.Point(1, 26)
        Me.GroupBoxFileDesc.Name = "GroupBoxFileDesc"
        Me.GroupBoxFileDesc.Size = New System.Drawing.Size(246, 373)
        Me.GroupBoxFileDesc.TabIndex = 3
        Me.GroupBoxFileDesc.TabStop = False
        '
        'LabelFileDesc
        '
        Me.LabelFileDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelFileDesc.Location = New System.Drawing.Point(3, 18)
        Me.LabelFileDesc.Name = "LabelFileDesc"
        Me.LabelFileDesc.Size = New System.Drawing.Size(240, 352)
        Me.LabelFileDesc.TabIndex = 0
        '
        'PanelFileDescTitle
        '
        Me.PanelFileDescTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelFileDescTitle, Me.btnCloseFileDesc})
        Me.PanelFileDescTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelFileDescTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelFileDescTitle.Name = "PanelFileDescTitle"
        Me.PanelFileDescTitle.Size = New System.Drawing.Size(246, 25)
        Me.PanelFileDescTitle.TabIndex = 2
        '
        'LabelFileDescTitle
        '
        Me.LabelFileDescTitle.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.LabelFileDescTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelFileDescTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFileDescTitle.Name = "LabelFileDescTitle"
        Me.LabelFileDescTitle.Size = New System.Drawing.Size(221, 25)
        Me.LabelFileDescTitle.TabIndex = 1
        Me.LabelFileDescTitle.Text = " File Desciption"
        Me.LabelFileDescTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCloseFileDesc
        '
        Me.btnCloseFileDesc.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseFileDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseFileDesc.Location = New System.Drawing.Point(221, 0)
        Me.btnCloseFileDesc.Name = "btnCloseFileDesc"
        Me.btnCloseFileDesc.Size = New System.Drawing.Size(25, 25)
        Me.btnCloseFileDesc.TabIndex = 0
        Me.btnCloseFileDesc.Text = "X"
        '
        'SplitterFileList
        '
        Me.SplitterFileList.Dock = System.Windows.Forms.DockStyle.Right
        Me.SplitterFileList.Location = New System.Drawing.Point(539, 0)
        Me.SplitterFileList.Name = "SplitterFileList"
        Me.SplitterFileList.Size = New System.Drawing.Size(3, 560)
        Me.SplitterFileList.TabIndex = 8
        Me.SplitterFileList.TabStop = False
        '
        'PanelInstr
        '
        Me.PanelInstr.ContextMenu = Me.ContextMenuInstr
        Me.PanelInstr.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBoxOptions, Me.SplitterInstrDetail, Me.GroupBoxInstr, Me.PanelInstrTitle})
        Me.PanelInstr.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelInstr.DockPadding.All = 1
        Me.PanelInstr.Location = New System.Drawing.Point(253, 160)
        Me.PanelInstr.Name = "PanelInstr"
        Me.PanelInstr.Size = New System.Drawing.Size(286, 400)
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
        'GroupBoxOptions
        '
        Me.GroupBoxOptions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBoxOptions.Location = New System.Drawing.Point(1, 26)
        Me.GroupBoxOptions.Name = "GroupBoxOptions"
        Me.GroupBoxOptions.Size = New System.Drawing.Size(284, 220)
        Me.GroupBoxOptions.TabIndex = 10
        Me.GroupBoxOptions.TabStop = False
        Me.GroupBoxOptions.Text = "Options"
        '
        'SplitterInstrDetail
        '
        Me.SplitterInstrDetail.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.SplitterInstrDetail.Location = New System.Drawing.Point(1, 246)
        Me.SplitterInstrDetail.Name = "SplitterInstrDetail"
        Me.SplitterInstrDetail.Size = New System.Drawing.Size(284, 3)
        Me.SplitterInstrDetail.TabIndex = 9
        Me.SplitterInstrDetail.TabStop = False
        '
        'GroupBoxInstr
        '
        Me.GroupBoxInstr.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelInstr})
        Me.GroupBoxInstr.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBoxInstr.Location = New System.Drawing.Point(1, 249)
        Me.GroupBoxInstr.Name = "GroupBoxInstr"
        Me.GroupBoxInstr.Size = New System.Drawing.Size(284, 150)
        Me.GroupBoxInstr.TabIndex = 8
        Me.GroupBoxInstr.TabStop = False
        Me.GroupBoxInstr.Text = "Instruction"
        '
        'LabelInstr
        '
        Me.LabelInstr.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelInstr.Location = New System.Drawing.Point(3, 18)
        Me.LabelInstr.Name = "LabelInstr"
        Me.LabelInstr.Size = New System.Drawing.Size(278, 129)
        Me.LabelInstr.TabIndex = 0
        '
        'PanelInstrTitle
        '
        Me.PanelInstrTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelInstrTitle, Me.btnCloseInstr})
        Me.PanelInstrTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelInstrTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelInstrTitle.Name = "PanelInstrTitle"
        Me.PanelInstrTitle.Size = New System.Drawing.Size(284, 25)
        Me.PanelInstrTitle.TabIndex = 6
        '
        'LabelInstrTitle
        '
        Me.LabelInstrTitle.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.LabelInstrTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelInstrTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelInstrTitle.Name = "LabelInstrTitle"
        Me.LabelInstrTitle.Size = New System.Drawing.Size(259, 25)
        Me.LabelInstrTitle.TabIndex = 1
        Me.LabelInstrTitle.Text = " Instruction Explorer"
        Me.LabelInstrTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCloseInstr
        '
        Me.btnCloseInstr.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseInstr.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseInstr.Location = New System.Drawing.Point(259, 0)
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
        Me.SplitterInstr.Size = New System.Drawing.Size(286, 3)
        Me.SplitterInstr.TabIndex = 10
        Me.SplitterInstr.TabStop = False
        '
        'ImageListChild
        '
        Me.ImageListChild.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageListChild.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageListChild.TransparentColor = System.Drawing.Color.Transparent
        '
        'PanelDisplay
        '
        Me.PanelDisplay.ContextMenu = Me.ContextMenuDisp
        Me.PanelDisplay.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControlDisplay, Me.PanelDispTitle})
        Me.PanelDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelDisplay.DockPadding.All = 1
        Me.PanelDisplay.Location = New System.Drawing.Point(253, 0)
        Me.PanelDisplay.Name = "PanelDisplay"
        Me.PanelDisplay.Size = New System.Drawing.Size(286, 157)
        Me.PanelDisplay.TabIndex = 11
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
        'TabControlDisplay
        '
        Me.TabControlDisplay.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.TabControlDisplay.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPageSource, Me.TabPageQuickView, Me.TabPageFileList})
        Me.TabControlDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControlDisplay.ItemSize = New System.Drawing.Size(78, 21)
        Me.TabControlDisplay.Location = New System.Drawing.Point(1, 26)
        Me.TabControlDisplay.Multiline = True
        Me.TabControlDisplay.Name = "TabControlDisplay"
        Me.TabControlDisplay.SelectedIndex = 0
        Me.TabControlDisplay.ShowToolTips = True
        Me.TabControlDisplay.Size = New System.Drawing.Size(284, 130)
        Me.TabControlDisplay.TabIndex = 55
        '
        'TabPageSource
        '
        Me.TabPageSource.Controls.AddRange(New System.Windows.Forms.Control() {Me.RichTextBox1})
        Me.TabPageSource.Location = New System.Drawing.Point(4, 4)
        Me.TabPageSource.Name = "TabPageSource"
        Me.TabPageSource.Size = New System.Drawing.Size(276, 101)
        Me.TabPageSource.TabIndex = 1
        Me.TabPageSource.Text = "Source"
        Me.TabPageSource.Visible = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(276, 101)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.WordWrap = False
        '
        'TabPageQuickView
        '
        Me.TabPageQuickView.Controls.AddRange(New System.Windows.Forms.Control() {Me.AxWebBrowser1})
        Me.TabPageQuickView.Location = New System.Drawing.Point(4, 4)
        Me.TabPageQuickView.Name = "TabPageQuickView"
        Me.TabPageQuickView.Size = New System.Drawing.Size(276, 101)
        Me.TabPageQuickView.TabIndex = 0
        Me.TabPageQuickView.Text = "Quick View"
        '
        'AxWebBrowser1
        '
        Me.AxWebBrowser1.ContainingControl = Me
        Me.AxWebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AxWebBrowser1.Enabled = True
        Me.AxWebBrowser1.OcxState = CType(resources.GetObject("AxWebBrowser1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxWebBrowser1.Size = New System.Drawing.Size(276, 101)
        Me.AxWebBrowser1.TabIndex = 47
        '
        'PanelDispTitle
        '
        Me.PanelDispTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelDispTitle, Me.btnCloseDisp})
        Me.PanelDispTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelDispTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelDispTitle.Name = "PanelDispTitle"
        Me.PanelDispTitle.Size = New System.Drawing.Size(284, 25)
        Me.PanelDispTitle.TabIndex = 54
        '
        'LabelDispTitle
        '
        Me.LabelDispTitle.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.LabelDispTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelDispTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDispTitle.Name = "LabelDispTitle"
        Me.LabelDispTitle.Size = New System.Drawing.Size(259, 25)
        Me.LabelDispTitle.TabIndex = 1
        Me.LabelDispTitle.Text = " Display Explorer"
        Me.LabelDispTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCloseDisp
        '
        Me.btnCloseDisp.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseDisp.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseDisp.Location = New System.Drawing.Point(259, 0)
        Me.btnCloseDisp.Name = "btnCloseDisp"
        Me.btnCloseDisp.Size = New System.Drawing.Size(25, 25)
        Me.btnCloseDisp.TabIndex = 0
        Me.btnCloseDisp.Text = "X"
        '
        'TabPageFileList
        '
        Me.TabPageFileList.Controls.AddRange(New System.Windows.Forms.Control() {Me.TreeView1})
        Me.TabPageFileList.Location = New System.Drawing.Point(4, 4)
        Me.TabPageFileList.Name = "TabPageFileList"
        Me.TabPageFileList.Size = New System.Drawing.Size(276, 101)
        Me.TabPageFileList.TabIndex = 2
        Me.TabPageFileList.Text = "File List"
        '
        'TreeView1
        '
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeView1.ImageIndex = -1
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.SelectedImageIndex = -1
        Me.TreeView1.Size = New System.Drawing.Size(276, 101)
        Me.TreeView1.TabIndex = 49
        '
        'FormChildBase
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 560)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDisplay, Me.SplitterInstr, Me.PanelInstr, Me.SplitterFileList, Me.PanelFileList, Me.SplitterStep, Me.PanelStep})
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Menu = Me.MainMenuChild
        Me.Name = "FormChildBase"
        Me.Text = "XinCoder"
        Me.PanelStep.ResumeLayout(False)
        Me.PanelStepDetail.ResumeLayout(False)
        Me.PanelStepTitle.ResumeLayout(False)
        Me.PanelFileList.ResumeLayout(False)
        Me.PanelFileListDetail.ResumeLayout(False)
        Me.PanelFileListTitle.ResumeLayout(False)
        Me.PanelFileDesc.ResumeLayout(False)
        Me.GroupBoxFileDesc.ResumeLayout(False)
        Me.PanelFileDescTitle.ResumeLayout(False)
        Me.PanelInstr.ResumeLayout(False)
        Me.GroupBoxInstr.ResumeLayout(False)
        Me.PanelInstrTitle.ResumeLayout(False)
        Me.PanelDisplay.ResumeLayout(False)
        Me.TabControlDisplay.ResumeLayout(False)
        Me.TabPageSource.ResumeLayout(False)
        Me.TabPageQuickView.ResumeLayout(False)
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelDispTitle.ResumeLayout(False)
        Me.TabPageFileList.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub FormChildBase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initLabelColors()
        Me.infoSaved = True

        ' This will be replaced by initParam function.
        Me.moduleType = ModuleMain.xcModuleType.xcTemplateBase

        'Me.LabelStep1.BackColor = Color.Blue
        'Me.LabelStep1.ForeColor = Color.White

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
        Me.moduleName = modName
        Me.moduleFullPath = modPath ' Is this needed? Can I remove this?
        Me.moduleType = modType

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
        setOffFocusColor(Me.LabelFileListTitle)
        setOffFocusColor(Me.LabelFileDescTitle)
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

    Private Sub LabelFileListTitle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelFileListTitle.Click
        setOnFocusColor(LabelFileListTitle)
    End Sub

    Private Sub LabelFileDescTitle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelFileDescTitle.Click
        setOnFocusColor(LabelFileDescTitle)
    End Sub

    Private Sub TabControlDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControlDisplay.Click
        setOnFocusColor(LabelDispTitle)
    End Sub

    Private Sub PanelStepDetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanelStepDetail.Click
        setOnFocusColor(LabelStepTitle)
    End Sub

    Private Sub GroupBoxOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxOptions.Click
        setOnFocusColor(LabelInstrTitle)
    End Sub

    Private Sub GroupBoxInstr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxInstr.Click
        setOnFocusColor(LabelInstrTitle)
    End Sub

    Private Sub LabelInstr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelInstr.Click
        setOnFocusColor(LabelInstrTitle)
    End Sub

    Private Sub TreeViewFileList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TreeViewFileList.Click
        setOnFocusColor(LabelFileListTitle)
    End Sub

    Private Sub LabelFileDesc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelFileDesc.Click
        setOnFocusColor(LabelFileDescTitle)
    End Sub

    Private Sub GroupBoxFileDesc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxFileDesc.Click
        setOnFocusColor(LabelFileDescTitle)
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

    Friend Sub hidePanelFileList()
        Me.PanelFileListDetail.Hide()
        Me.PanelFileDesc.Dock = DockStyle.Fill
    End Sub

    Friend Sub hidePanelFileDesc()
        Me.PanelFileDesc.Hide()
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

    Private Sub btnCloseFileList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseFileList.Click
        Me.hidePanelFileList()
    End Sub

    Private Sub btnCloseFileDesc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseFileDesc.Click
        Me.hidePanelFileDesc()
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

    Friend Sub viewFileListExplorer()
        If (Not Me.fileListExplorer Is Nothing) AndAlso Me.fileListExplorer.isFloating Then
            Me.setAllOffFocusColor()
            Me.fileListExplorer.Focus()
            Exit Sub
        End If
        Me.PanelFileDesc.Dock = DockStyle.Bottom
        Me.setOnFocusColor(Me.LabelFileListTitle)
        Me.PanelFileListDetail.Show()
    End Sub

    Friend Sub viewFileDescExplorer()
        If (Not Me.fileDescExplorer Is Nothing) AndAlso Me.fileDescExplorer.isFloating Then
            Me.setAllOffFocusColor()
            Me.fileDescExplorer.Focus()
            Exit Sub
        End If
        Me.setOnFocusColor(Me.LabelFileDescTitle)
        Me.PanelFileDesc.Show()
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
        adjustLabelText(LabelFileListTitle)
        adjustLabelText(LabelFileDescTitle)
    End Sub

    Private Sub SplitterStep_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitterStep.SplitterMoved
        adjustLabelText(LabelStepTitle)
        adjustLabelText(LabelDispTitle)
        adjustLabelText(LabelInstrTitle)
    End Sub


    Private Sub SplitterFileList_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitterFileList.SplitterMoved
        adjustLabelText(LabelDispTitle)
        adjustLabelText(LabelInstrTitle)
        adjustLabelText(LabelFileListTitle)
        adjustLabelText(LabelFileDescTitle)
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
        ElseIf lbl Is LabelFileListTitle Then
            lblTextLength = LabelFileListTitleText.Length
            tmpStr = LabelFileListTitleText
        ElseIf lbl Is LabelFileDescTitle Then
            lblTextLength = LabelFileDescTitleText.Length
            tmpStr = LabelFileDescTitleText
        End If

        Dim lblTextWidth = lblFontSize * lblTextLength

        '        Me.TextBoxFileDesc.Text = "label width = " & lblWidth & vbCrLf & _
        '                                  "label text size = " & lblFontSize & vbCrLf & _
        '                                  "label text length = " & lblTextLength & vbCrLf & _
        '                                  "label total width = " & lblTextWidth

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

    ' FileListExplorer
    Private Sub floatPanelFileList()
        Me.PanelFileListDetail.Hide()
        Me.PanelFileDesc.Dock = DockStyle.Fill
        If fileListExplorer Is Nothing Then
            fileListExplorer = New FormChildExplorerFileList()
            fileListExplorer.init(Me)
            Me.explorers.Add(fileListExplorer)
        End If
        fileListExplorer.floatMe()
    End Sub

    Private Sub LabelFileListTitle_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelFileListTitle.DoubleClick
        floatPanelFileList()
    End Sub

    Private Sub mnuFileListFloat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileListFloat.Click
        floatPanelFileList()
    End Sub

    Private Sub mnuFileListHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileListHide.Click
        hidePanelFileList()
    End Sub

    ' FileDescriptionExplorer

    Private Sub floatPanelFileDesc()
        Me.PanelFileDesc.Hide()
        If fileDescExplorer Is Nothing Then
            fileDescExplorer = New FormChildExplorerFileDesc()
            fileDescExplorer.init(Me)
            Me.explorers.Add(fileDescExplorer)
        End If
        fileDescExplorer.floatMe()
    End Sub

    Private Sub LabelFileDescTitle_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelFileDescTitle.DoubleClick
        floatPanelFileDesc()
    End Sub

    Private Sub mnuFileDescFloat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileDescFloat.Click
        floatPanelFileDesc()
    End Sub

    Private Sub mnuFileDescHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileDescHide.Click
        hidePanelFileDesc()
    End Sub

    Protected Friend Overridable Sub saveModule()
        Dim str As String = ""
        str = str & "CreateDate=" & Now() & vbCrLf
        str = str & "ModuleName=" & Me.moduleName & vbCrLf
        str = str & "ModuleType=" & Me.moduleType & vbCrLf
        str = str & vbCrLf

        IOManager.SaveTextToFile(str, Me.moduleFullPath)

        Me.infoSaved = True
    End Sub

End Class
