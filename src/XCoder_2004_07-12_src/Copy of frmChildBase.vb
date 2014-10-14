Public Class frmChildBase
    Inherits System.Windows.Forms.Form

    Private Sub Timer1_Timer(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim locPos As Point
        locPos = Me.PointToClient(Me.MousePosition)

        If Me.MouseButtons.Right Then
            Me.LabelInstrTitle.Text = "Right Click: " & locPos.X & ", " & locPos.Y
            getMousePos()
        ElseIf Me.MouseButtons.Left Then
            Me.LabelInstrTitle.Text = "Left Click: " & locPos.X & ", " & locPos.Y
            getMousePos()
        Else
            Me.LabelInstrTitle.Text = ""
        End If

    End Sub

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
    Protected WithEvents PanelStep As System.Windows.Forms.Panel
    Protected WithEvents SplitterStep As System.Windows.Forms.Splitter
    Protected WithEvents btnNext As System.Windows.Forms.Button
    Protected WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents SplitterInstr As System.Windows.Forms.Splitter
    Friend WithEvents PanelDisplay As System.Windows.Forms.Panel
    Protected WithEvents PanelFileList As System.Windows.Forms.Panel
    Protected WithEvents SplitterFileList As System.Windows.Forms.Splitter
    Protected WithEvents PanelInstr As System.Windows.Forms.Panel
    Friend WithEvents TabControlDisplay As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents AxWebBrowser1 As AxSHDocVw.AxWebBrowser
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents PanelFileDesc As System.Windows.Forms.Panel
    Friend WithEvents SplitterFileDesc As System.Windows.Forms.Splitter
    Friend WithEvents PanelFileListDetail As System.Windows.Forms.Panel
    Protected WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents TreeViewFileList As System.Windows.Forms.TreeView
    Friend WithEvents TextBoxInstr As System.Windows.Forms.TextBox
    Friend WithEvents SplitterInstrDetail As System.Windows.Forms.Splitter
    Friend WithEvents PanelDispTitle As System.Windows.Forms.Panel
    Friend WithEvents btnCloseDisp As System.Windows.Forms.Button
    Friend WithEvents LabelDispTitle As System.Windows.Forms.Label
    Friend WithEvents PanelStepTitle As System.Windows.Forms.Panel
    Friend WithEvents btnCloseStep As System.Windows.Forms.Button
    Friend WithEvents LabelStepTitle As System.Windows.Forms.Label
    Friend WithEvents PanelInstrTitle As System.Windows.Forms.Panel
    Friend WithEvents btnCloseInstr As System.Windows.Forms.Button
    Friend WithEvents LabelInstrTitle As System.Windows.Forms.Label
    Friend WithEvents PanelFileListTitle As System.Windows.Forms.Panel
    Friend WithEvents btnCloseFileList As System.Windows.Forms.Button
    Friend WithEvents LabelFileListTitle As System.Windows.Forms.Label
    Friend WithEvents PanelFileDescTitle As System.Windows.Forms.Panel
    Friend WithEvents btnCloseFileDesc As System.Windows.Forms.Button
    Friend WithEvents TextBoxFileDesc As System.Windows.Forms.TextBox
    Friend WithEvents LabelFileDescTitle As System.Windows.Forms.Label
    Friend WithEvents PanelStepDetail As System.Windows.Forms.Panel
    Friend WithEvents PanelOptions As System.Windows.Forms.Panel
    Friend WithEvents MainMenuChild As System.Windows.Forms.MainMenu
    Friend WithEvents mnuChildView As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChildViewStep As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChildViewDisp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChildViewInstr As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChildViewFileList As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChildViewFileDesc As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmChildBase))
        Me.PanelStep = New System.Windows.Forms.Panel()
        Me.PanelStepTitle = New System.Windows.Forms.Panel()
        Me.LabelStepTitle = New System.Windows.Forms.Label()
        Me.btnCloseStep = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.SplitterStep = New System.Windows.Forms.Splitter()
        Me.PanelFileList = New System.Windows.Forms.Panel()
        Me.PanelFileListDetail = New System.Windows.Forms.Panel()
        Me.PanelFileListTitle = New System.Windows.Forms.Panel()
        Me.LabelFileListTitle = New System.Windows.Forms.Label()
        Me.btnCloseFileList = New System.Windows.Forms.Button()
        Me.TreeViewFileList = New System.Windows.Forms.TreeView()
        Me.SplitterFileDesc = New System.Windows.Forms.Splitter()
        Me.PanelFileDesc = New System.Windows.Forms.Panel()
        Me.PanelFileDescTitle = New System.Windows.Forms.Panel()
        Me.btnCloseFileDesc = New System.Windows.Forms.Button()
        Me.SplitterFileList = New System.Windows.Forms.Splitter()
        Me.PanelInstr = New System.Windows.Forms.Panel()
        Me.PanelInstrTitle = New System.Windows.Forms.Panel()
        Me.LabelInstrTitle = New System.Windows.Forms.Label()
        Me.btnCloseInstr = New System.Windows.Forms.Button()
        Me.SplitterInstrDetail = New System.Windows.Forms.Splitter()
        Me.TextBoxInstr = New System.Windows.Forms.TextBox()
        Me.SplitterInstr = New System.Windows.Forms.Splitter()
        Me.PanelDisplay = New System.Windows.Forms.Panel()
        Me.PanelDispTitle = New System.Windows.Forms.Panel()
        Me.LabelDispTitle = New System.Windows.Forms.Label()
        Me.btnCloseDisp = New System.Windows.Forms.Button()
        Me.TabControlDisplay = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.AxWebBrowser1 = New AxSHDocVw.AxWebBrowser()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MainMenuChild = New System.Windows.Forms.MainMenu()
        Me.mnuChildView = New System.Windows.Forms.MenuItem()
        Me.mnuChildViewStep = New System.Windows.Forms.MenuItem()
        Me.mnuChildViewDisp = New System.Windows.Forms.MenuItem()
        Me.mnuChildViewInstr = New System.Windows.Forms.MenuItem()
        Me.TextBoxFileDesc = New System.Windows.Forms.TextBox()
        Me.LabelFileDescTitle = New System.Windows.Forms.Label()
        Me.mnuChildViewFileList = New System.Windows.Forms.MenuItem()
        Me.mnuChildViewFileDesc = New System.Windows.Forms.MenuItem()
        Me.PanelStepDetail = New System.Windows.Forms.Panel()
        Me.PanelOptions = New System.Windows.Forms.Panel()
        Me.PanelStep.SuspendLayout()
        Me.PanelStepTitle.SuspendLayout()
        Me.PanelFileList.SuspendLayout()
        Me.PanelFileListDetail.SuspendLayout()
        Me.PanelFileListTitle.SuspendLayout()
        Me.PanelFileDesc.SuspendLayout()
        Me.PanelFileDescTitle.SuspendLayout()
        Me.PanelInstr.SuspendLayout()
        Me.PanelInstrTitle.SuspendLayout()
        Me.PanelDisplay.SuspendLayout()
        Me.PanelDispTitle.SuspendLayout()
        Me.TabControlDisplay.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.PanelStepDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelStep
        '
        Me.PanelStep.BackColor = System.Drawing.SystemColors.ControlLight
        Me.PanelStep.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelStepDetail, Me.PanelStepTitle})
        Me.PanelStep.Dock = System.Windows.Forms.DockStyle.Left
        Me.PanelStep.DockPadding.All = 1
        Me.PanelStep.Name = "PanelStep"
        Me.PanelStep.Size = New System.Drawing.Size(250, 560)
        Me.PanelStep.TabIndex = 1
        '
        'PanelStepTitle
        '
        Me.PanelStepTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelStepTitle, Me.btnCloseStep})
        Me.PanelStepTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelStepTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelStepTitle.Name = "PanelStepTitle"
        Me.PanelStepTitle.Size = New System.Drawing.Size(248, 25)
        Me.PanelStepTitle.TabIndex = 5
        '
        'LabelStepTitle
        '
        Me.LabelStepTitle.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStepTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelStepTitle.Name = "LabelStepTitle"
        Me.LabelStepTitle.Size = New System.Drawing.Size(223, 25)
        Me.LabelStepTitle.TabIndex = 1
        Me.LabelStepTitle.Text = " Step Explorer"
        Me.LabelStepTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCloseStep
        '
        Me.btnCloseStep.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseStep.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseStep.Location = New System.Drawing.Point(223, 0)
        Me.btnCloseStep.Name = "btnCloseStep"
        Me.btnCloseStep.Size = New System.Drawing.Size(25, 25)
        Me.btnCloseStep.TabIndex = 0
        Me.btnCloseStep.Text = "X"
        '
        'btnNext
        '
        Me.btnNext.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.btnNext.BackColor = System.Drawing.SystemColors.Control
        Me.btnNext.Location = New System.Drawing.Point(120, 472)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(96, 32)
        Me.btnNext.TabIndex = 4
        Me.btnNext.Text = "Next"
        '
        'btnBack
        '
        Me.btnBack.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.btnBack.BackColor = System.Drawing.SystemColors.Control
        Me.btnBack.Location = New System.Drawing.Point(32, 472)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(88, 32)
        Me.btnBack.TabIndex = 3
        Me.btnBack.Text = "Back"
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
        Me.PanelFileListDetail.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelFileListTitle, Me.TreeViewFileList})
        Me.PanelFileListDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelFileListDetail.DockPadding.All = 1
        Me.PanelFileListDetail.Location = New System.Drawing.Point(1, 1)
        Me.PanelFileListDetail.Name = "PanelFileListDetail"
        Me.PanelFileListDetail.Size = New System.Drawing.Size(248, 155)
        Me.PanelFileListDetail.TabIndex = 2
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
        Me.LabelFileListTitle.Dock = System.Windows.Forms.DockStyle.Fill
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
        'TreeViewFileList
        '
        Me.TreeViewFileList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeViewFileList.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeViewFileList.FullRowSelect = True
        Me.TreeViewFileList.ImageIndex = -1
        Me.TreeViewFileList.LabelEdit = True
        Me.TreeViewFileList.Location = New System.Drawing.Point(1, 1)
        Me.TreeViewFileList.Name = "TreeViewFileList"
        Me.TreeViewFileList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TreeViewFileList.SelectedImageIndex = -1
        Me.TreeViewFileList.Size = New System.Drawing.Size(246, 153)
        Me.TreeViewFileList.TabIndex = 46
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
        Me.PanelFileDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PanelFileDesc.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBoxFileDesc, Me.PanelFileDescTitle})
        Me.PanelFileDesc.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelFileDesc.DockPadding.All = 1
        Me.PanelFileDesc.Location = New System.Drawing.Point(1, 159)
        Me.PanelFileDesc.Name = "PanelFileDesc"
        Me.PanelFileDesc.Size = New System.Drawing.Size(248, 400)
        Me.PanelFileDesc.TabIndex = 0
        '
        'PanelFileDescTitle
        '
        Me.PanelFileDescTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelFileDescTitle, Me.btnCloseFileDesc})
        Me.PanelFileDescTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelFileDescTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelFileDescTitle.Name = "PanelFileDescTitle"
        Me.PanelFileDescTitle.Size = New System.Drawing.Size(242, 25)
        Me.PanelFileDescTitle.TabIndex = 2
        '
        'btnCloseFileDesc
        '
        Me.btnCloseFileDesc.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseFileDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseFileDesc.Location = New System.Drawing.Point(217, 0)
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
        Me.PanelInstr.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PanelInstr.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelOptions, Me.PanelInstrTitle, Me.SplitterInstrDetail, Me.TextBoxInstr})
        Me.PanelInstr.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelInstr.DockPadding.All = 1
        Me.PanelInstr.Location = New System.Drawing.Point(253, 160)
        Me.PanelInstr.Name = "PanelInstr"
        Me.PanelInstr.Size = New System.Drawing.Size(286, 400)
        Me.PanelInstr.TabIndex = 9
        '
        'PanelInstrTitle
        '
        Me.PanelInstrTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelInstrTitle, Me.btnCloseInstr})
        Me.PanelInstrTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelInstrTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelInstrTitle.Name = "PanelInstrTitle"
        Me.PanelInstrTitle.Size = New System.Drawing.Size(280, 25)
        Me.PanelInstrTitle.TabIndex = 6
        '
        'LabelInstrTitle
        '
        Me.LabelInstrTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelInstrTitle.Name = "LabelInstrTitle"
        Me.LabelInstrTitle.Size = New System.Drawing.Size(255, 25)
        Me.LabelInstrTitle.TabIndex = 1
        Me.LabelInstrTitle.Text = " Instruction Explorer"
        Me.LabelInstrTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCloseInstr
        '
        Me.btnCloseInstr.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseInstr.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseInstr.Location = New System.Drawing.Point(255, 0)
        Me.btnCloseInstr.Name = "btnCloseInstr"
        Me.btnCloseInstr.Size = New System.Drawing.Size(25, 25)
        Me.btnCloseInstr.TabIndex = 0
        Me.btnCloseInstr.Text = "X"
        '
        'SplitterInstrDetail
        '
        Me.SplitterInstrDetail.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.SplitterInstrDetail.Location = New System.Drawing.Point(1, 192)
        Me.SplitterInstrDetail.Name = "SplitterInstrDetail"
        Me.SplitterInstrDetail.Size = New System.Drawing.Size(280, 3)
        Me.SplitterInstrDetail.TabIndex = 5
        Me.SplitterInstrDetail.TabStop = False
        '
        'TextBoxInstr
        '
        Me.TextBoxInstr.BackColor = System.Drawing.SystemColors.Window
        Me.TextBoxInstr.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TextBoxInstr.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TextBoxInstr.Location = New System.Drawing.Point(1, 195)
        Me.TextBoxInstr.Multiline = True
        Me.TextBoxInstr.Name = "TextBoxInstr"
        Me.TextBoxInstr.ReadOnly = True
        Me.TextBoxInstr.Size = New System.Drawing.Size(280, 200)
        Me.TextBoxInstr.TabIndex = 4
        Me.TextBoxInstr.Text = "Instruction details goes here ..."
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
        'PanelDisplay
        '
        Me.PanelDisplay.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDispTitle, Me.TabControlDisplay})
        Me.PanelDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelDisplay.DockPadding.All = 1
        Me.PanelDisplay.Location = New System.Drawing.Point(253, 0)
        Me.PanelDisplay.Name = "PanelDisplay"
        Me.PanelDisplay.Size = New System.Drawing.Size(286, 157)
        Me.PanelDisplay.TabIndex = 11
        '
        'PanelDispTitle
        '
        Me.PanelDispTitle.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelDispTitle, Me.btnCloseDisp})
        Me.PanelDispTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelDispTitle.Location = New System.Drawing.Point(1, 1)
        Me.PanelDispTitle.Name = "PanelDispTitle"
        Me.PanelDispTitle.Size = New System.Drawing.Size(284, 25)
        Me.PanelDispTitle.TabIndex = 50
        '
        'LabelDispTitle
        '
        Me.LabelDispTitle.Dock = System.Windows.Forms.DockStyle.Fill
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
        'TabControlDisplay
        '
        Me.TabControlDisplay.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.TabControlDisplay.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPage1, Me.TabPage2})
        Me.TabControlDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControlDisplay.Location = New System.Drawing.Point(1, 1)
        Me.TabControlDisplay.Multiline = True
        Me.TabControlDisplay.Name = "TabControlDisplay"
        Me.TabControlDisplay.SelectedIndex = 0
        Me.TabControlDisplay.ShowToolTips = True
        Me.TabControlDisplay.Size = New System.Drawing.Size(284, 155)
        Me.TabControlDisplay.TabIndex = 49
        '
        'TabPage1
        '
        Me.TabPage1.Controls.AddRange(New System.Windows.Forms.Control() {Me.AxWebBrowser1})
        Me.TabPage1.Location = New System.Drawing.Point(4, 4)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(276, 126)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Quick View"
        '
        'AxWebBrowser1
        '
        Me.AxWebBrowser1.ContainingControl = Me
        Me.AxWebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AxWebBrowser1.Enabled = True
        Me.AxWebBrowser1.OcxState = CType(resources.GetObject("AxWebBrowser1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxWebBrowser1.Size = New System.Drawing.Size(276, 126)
        Me.AxWebBrowser1.TabIndex = 47
        '
        'TabPage2
        '
        Me.TabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.RichTextBox1})
        Me.TabPage2.Location = New System.Drawing.Point(4, 4)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(276, 126)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Source"
        Me.TabPage2.Visible = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(276, 126)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.WordWrap = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'MainMenuChild
        '
        Me.MainMenuChild.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuChildView})
        '
        'mnuChildView
        '
        Me.mnuChildView.Index = 0
        Me.mnuChildView.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuChildViewStep, Me.mnuChildViewDisp, Me.mnuChildViewInstr, Me.mnuChildViewFileList, Me.mnuChildViewFileDesc})
        Me.mnuChildView.Text = "View"
        '
        'mnuChildViewStep
        '
        Me.mnuChildViewStep.Index = 0
        Me.mnuChildViewStep.Text = "Step Explorer"
        '
        'mnuChildViewDisp
        '
        Me.mnuChildViewDisp.Index = 1
        Me.mnuChildViewDisp.Text = "Display Explorer"
        '
        'mnuChildViewInstr
        '
        Me.mnuChildViewInstr.Index = 2
        Me.mnuChildViewInstr.Text = "Instruction Explorer"
        '
        'TextBoxFileDesc
        '
        Me.TextBoxFileDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBoxFileDesc.Location = New System.Drawing.Point(1, 26)
        Me.TextBoxFileDesc.Multiline = True
        Me.TextBoxFileDesc.Name = "TextBoxFileDesc"
        Me.TextBoxFileDesc.Size = New System.Drawing.Size(242, 369)
        Me.TextBoxFileDesc.TabIndex = 3
        Me.TextBoxFileDesc.Text = "File Description goes here ..."
        '
        'LabelFileDescTitle
        '
        Me.LabelFileDescTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelFileDescTitle.Name = "LabelFileDescTitle"
        Me.LabelFileDescTitle.Size = New System.Drawing.Size(217, 25)
        Me.LabelFileDescTitle.TabIndex = 1
        Me.LabelFileDescTitle.Text = " File Desciption Explorer"
        Me.LabelFileDescTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'mnuChildViewFileList
        '
        Me.mnuChildViewFileList.Index = 3
        Me.mnuChildViewFileList.Text = "File List Explorer"
        '
        'mnuChildViewFileDesc
        '
        Me.mnuChildViewFileDesc.Index = 4
        Me.mnuChildViewFileDesc.Text = "File Description Explorer"
        '
        'PanelStepDetail
        '
        Me.PanelStepDetail.BackColor = System.Drawing.SystemColors.Control
        Me.PanelStepDetail.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnNext, Me.btnBack})
        Me.PanelStepDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelStepDetail.Location = New System.Drawing.Point(1, 26)
        Me.PanelStepDetail.Name = "PanelStepDetail"
        Me.PanelStepDetail.Size = New System.Drawing.Size(248, 533)
        Me.PanelStepDetail.TabIndex = 6
        '
        'PanelOptions
        '
        Me.PanelOptions.BackColor = System.Drawing.SystemColors.ControlLight
        Me.PanelOptions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelOptions.Location = New System.Drawing.Point(1, 26)
        Me.PanelOptions.Name = "PanelOptions"
        Me.PanelOptions.Size = New System.Drawing.Size(280, 166)
        Me.PanelOptions.TabIndex = 7
        '
        'frmChildBase
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 560)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDisplay, Me.SplitterInstr, Me.PanelInstr, Me.SplitterFileList, Me.PanelFileList, Me.SplitterStep, Me.PanelStep})
        Me.Menu = Me.MainMenuChild
        Me.Name = "frmChildBase"
        Me.Text = "XinCoder"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.PanelStep.ResumeLayout(False)
        Me.PanelStepTitle.ResumeLayout(False)
        Me.PanelFileList.ResumeLayout(False)
        Me.PanelFileListDetail.ResumeLayout(False)
        Me.PanelFileListTitle.ResumeLayout(False)
        Me.PanelFileDesc.ResumeLayout(False)
        Me.PanelFileDescTitle.ResumeLayout(False)
        Me.PanelInstr.ResumeLayout(False)
        Me.PanelInstrTitle.ResumeLayout(False)
        Me.PanelDisplay.ResumeLayout(False)
        Me.PanelDispTitle.ResumeLayout(False)
        Me.TabControlDisplay.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.PanelStepDetail.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private labelDefault As New Label()
    Private labelHighlight As New Label()

    Private Sub frmChildBase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        labelDefault.BackColor = Me.LabelStepTitle.BackColor
        labelDefault.ForeColor = Me.LabelStepTitle.ForeColor

        labelHighlight.BackColor = Color.Blue
        labelHighlight.ForeColor = Color.White

        Timer1.Interval = 100
        'setAllOnFocusColor()
    End Sub


    ''''''''''''''''''''''''''''''''''''''
    Private Sub PanelStep_Paint(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanelStep.Click
        setAllOutOfFocusColor()
        setLabelColor(Me.LabelStepTitle, labelHighlight)
        Me.getMousePos()
    End Sub


    Private Sub PanelFileDesc_Paint(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanelFileDesc.Click
        setAllOutOfFocusColor()
        'setLabelColor(Me.LabelFileDesc, labelHighlight)
        Me.getMousePos()
    End Sub


    ''''''''''''''''''''''''''''''''''''''
    Private Sub setAllOutOfFocusColor()
        setLabelColor(Me.LabelStepTitle, labelDefault)
        setLabelColor(Me.LabelInstrTitle, labelDefault)
    End Sub

    Private Sub setAllOnFocusColor()
        setLabelColor(Me.LabelStepTitle, labelHighlight)
        setLabelColor(Me.LabelInstrTitle, labelHighlight)
    End Sub

    Private Sub setLabelColor(ByRef lbl As Label, ByRef lblColor As Label)
        lbl.ForeColor = lblColor.ForeColor
        lbl.BackColor = lblColor.BackColor
    End Sub

    Private Sub getMousePos()
        Me.TextBoxFileDesc.Text = _
            "Me: " & Me.Left & ", " & Me.Top & ", " & Me.Width & ", " & Me.Height & vbCrLf & _
            "MousePosition: " & Me.MousePosition.X & ", " & Me.MousePosition.Y
    End Sub

    '    Private Sub frmChildBase_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.MouseHover
    '       Dim LocalMousePosition As Point

    '      LocalMousePosition = Cursor.Position
    '    'Debug.WriteLine("X = " & LocalMousePosition.X & ", " & LocalMousePosition.Y)
    '     Me.TextBoxFileDesc.Text = Me.MousePosition.X & ", " & Me.MousePosition.Y
    ' End Sub

    Friend WithEvents Timer1 As System.Windows.Forms.Timer

    Private Sub btnCloseStep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.PanelStep.Hide()
    End Sub

    Private Sub btnCloseDisp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.PanelDisplay.Hide()
    End Sub

    Private Sub btnCloseInstr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.PanelInstr.Hide()
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuChildViewStep.Click
        Me.PanelStep.Show()
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuChildViewDisp.Click
        Me.PanelDisplay.Show()
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuChildViewInstr.Click
        Me.PanelInstr.Show()
    End Sub
End Class
