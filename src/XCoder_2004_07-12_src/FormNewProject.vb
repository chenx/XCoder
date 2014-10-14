Public Class FormNewProject
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
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnHelp As System.Windows.Forms.Button
    Friend WithEvents ImageListChooseNew As System.Windows.Forms.ImageList
    Friend WithEvents ListViewChooseNew As System.Windows.Forms.ListView
    Friend WithEvents ToolBarButtonLargeIcon As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonSmallIcon As System.Windows.Forms.ToolBarButton
    Friend WithEvents ImageListViewStyle As System.Windows.Forms.ImageList
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ToolBarViewStyle As System.Windows.Forms.ToolBar
    Friend WithEvents LabelFileName As System.Windows.Forms.Label
    Friend WithEvents LabelFileLoc As System.Windows.Forms.Label
    Friend WithEvents btnBrowseFolder As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBoxProjName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxProjLoc As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ListViewItem1 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "General Website", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 0)
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormNewProject))
        Me.ListViewChooseNew = New System.Windows.Forms.ListView()
        Me.ImageListChooseNew = New System.Windows.Forms.ImageList(Me.components)
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnHelp = New System.Windows.Forms.Button()
        Me.ToolBarViewStyle = New System.Windows.Forms.ToolBar()
        Me.ToolBarButtonLargeIcon = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButtonSmallIcon = New System.Windows.Forms.ToolBarButton()
        Me.ImageListViewStyle = New System.Windows.Forms.ImageList(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LabelFileName = New System.Windows.Forms.Label()
        Me.LabelFileLoc = New System.Windows.Forms.Label()
        Me.TextBoxProjName = New System.Windows.Forms.TextBox()
        Me.btnBrowseFolder = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TextBoxProjLoc = New System.Windows.Forms.TextBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'ListViewChooseNew
        '
        Me.ListViewChooseNew.BackColor = System.Drawing.SystemColors.Window
        ListViewItem1.Tag = "General Information Management System (IMS) Website Project"
        Me.ListViewChooseNew.Items.AddRange(New System.Windows.Forms.ListViewItem() {ListViewItem1})
        Me.ListViewChooseNew.LargeImageList = Me.ImageListChooseNew
        Me.ListViewChooseNew.Location = New System.Drawing.Point(16, 40)
        Me.ListViewChooseNew.MultiSelect = False
        Me.ListViewChooseNew.Name = "ListViewChooseNew"
        Me.ListViewChooseNew.Size = New System.Drawing.Size(480, 192)
        Me.ListViewChooseNew.SmallImageList = Me.ImageListChooseNew
        Me.ListViewChooseNew.TabIndex = 0
        '
        'ImageListChooseNew
        '
        Me.ImageListChooseNew.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageListChooseNew.ImageSize = New System.Drawing.Size(32, 32)
        Me.ImageListChooseNew.ImageStream = CType(resources.GetObject("ImageListChooseNew.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListChooseNew.TransparentColor = System.Drawing.Color.Transparent
        '
        'btnOK
        '
        Me.btnOK.Enabled = False
        Me.btnOK.Location = New System.Drawing.Point(192, 328)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(96, 28)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(296, 328)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(96, 28)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        '
        'btnHelp
        '
        Me.btnHelp.Location = New System.Drawing.Point(400, 328)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.Size = New System.Drawing.Size(96, 28)
        Me.btnHelp.TabIndex = 3
        Me.btnHelp.Text = "Help"
        '
        'ToolBarViewStyle
        '
        Me.ToolBarViewStyle.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButtonLargeIcon, Me.ToolBarButtonSmallIcon})
        Me.ToolBarViewStyle.ButtonSize = New System.Drawing.Size(24, 24)
        Me.ToolBarViewStyle.Divider = False
        Me.ToolBarViewStyle.Dock = System.Windows.Forms.DockStyle.None
        Me.ToolBarViewStyle.DropDownArrows = True
        Me.ToolBarViewStyle.ImageList = Me.ImageListViewStyle
        Me.ToolBarViewStyle.Location = New System.Drawing.Point(448, 8)
        Me.ToolBarViewStyle.Name = "ToolBarViewStyle"
        Me.ToolBarViewStyle.ShowToolTips = True
        Me.ToolBarViewStyle.Size = New System.Drawing.Size(48, 25)
        Me.ToolBarViewStyle.TabIndex = 8
        '
        'ToolBarButtonLargeIcon
        '
        Me.ToolBarButtonLargeIcon.ImageIndex = 0
        Me.ToolBarButtonLargeIcon.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.ToolBarButtonLargeIcon.ToolTipText = "Large Icon"
        '
        'ToolBarButtonSmallIcon
        '
        Me.ToolBarButtonSmallIcon.ImageIndex = 1
        Me.ToolBarButtonSmallIcon.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.ToolBarButtonSmallIcon.ToolTipText = "Small Icon"
        '
        'ImageListViewStyle
        '
        Me.ImageListViewStyle.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageListViewStyle.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageListViewStyle.ImageStream = CType(resources.GetObject("ImageListViewStyle.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListViewStyle.TransparentColor = System.Drawing.Color.Transparent
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(400, 24)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Select the type of project you will create:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelFileName
        '
        Me.LabelFileName.Location = New System.Drawing.Point(16, 240)
        Me.LabelFileName.Name = "LabelFileName"
        Me.LabelFileName.Size = New System.Drawing.Size(72, 23)
        Me.LabelFileName.TabIndex = 10
        Me.LabelFileName.Text = "Name:"
        Me.LabelFileName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelFileLoc
        '
        Me.LabelFileLoc.Location = New System.Drawing.Point(16, 272)
        Me.LabelFileLoc.Name = "LabelFileLoc"
        Me.LabelFileLoc.Size = New System.Drawing.Size(72, 23)
        Me.LabelFileLoc.TabIndex = 11
        Me.LabelFileLoc.Text = "Location:"
        Me.LabelFileLoc.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBoxProjName
        '
        Me.TextBoxProjName.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxProjName.HideSelection = False
        Me.TextBoxProjName.Location = New System.Drawing.Point(80, 240)
        Me.TextBoxProjName.Name = "TextBoxProjName"
        Me.TextBoxProjName.Size = New System.Drawing.Size(416, 22)
        Me.TextBoxProjName.TabIndex = 12
        Me.TextBoxProjName.Text = ""
        '
        'btnBrowseFolder
        '
        Me.btnBrowseFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowseFolder.Location = New System.Drawing.Point(424, 272)
        Me.btnBrowseFolder.Name = "btnBrowseFolder"
        Me.btnBrowseFolder.Size = New System.Drawing.Size(72, 24)
        Me.btnBrowseFolder.TabIndex = 14
        Me.btnBrowseFolder.Text = "Browse..."
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(16, 312)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(480, 3)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        '
        'TextBoxProjLoc
        '
        Me.TextBoxProjLoc.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxProjLoc.HideSelection = False
        Me.TextBoxProjLoc.Location = New System.Drawing.Point(80, 272)
        Me.TextBoxProjLoc.Name = "TextBoxProjLoc"
        Me.TextBoxProjLoc.ReadOnly = True
        Me.TextBoxProjLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TextBoxProjLoc.Size = New System.Drawing.Size(335, 22)
        Me.TextBoxProjLoc.TabIndex = 17
        Me.TextBoxProjLoc.Text = ""
        '
        'CheckBox1
        '
        Me.CheckBox1.Location = New System.Drawing.Point(16, 328)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(128, 24)
        Me.CheckBox1.TabIndex = 18
        Me.CheckBox1.Text = "Show me at start"
        '
        'FormNewProject
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(514, 370)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.ListViewChooseNew, Me.CheckBox1, Me.TextBoxProjLoc, Me.GroupBox1, Me.btnBrowseFolder, Me.TextBoxProjName, Me.LabelFileLoc, Me.LabelFileName, Me.Label1, Me.ToolBarViewStyle, Me.btnHelp, Me.btnCancel, Me.btnOK})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormNewProject"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "New Project"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim projName As String
    Dim projLocBase As String = SetupManager.OUTPUT_PATH


    Private Sub FormChooseNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ListViewChooseNew.Items.Item(0).Selected = True
        Me.ListViewChooseNew.Items.Item(0).Focused = True

        Me.ToolBarButtonLargeIcon.Pushed = True
        Me.ToolBarButtonSmallIcon.Pushed = False
        Me.ListViewChooseNew.View = View.LargeIcon

        Me.projName = ""
        Me.TextBoxProjName.Text = ""
        Me.initProjOutputPath()

        Me.TextBoxProjName.Select()
    End Sub


    Private Sub initProjOutputPath()
        Me.TextBoxProjLoc.Text = Me.getProjOutputPath()

        Me.TextBoxProjLoc.SelectionStart = Me.TextBoxProjLoc.Text.Length
        Me.TextBoxProjLoc.ScrollToCaret()

        Me.toggleOkButton()
    End Sub


    Private Function getProjOutputPath()
        Dim str As String = ""
        If Not projLocBase.EndsWith("\") Then
            str = "\"
        End If
        If projName = "" Then
            str = projLocBase & str
        Else
            str = projLocBase & str & projName & "\"
        End If

        Return str
    End Function


    Private Sub ToolBarViewStyle_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBarViewStyle.ButtonClick
        If e.Button Is Me.ToolBarButtonLargeIcon And Me.ToolBarButtonLargeIcon.Pushed = True Then
            Me.ToolBarButtonSmallIcon.Pushed = False
            Me.ListViewChooseNew.View = View.LargeIcon
        ElseIf e.Button Is Me.ToolBarButtonSmallIcon And Me.ToolBarButtonSmallIcon.Pushed = True Then
            Me.ToolBarButtonLargeIcon.Pushed = False
            Me.ListViewChooseNew.View = View.SmallIcon
        End If
    End Sub


    ' ListView Functions.

    Private Sub ListViewChooseNew_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListViewChooseNew.MouseMove
        Dim lvi As ListViewItem = Me.ListViewChooseNew.GetItemAt(e.X, e.Y)

        ' Check if mouse is paused over an actual node.
        If Not (lvi Is Nothing) Then
            ' Only update the ToolTip if tip needs to be changed.
            If (lvi.Tag <> ToolTip1.GetToolTip(Me.ListViewChooseNew)) Then
                ToolTip1.SetToolTip(Me.ListViewChooseNew, lvi.Tag)
            End If
        Else
            ' Mouse is not paused over a node. Therefore, clear the ToolTip.
            ToolTip1.SetToolTip(Me.ListViewChooseNew, "")
        End If

    End Sub


    Private Sub ListViewChooseNew_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewChooseNew.SelectedIndexChanged
        Me.toggleOkButton()
    End Sub


    ' Functions for control buttons.

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.processOK()
    End Sub

    Private Sub processOK()
        Dim projPath As String = Me.getProjOutputPath()
        Dim errInfo As ModuleMain.xcValidateNewFileStatus
        Dim projType As ModuleMain.xcProjectType

        If Me.ListViewChooseNew.SelectedIndices.Item(0) = 0 Then
            projType = ModuleMain.xcProjectType.xcGeneralIMSWebsiteProject
        Else
            MsgBox("Unknown project type.")
            Exit Sub
        End If

        If CType(Me.Owner, FormMain).createNewProject(Me.projName, projPath, projType) = False Then
            Me.TextBoxProjName.SelectAll()
            Me.TextBoxProjName.Focus()
            Exit Sub
        End If

        Me.Hide()
    End Sub


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Hide()
    End Sub


    Private Sub btnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelp.Click
        Help.ShowHelp(Me, SetupManager.HELP_FILE)
    End Sub


    ' Helper functions

    ' After function ListView_SelctedIndexChanged
    Private Sub toggleOkButton()
        If Me.ListViewChooseNew.SelectedItems.Count = 0 Then
            Me.resetListViewBackColor(Nothing)
            Me.btnOK.Enabled = False
        Else
            Me.resetListViewBackColor(Me.ListViewChooseNew.SelectedItems.Item(0))
            If Me.TextBoxProjName.Text = "" Then
                Me.btnOK.Enabled = False
            Else
                Me.btnOK.Enabled = True
            End If
        End If
    End Sub


    Private Sub resetListViewBackColor(ByRef selectedItem)
        Dim item As ListViewItem
        For Each item In Me.ListViewChooseNew.Items
            If item Is selectedItem Then
                item.BackColor = Color.LightGray
            Else
                item.BackColor = Color.White
            End If
        Next
    End Sub


    ' Open folder
    Private Sub openFolder()
        ' Use the FormOpenFolderDialog I wrote myself
        'If Me.frmOpenFolder Is Nothing Then
        'frmOpenFolder = New FormOpenFolderDialog()
        'End If
        'frmOpenFolder.ShowDialog(Me)

        Dim fBrowser As New clsFolderBrowser()
        fBrowser.Description = "The selected folder will be used as the location of the project folder. Please select the target folder and click OK:"

        fBrowser.ShowDialog()
        If fBrowser.DirectoryPath <> "" Then
            'MsgBox(fBrowser.DirectoryPath)
            Me.projLocBase = fBrowser.DirectoryPath
            Me.initProjOutputPath()
        End If
    End Sub


    Private Sub btnBrowseFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseFolder.Click
        openFolder()
    End Sub


    Private Sub TextBoxFileName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxProjName.TextChanged
        If IOManager.isValidFileName(Me.TextBoxProjName.Text, True) Then
            Me.projName = Me.TextBoxProjName.Text
            Me.initProjOutputPath()
        Else
            Me.TextBoxProjName.Text = Me.projName
            Me.TextBoxProjName.SelectionStart = Me.TextBoxProjName.Text.Length
            Me.TextBoxProjName.ScrollToCaret()
        End If
    End Sub


    Private Sub TextBoxFileName_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxProjName.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Me.processOK()
        End If
    End Sub

End Class
