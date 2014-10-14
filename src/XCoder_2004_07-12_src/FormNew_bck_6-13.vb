Public Class FormNew
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
    Friend WithEvents TreeViewNew As System.Windows.Forms.TreeView
    Friend WithEvents ImageListNewModule As System.Windows.Forms.ImageList
    Friend WithEvents ImageListNewType As System.Windows.Forms.ImageList
    Friend WithEvents ListViewNewModule As System.Windows.Forms.ListView
    Friend WithEvents TextBoxDesc As System.Windows.Forms.TextBox
    Friend WithEvents ListViewNewProject As System.Windows.Forms.ListView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ListViewItem1 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Web site framework", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 0)
        Dim ListViewItem2 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Login Module", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim ListViewItem3 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Information Management Module", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim ListViewItem4 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Calendar Module", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim ListViewItem5 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Discussion Board", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim ListViewItem6 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Search Module", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim ListViewItem7 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Upload Module", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormNew))
        Dim ListViewItem8 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"Project"}, 0, System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
        Me.ListViewNewModule = New System.Windows.Forms.ListView()
        Me.ImageListNewModule = New System.Windows.Forms.ImageList(Me.components)
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
        Me.TreeViewNew = New System.Windows.Forms.TreeView()
        Me.ImageListNewType = New System.Windows.Forms.ImageList(Me.components)
        Me.TextBoxDesc = New System.Windows.Forms.TextBox()
        Me.ListViewNewProject = New System.Windows.Forms.ListView()
        Me.SuspendLayout()
        '
        'ListViewNewModule
        '
        Me.ListViewNewModule.BackColor = System.Drawing.SystemColors.Window
        ListViewItem1.Tag = "framework"
        ListViewItem7.Tag = "Upload"
        Me.ListViewNewModule.Items.AddRange(New System.Windows.Forms.ListViewItem() {ListViewItem1, ListViewItem2, ListViewItem3, ListViewItem4, ListViewItem5, ListViewItem6, ListViewItem7})
        Me.ListViewNewModule.LargeImageList = Me.ImageListNewModule
        Me.ListViewNewModule.Location = New System.Drawing.Point(184, 40)
        Me.ListViewNewModule.MultiSelect = False
        Me.ListViewNewModule.Name = "ListViewNewModule"
        Me.ListViewNewModule.Size = New System.Drawing.Size(384, 200)
        Me.ListViewNewModule.SmallImageList = Me.ImageListNewModule
        Me.ListViewNewModule.TabIndex = 0
        Me.ListViewNewModule.Visible = False
        '
        'ImageListNewModule
        '
        Me.ImageListNewModule.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageListNewModule.ImageSize = New System.Drawing.Size(32, 32)
        Me.ImageListNewModule.ImageStream = CType(resources.GetObject("ImageListNewModule.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListNewModule.TransparentColor = System.Drawing.Color.Transparent
        '
        'btnOK
        '
        Me.btnOK.Enabled = False
        Me.btnOK.Location = New System.Drawing.Point(264, 368)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(96, 28)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(368, 368)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(96, 28)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        '
        'btnHelp
        '
        Me.btnHelp.Location = New System.Drawing.Point(472, 368)
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
        Me.ToolBarViewStyle.Location = New System.Drawing.Point(520, 8)
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
        Me.ToolBarButtonSmallIcon.ImageIndex = 2
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
        Me.Label1.Location = New System.Drawing.Point(24, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(400, 24)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Select the type of file you will create:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelFileName
        '
        Me.LabelFileName.Location = New System.Drawing.Point(24, 272)
        Me.LabelFileName.Name = "LabelFileName"
        Me.LabelFileName.Size = New System.Drawing.Size(72, 23)
        Me.LabelFileName.TabIndex = 10
        Me.LabelFileName.Text = "Name:"
        Me.LabelFileName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelFileLoc
        '
        Me.LabelFileLoc.Location = New System.Drawing.Point(24, 304)
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
        Me.TextBoxProjName.Location = New System.Drawing.Point(88, 272)
        Me.TextBoxProjName.Name = "TextBoxProjName"
        Me.TextBoxProjName.Size = New System.Drawing.Size(480, 22)
        Me.TextBoxProjName.TabIndex = 12
        Me.TextBoxProjName.Text = ""
        '
        'btnBrowseFolder
        '
        Me.btnBrowseFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowseFolder.Location = New System.Drawing.Point(480, 304)
        Me.btnBrowseFolder.Name = "btnBrowseFolder"
        Me.btnBrowseFolder.Size = New System.Drawing.Size(88, 28)
        Me.btnBrowseFolder.TabIndex = 14
        Me.btnBrowseFolder.Text = "Browse..."
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(24, 352)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(544, 3)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        '
        'TextBoxProjLoc
        '
        Me.TextBoxProjLoc.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxProjLoc.HideSelection = False
        Me.TextBoxProjLoc.Location = New System.Drawing.Point(88, 304)
        Me.TextBoxProjLoc.Name = "TextBoxProjLoc"
        Me.TextBoxProjLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TextBoxProjLoc.Size = New System.Drawing.Size(384, 22)
        Me.TextBoxProjLoc.TabIndex = 17
        Me.TextBoxProjLoc.Text = ""
        '
        'CheckBox1
        '
        Me.CheckBox1.Location = New System.Drawing.Point(24, 368)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(128, 24)
        Me.CheckBox1.TabIndex = 18
        Me.CheckBox1.Text = "Show me at start"
        '
        'TreeViewNew
        '
        Me.TreeViewNew.ImageList = Me.ImageListNewType
        Me.TreeViewNew.Location = New System.Drawing.Point(24, 40)
        Me.TreeViewNew.Name = "TreeViewNew"
        Me.TreeViewNew.Nodes.AddRange(New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Project", 0, 1), New System.Windows.Forms.TreeNode("Module", 0, 1)})
        Me.TreeViewNew.Scrollable = False
        Me.TreeViewNew.Size = New System.Drawing.Size(152, 200)
        Me.TreeViewNew.TabIndex = 19
        '
        'ImageListNewType
        '
        Me.ImageListNewType.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageListNewType.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageListNewType.ImageStream = CType(resources.GetObject("ImageListNewType.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListNewType.TransparentColor = System.Drawing.Color.Transparent
        '
        'TextBoxDesc
        '
        Me.TextBoxDesc.Location = New System.Drawing.Point(24, 240)
        Me.TextBoxDesc.Name = "TextBoxDesc"
        Me.TextBoxDesc.ReadOnly = True
        Me.TextBoxDesc.Size = New System.Drawing.Size(544, 22)
        Me.TextBoxDesc.TabIndex = 20
        Me.TextBoxDesc.Text = ""
        '
        'ListViewNewProject
        '
        Me.ListViewNewProject.Items.AddRange(New System.Windows.Forms.ListViewItem() {ListViewItem8})
        Me.ListViewNewProject.LargeImageList = Me.ImageListNewModule
        Me.ListViewNewProject.Location = New System.Drawing.Point(184, 40)
        Me.ListViewNewProject.MultiSelect = False
        Me.ListViewNewProject.Name = "ListViewNewProject"
        Me.ListViewNewProject.Size = New System.Drawing.Size(384, 200)
        Me.ListViewNewProject.SmallImageList = Me.ImageListNewModule
        Me.ListViewNewProject.TabIndex = 21
        '
        'FormNew
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(594, 410)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBoxDesc, Me.TreeViewNew, Me.CheckBox1, Me.TextBoxProjLoc, Me.GroupBox1, Me.btnBrowseFolder, Me.TextBoxProjName, Me.LabelFileLoc, Me.LabelFileName, Me.Label1, Me.ToolBarViewStyle, Me.btnHelp, Me.btnCancel, Me.btnOK, Me.ListViewNewProject, Me.ListViewNewModule})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormNew"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "New Project"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Enum newFileType
        newProject
        newModule
    End Enum

    Dim newFile As newFileType
    Dim curListView As ListView

    Dim projName As String
    Dim projLocBase As String = SetupManager.OUTPUT_PATH

    Private Function getProjOutputPath()
        Dim str As String = ""
        If Not projLocBase.EndsWith("\") Then
            str = "\"
        End If
        str = projLocBase & str & projName
        Return str
    End Function

    Private Sub FormChooseNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TreeViewNew.SelectedNode = Me.TreeViewNew.Nodes.Item(0)

        Me.curListView = Me.ListViewNewProject

        Me.ListViewNewModule.Items.Item(0).Selected = True
        Me.ListViewNewModule.Items.Item(0).Focused = True
        Me.ListViewNewModule.Visible = False
        Me.ListViewNewModule.View = View.LargeIcon

        Me.ListViewNewProject.Items.Item(0).Selected = True
        Me.ListViewNewProject.Items.Item(0).Focused = True
        Me.ListViewNewProject.Visible = True
        Me.ListViewNewProject.View = View.LargeIcon

        Me.ToolBarButtonLargeIcon.Pushed = True
        Me.ToolBarButtonSmallIcon.Pushed = False

        Me.projName = ""
        Me.TextBoxProjName.Text = ""
        Me.initProjOutputPath()

        Me.TextBoxProjName.Select()
    End Sub

    Private Sub initProjOutputPath()
        Me.TextBoxProjLoc.Text = Me.getProjOutputPath()
        If Not Me.TextBoxProjLoc.Text.EndsWith("\") Then
            Me.TextBoxProjLoc.Text += "\"
        End If

        Me.TextBoxProjLoc.SelectionStart = Me.TextBoxProjLoc.Text.Length
        Me.TextBoxProjLoc.ScrollToCaret()

        If projName = "" Or Me.ListViewNewModule.SelectedItems.Count = 0 Then
            Me.btnOK.Enabled = False
        Else
            Me.btnOK.Enabled = True
        End If
    End Sub

    Private Sub ToolBarViewStyle_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBarViewStyle.ButtonClick
        If e.Button Is Me.ToolBarButtonLargeIcon And Me.ToolBarButtonLargeIcon.Pushed = True Then
            Me.ToolBarButtonSmallIcon.Pushed = False
            Me.ListViewNewModule.View = View.LargeIcon
            Me.ListViewNewProject.View = View.LargeIcon
        ElseIf e.Button Is Me.ToolBarButtonSmallIcon And Me.ToolBarButtonSmallIcon.Pushed = True Then
            Me.ToolBarButtonLargeIcon.Pushed = False
            Me.ListViewNewModule.View = View.SmallIcon
            Me.ListViewNewProject.View = View.SmallIcon
        End If
    End Sub

    ' ListView Functions.

    Private Sub ListViewNewModule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewNewModule.SelectedIndexChanged
        Me.toggleOkButton()
    End Sub

    Private Sub ListViewNewModule_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewNewModule.DoubleClick
        If Not Me.ListViewNewModule.FocusedItem Is Nothing Then
            Me.openChildForm()
        End If
    End Sub

    ' Functions for control buttons.

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If Me.curListView Is Me.ListViewNewModule Then
            Me.openChildForm()
        Else
            Dim projOutputPath As String = Me.getProjOutputPath()

            If CType(Me.Owner, FormMain).openingProjects.Contains(UCase(projOutputPath)) Then
                MsgBox("Project " & projName & " in folder " & projOutputPath & _
                       " is currently openned. " & _
                       vbCrLf & "Please choose another project name.")
                'MsgBox("Project " & projName & " is currently openned. Please choose another project name.")
                Exit Sub
            End If

            If IOManager.folderExists(projOutputPath) Then
                'MsgBox("Folder " & projOutputPath & " already exists. Please choose another project name.")
                'Exit Sub
                If MessageBox.Show("Folder " & projOutputPath & " already exists. Do you want to overwrite this folder?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                    Me.TextBoxProjName.SelectAll()
                    Me.TextBoxProjName.Focus()
                    Exit Sub
                Else
                    If MessageBox.Show("Current contents in " & projOutputPath & " will be deleted. Do you want to proceed?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                        Me.TextBoxProjName.Focus()
                        Me.TextBoxProjName.SelectAll()
                        Exit Sub
                    End If
                    IOManager.DeleteFolder(projOutputPath, True)
                End If
            End If

            Me.createNewProject(projOutputPath)
            CType(Me.Owner, FormMain).setCurProject(Me.projName, Me.getProjOutputPath())
            Me.Hide()
        End If
    End Sub

    Private Sub createNewProject(ByVal projOutputPath As String)
        ' Create folders 
        IOManager.CreateFolder(projOutputPath) ' project folder, contains project files
        If Not projOutputPath.EndsWith("\") Then
            projOutputPath += "\"
        End If
        IOManager.CreateFolder(projOutputPath & SetupManager.OUTPUT_FOLDER) ' project output folder

        ' Create project file
        Dim projFile As String = projOutputPath & projName & ".xcproj"
        Dim str As String = ""
        str = str & "CreateDate=" & Now() & vbCrLf
        str = str & "ProjName=" & projName & vbCrLf
        str = str & "ProjFile=" & projFile & vbCrLf
        str = str & "ProjPath=" & projOutputPath & vbCrLf
        str = str & "ProjOutputPath=" & projOutputPath & SetupManager.OUTPUT_FOLDER & vbCrLf
        str = str & vbCrLf

        IOManager.SaveTextToFile(str, projFile)
    End Sub

    Private Sub openChildForm()
        ' Open corresponding child window.
        'MsgBox("Open Form: " & Me.ListViewNewModule.FocusedItem.Text)
        If (Not Me.ListViewNewModule.FocusedItem Is Nothing) Then
            Dim txt As String = LCase(Me.ListViewNewModule.FocusedItem.Text)
            If txt = "web site framework" Then
                CType(Me.Owner, FormMain).openNewProjFramework( _
                    ModuleMain.xcModuleType.xcWebSiteFramework, _
                    Me.getProjOutputPath(), projName)
            End If
            Me.Hide()
        Else
            MsgBox("Please select a project to open")
        End If
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
        If Me.curListView.SelectedItems.Count = 0 Then
            Me.resetListViewItemsBackColor(Nothing)
            Me.btnOK.Enabled = False
        Else
            Me.resetListViewItemsBackColor(Me.curListView.SelectedItems.Item(0))
            If Me.TextBoxProjName.Text = "" Then
                Me.btnOK.Enabled = False
            Else
                Me.btnOK.Enabled = True
            End If
        End If
    End Sub

    Private Sub resetListViewItemsBackColor(ByRef selectedItem As ListViewItem)
        Dim item As ListViewItem
        For Each item In Me.curListView.Items
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
        fBrowser.Description = "The selected folder will be used as the output location of the project folder. Please select the target folder and click OK:"

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
        Me.projName = Me.TextBoxProjName.Text
        Me.initProjOutputPath()
    End Sub

    Private Sub TreeViewNew_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeViewNew.AfterSelect
        If LCase(TreeViewNew.SelectedNode.Text) = "project" Then
            showLocInput(True)

            Me.curListView = Me.ListViewNewProject
            Me.newFile = newFileType.newProject
            Me.TextBoxDesc.Text = "A project contains multiple modules."
            Me.ListViewNewModule.Visible = False
            Me.ListViewNewProject.Visible = True
        ElseIf LCase(TreeViewNew.SelectedNode.Text) = "module" Then
            ' For module, no path is needed. 
            ' If open module in a project, use the project path,
            ' If open module individually, save it by itself with saveFileDialog
            showLocInput(False)

            Me.curListView = Me.ListViewNewModule
            Me.newFile = newFileType.newModule
            Me.TextBoxDesc.Text = "A module created individually is stored individually."
            Me.ListViewNewModule.Visible = True
            Me.ListViewNewProject.Visible = False
        End If
        resetTreeViewNodesBackColor(TreeViewNew.SelectedNode)
    End Sub

    Private Sub showLocInput(ByVal show As Boolean)
        Me.LabelFileLoc.Visible = show
        Me.TextBoxProjLoc.Visible = show
        Me.btnBrowseFolder.Visible = show
    End Sub

    Private Sub resetTreeViewNodesBackColor(ByRef selectedNode As TreeNode)
        Dim node As TreeNode
        For Each node In Me.TreeViewNew.Nodes
            If node Is selectedNode Then
                node.BackColor = Color.LightGray
            Else
                node.BackColor = Color.White
            End If
        Next
    End Sub

    Private Sub ListViewNewProject_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewNewProject.SelectedIndexChanged
        Me.toggleOkButton()
    End Sub
End Class
