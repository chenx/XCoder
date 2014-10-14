Public Class FormNewModule_bck
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnBrowseFolder As System.Windows.Forms.Button
    Friend WithEvents LabelFileLoc As System.Windows.Forms.Label
    Friend WithEvents TextBoxName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxLoc As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ListViewItem1 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Web site framework", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 0)
        Dim ListViewItem2 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Login Module", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim ListViewItem3 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Information Management Module", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim ListViewItem4 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Calendar Module", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim ListViewItem5 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Discussion Board", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim ListViewItem6 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Search Module", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim ListViewItem7 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Upload Module", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, 1)
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormNewModule))
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
        Me.TextBoxName = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TextBoxLoc = New System.Windows.Forms.TextBox()
        Me.btnBrowseFolder = New System.Windows.Forms.Button()
        Me.LabelFileLoc = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ListViewChooseNew
        '
        Me.ListViewChooseNew.BackColor = System.Drawing.SystemColors.Window
        ListViewItem1.Tag = "framework"
        ListViewItem7.Tag = "Upload"
        Me.ListViewChooseNew.Items.AddRange(New System.Windows.Forms.ListViewItem() {ListViewItem1, ListViewItem2, ListViewItem3, ListViewItem4, ListViewItem5, ListViewItem6, ListViewItem7})
        Me.ListViewChooseNew.LargeImageList = Me.ImageListChooseNew
        Me.ListViewChooseNew.Location = New System.Drawing.Point(16, 40)
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
        Me.Label1.Text = "Select the type of module you will create:"
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
        'TextBoxName
        '
        Me.TextBoxName.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxName.HideSelection = False
        Me.TextBoxName.Location = New System.Drawing.Point(80, 240)
        Me.TextBoxName.Name = "TextBoxName"
        Me.TextBoxName.Size = New System.Drawing.Size(416, 22)
        Me.TextBoxName.TabIndex = 12
        Me.TextBoxName.Text = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(16, 312)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(480, 3)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        '
        'TextBoxLoc
        '
        Me.TextBoxLoc.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxLoc.HideSelection = False
        Me.TextBoxLoc.Location = New System.Drawing.Point(80, 272)
        Me.TextBoxLoc.Name = "TextBoxLoc"
        Me.TextBoxLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TextBoxLoc.Size = New System.Drawing.Size(335, 22)
        Me.TextBoxLoc.TabIndex = 17
        Me.TextBoxLoc.Text = ""
        '
        'btnBrowseFolder
        '
        Me.btnBrowseFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowseFolder.Location = New System.Drawing.Point(424, 272)
        Me.btnBrowseFolder.Name = "btnBrowseFolder"
        Me.btnBrowseFolder.Size = New System.Drawing.Size(72, 28)
        Me.btnBrowseFolder.TabIndex = 14
        Me.btnBrowseFolder.Text = "Browse..."
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
        'FormNewModule
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(514, 370)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBoxLoc, Me.GroupBox1, Me.btnBrowseFolder, Me.TextBoxName, Me.LabelFileLoc, Me.LabelFileName, Me.Label1, Me.ToolBarViewStyle, Me.btnHelp, Me.btnCancel, Me.btnOK, Me.ListViewChooseNew})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormNewModule"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "New Module"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim moduleName As String
    Dim modulePathBase As String


    Private Sub FormNewModule_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ListViewChooseNew.Items.Item(0).Selected = True
        Me.ListViewChooseNew.Items.Item(0).Focused = True

        Me.ToolBarButtonLargeIcon.Pushed = True
        Me.ToolBarButtonSmallIcon.Pushed = False
        Me.ListViewChooseNew.View = View.LargeIcon

        Me.moduleName = ""
        Me.TextBoxName.Text = ""
        Me.modulePathBase = MyUtil.appendToEnd(CType(Me.Owner, FormMain).curProjectPath, "\")
        Me.initModulePath()

        Me.TextBoxName.Select()
    End Sub


    Private Sub initModulePath()
        Me.TextBoxLoc.Text = Me.modulePathBase

        Me.TextBoxLoc.SelectionStart = Me.TextBoxLoc.Text.Length
        Me.TextBoxLoc.ScrollToCaret()

        Me.toggleOkButton()
    End Sub


    Private Function getModuleFullPath()
        Return modulePathBase & moduleName & ".xc"
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

    Private Sub ListViewChooseNew_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewChooseNew.SelectedIndexChanged
        Me.toggleOkButton()
    End Sub


    ' Functions for control buttons.

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        processOK()
    End Sub

    Private Sub processOK()
        Dim moduleFullPath As String = Me.getModuleFullPath()
        Dim errInfo As ModuleMain.xcValidateNewFileStatus

        If Me.ListViewChooseNew.FocusedItem Is Nothing Then
            MsgBox("Please select a module type.")
            Exit Sub
        End If

        Dim moduleType As ModuleMain.xcModuleType = getModuleType(Me.ListViewChooseNew.FocusedItem)

        If CType(Me.Owner, FormMain).createNewModule(Me.moduleName, moduleFullPath, _
            moduleType, errInfo) = False Then
            If errInfo = ModuleMain.xcValidateNewFileStatus.isOpening Then
                MsgBox("Module " & moduleFullPath & " is currently openned. ")
            ElseIf errInfo = ModuleMain.xcValidateNewFileStatus.alreadyExists Then
                MsgBox("Module " & moduleFullPath & " already exists. Please choose another name.")
            End If
            Exit Sub
        End If

        Me.Hide()

    End Sub


    Private Function getModuleType(ByVal item As ListViewItem) As ModuleMain.xcModuleType
        If (Not item Is Nothing) Then
            Dim txt As String = LCase(Trim(item.Text))

            If txt = "web site framework" Then
                Return ModuleMain.xcModuleType.xcWebSiteFramework
            ElseIf txt = "login module" Then
                Return ModuleMain.xcModuleType.xcLogin
            ElseIf txt = "information management module" Then
                Return ModuleMain.xcModuleType.xcIMS
            ElseIf txt = "calendar module" Then
                Return ModuleMain.xcModuleType.xcCalendar
            ElseIf txt = "discussion board" Then
                Return ModuleMain.xcModuleType.xcDiscBoard
            ElseIf txt = "search module" Then
                Return ModuleMain.xcModuleType.xcSearch
            ElseIf txt = "upload module" Then
                Return ModuleMain.xcModuleType.xcUpload
            Else
                Return ModuleMain.xcModuleType.xcTemplateBase
            End If
        End If
    End Function


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
            If Me.TextBoxName.Text = "" Then
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
        Dim fBrowser As New clsFolderBrowser()
        fBrowser.Description = "The selected folder will be used to store the module file. Please select the target folder and click OK:"

        fBrowser.ShowDialog()
        If fBrowser.DirectoryPath <> "" Then
            Me.modulePathBase = fBrowser.DirectoryPath
            Me.initModulePath()
        End If
    End Sub


    Private Sub btnBrowseFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseFolder.Click
        openFolder()
    End Sub


    Private Sub TextBoxFileName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxName.TextChanged
        If IOManager.isValidFileName(Me.TextBoxName.Text, True) Then
            Me.moduleName = Me.TextBoxName.Text
            Me.initModulePath()
        Else
            Me.TextBoxName.Text = Me.moduleName
            Me.TextBoxName.SelectionStart = Me.TextBoxName.Text.Length
            Me.TextBoxName.ScrollToCaret()
        End If
    End Sub


    Private Sub TextBoxFileName_KeyPress(ByVal sender As System.Object, ByVal e As KeyPressEventArgs) Handles TextBoxName.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Me.processOK()
        End If
    End Sub
End Class
