

Public Class FormGenSearch
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormGenSearch))
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
        Me.GroupBoxStep.SuspendLayout()
        Me.SuspendLayout()
        '
        'RichTextBox1
        '
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
        'FormGenSearch
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 559)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelDisplay, Me.SplitterInstr, Me.PanelInstr, Me.SplitterStep, Me.PanelStep})
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FormGenSearch"
        Me.Text = "Search"
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
        Me.GroupBoxStep.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FormGenSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.moduleType = ModuleMain.xcModuleType.xcSearch
        Me.NUM_STEPS = 0
        Me.initModuleContent()
    End Sub


    Private Sub FormGenSearch_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize

    End Sub



    Protected Overrides Sub initTreeNodesText()

    End Sub


    Protected Overrides Sub initModuleContent()
        tmpFileFolder = SetupManager.ROOT_PATH & "tmp\" & Me.projectName & "\" & Me.moduleName & "\"
        If Not IOManager.folderExists(tmpFileFolder) Then IOManager.CreateFolder(tmpFileFolder)
        'MsgBox(Me.moduleName & ", tmpfilefolder=" & tmpFileFolder & ", tmpImagefolder=" & tmpImageFolder)

        Me.moduleOutputPath = Me.projectOutputPath & Me.moduleName & "\"

        'Me.LabelOutputPath.Text = Me.moduleOutputPath
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
