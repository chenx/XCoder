Imports System.io

Public Class FormMain
    Inherits System.Windows.Forms.Form

    ' Hide the FormChildBase type activeMDIChild's docked toolbox(es) if any.
    ' disable Alt+F4 and close button "X" on upper right corner
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Dim SC_MINIMIZE As Integer = 61472
        Dim WM_SYSCOMMAND As Integer = 274
        Dim curChild As FormChildBase
        Dim childFrm As Form
        Dim explorer As FormChildExplorerBase

        For Each childFrm In Me.MdiChildren
            curChild = Nothing
            If TypeOf (childFrm) Is FormChildBase Then
                curChild = childFrm

                If m.Msg = WM_SYSCOMMAND And m.WParam.ToInt32 = SC_MINIMIZE Then
                    For Each explorer In curChild.explorers
                        If (Not explorer Is Nothing) AndAlso explorer.isFloating Then
                            explorer.Hide()
                        End If
                    Next
                ElseIf m.Msg = WM_SYSCOMMAND And Not curChild.WindowState = FormWindowState.Minimized Then
                    For Each explorer In curChild.explorers
                        If (Not explorer Is Nothing) AndAlso explorer.isFloating Then
                            explorer.Show()
                        End If
                    Next
                End If
            End If
        Next

        MyBase.WndProc(m)
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
        Me.writeRecentProjects()    ' save list of recent projects.

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
    Friend WithEvents MainMenu As System.Windows.Forms.MainMenu
    Friend WithEvents OFDialogGenFW As System.Windows.Forms.OpenFileDialog
    Friend WithEvents StatusBarMain As System.Windows.Forms.StatusBar
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNew As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarMain As System.Windows.Forms.ToolBar
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWindow As System.Windows.Forms.MenuItem
    Friend WithEvents mnuView As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewStep As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewDisp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewInstr As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAbout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpContent As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpIndex As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWinArrange As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWinCascade As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWinTileH As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWinTileV As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOptions As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNewProj As System.Windows.Forms.MenuItem
    Friend WithEvents ImageListTooBar As System.Windows.Forms.ImageList
    Friend WithEvents ToolBarButtonSeparator1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuEditUndo As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditRedo As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditCut As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditCopy As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditPaste As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditDelete As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSeparator As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWinSeparator As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpSeparator As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditSeparator As System.Windows.Forms.MenuItem
    Friend WithEvents SaveFileDialogMain As System.Windows.Forms.SaveFileDialog
    Friend WithEvents mnuOpen As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOpenProj As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewProjExplorer As System.Windows.Forms.MenuItem
    Friend WithEvents PanelMainRight As System.Windows.Forms.Panel
    Friend WithEvents SplitterMainRight As System.Windows.Forms.Splitter
    Friend WithEvents PanelProperties As System.Windows.Forms.Panel
    Friend WithEvents PanelPropertiesHead As System.Windows.Forms.Panel
    Friend WithEvents SplitterProperties As System.Windows.Forms.Splitter
    Friend WithEvents PanelProjExplorer As System.Windows.Forms.Panel
    Friend WithEvents PanelProjExplorerHead As System.Windows.Forms.Panel
    Friend WithEvents btnCloseProjExplorer As System.Windows.Forms.Button
    Friend WithEvents LabelProjExplorerTitle As System.Windows.Forms.Label
    Friend WithEvents PanelProjExplorerBody As System.Windows.Forms.Panel
    Friend WithEvents PanelProjExplorerToolbar As System.Windows.Forms.Panel
    Friend WithEvents TreeViewProjExplorer As System.Windows.Forms.TreeView
    Friend WithEvents ContextMenuProjExplorer As System.Windows.Forms.ContextMenu
    Friend WithEvents btnCloseProperties As System.Windows.Forms.Button
    Friend WithEvents LabelPropertiesTitle As System.Windows.Forms.Label
    Friend WithEvents PanelPropertiesToolbar As System.Windows.Forms.Panel
    Friend WithEvents PanelPropertiesBody As System.Windows.Forms.Panel
    Friend WithEvents SplitterPropertiesDesc As System.Windows.Forms.Splitter
    Friend WithEvents ListViewProperties As System.Windows.Forms.ListView
    Friend WithEvents CHPropertyName As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHPropertyValue As System.Windows.Forms.ColumnHeader
    Friend WithEvents mnuViewProperties As System.Windows.Forms.MenuItem
    Friend WithEvents PanelProjExplorerBottom As System.Windows.Forms.Panel
    Friend WithEvents ImageListProjExplorer As System.Windows.Forms.ImageList
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCloseMod As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCloseAllMod As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarButtonNewModule As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonOpenProject As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonSaveProject As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuAddNewMod As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddOldMod As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCloseProj As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarButtonNewProject As System.Windows.Forms.ToolBarButton
    Friend WithEvents OpenFileDialogModule As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ContextMenuModuleNode As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMNFullPath As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMNOpen As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMNExclude As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMNDelete As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMNRename As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMNProperties As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMNClose As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenuProjNode As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuPNCloseAllMod As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPNCloseProj As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPNAddNewMod As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPNAddOldMod As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPNProperties As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPNRename As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPNSaveProj As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPNFullPath As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPEFloat As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPEHide As System.Windows.Forms.MenuItem
    Friend WithEvents TextBoxProperty As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialogProj As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ToolBarButtonExistModule As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonSeparator2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuSaveProj As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRecentProj As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSaveModule As System.Windows.Forms.MenuItem
    Friend WithEvents GroupBoxPropertyDesc As System.Windows.Forms.GroupBox
    Friend WithEvents LabelPropertyDesc As System.Windows.Forms.Label
    Friend WithEvents mnuSaveModuleAs As System.Windows.Forms.MenuItem
    Friend WithEvents SaveFileDialogModule As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ContextMenuProperty As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuPropFloat As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPropHide As System.Windows.Forms.MenuItem
    Friend WithEvents mnuBuild As System.Windows.Forms.MenuItem
    Friend WithEvents mnuBuildProject As System.Windows.Forms.MenuItem
    Friend WithEvents mnuBuildModule As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarButtonCloseProject As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonBuildProject As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonDeleteModule As System.Windows.Forms.ToolBarButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormMain))
        Me.StatusBarMain = New System.Windows.Forms.StatusBar()
        Me.MainMenu = New System.Windows.Forms.MainMenu()
        Me.mnuFile = New System.Windows.Forms.MenuItem()
        Me.mnuNew = New System.Windows.Forms.MenuItem()
        Me.mnuNewProj = New System.Windows.Forms.MenuItem()
        Me.mnuOpen = New System.Windows.Forms.MenuItem()
        Me.mnuOpenProj = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.mnuAddNewMod = New System.Windows.Forms.MenuItem()
        Me.mnuAddOldMod = New System.Windows.Forms.MenuItem()
        Me.MenuItem8 = New System.Windows.Forms.MenuItem()
        Me.mnuCloseMod = New System.Windows.Forms.MenuItem()
        Me.mnuCloseAllMod = New System.Windows.Forms.MenuItem()
        Me.mnuCloseProj = New System.Windows.Forms.MenuItem()
        Me.mnuFileSeparator = New System.Windows.Forms.MenuItem()
        Me.mnuSaveModule = New System.Windows.Forms.MenuItem()
        Me.mnuSaveModuleAs = New System.Windows.Forms.MenuItem()
        Me.mnuSaveProj = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.mnuRecentProj = New System.Windows.Forms.MenuItem()
        Me.MenuItem10 = New System.Windows.Forms.MenuItem()
        Me.mnuExit = New System.Windows.Forms.MenuItem()
        Me.mnuEdit = New System.Windows.Forms.MenuItem()
        Me.mnuEditUndo = New System.Windows.Forms.MenuItem()
        Me.mnuEditRedo = New System.Windows.Forms.MenuItem()
        Me.mnuEditSeparator = New System.Windows.Forms.MenuItem()
        Me.mnuEditCut = New System.Windows.Forms.MenuItem()
        Me.mnuEditCopy = New System.Windows.Forms.MenuItem()
        Me.mnuEditPaste = New System.Windows.Forms.MenuItem()
        Me.mnuEditDelete = New System.Windows.Forms.MenuItem()
        Me.mnuView = New System.Windows.Forms.MenuItem()
        Me.mnuViewStep = New System.Windows.Forms.MenuItem()
        Me.mnuViewDisp = New System.Windows.Forms.MenuItem()
        Me.mnuViewInstr = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuViewProjExplorer = New System.Windows.Forms.MenuItem()
        Me.mnuViewProperties = New System.Windows.Forms.MenuItem()
        Me.mnuBuild = New System.Windows.Forms.MenuItem()
        Me.mnuBuildProject = New System.Windows.Forms.MenuItem()
        Me.mnuBuildModule = New System.Windows.Forms.MenuItem()
        Me.mnuTools = New System.Windows.Forms.MenuItem()
        Me.mnuOptions = New System.Windows.Forms.MenuItem()
        Me.mnuWindow = New System.Windows.Forms.MenuItem()
        Me.mnuWinArrange = New System.Windows.Forms.MenuItem()
        Me.mnuWinCascade = New System.Windows.Forms.MenuItem()
        Me.mnuWinTileH = New System.Windows.Forms.MenuItem()
        Me.mnuWinTileV = New System.Windows.Forms.MenuItem()
        Me.mnuWinSeparator = New System.Windows.Forms.MenuItem()
        Me.mnuHelp = New System.Windows.Forms.MenuItem()
        Me.mnuHelpContent = New System.Windows.Forms.MenuItem()
        Me.mnuHelpIndex = New System.Windows.Forms.MenuItem()
        Me.mnuHelpSeparator = New System.Windows.Forms.MenuItem()
        Me.mnuAbout = New System.Windows.Forms.MenuItem()
        Me.ImageListTooBar = New System.Windows.Forms.ImageList(Me.components)
        Me.OFDialogGenFW = New System.Windows.Forms.OpenFileDialog()
        Me.ToolBarMain = New System.Windows.Forms.ToolBar()
        Me.ToolBarButtonNewProject = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButtonOpenProject = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButtonCloseProject = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButtonSeparator1 = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButtonNewModule = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButtonExistModule = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButtonDeleteModule = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButtonSeparator2 = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButtonSaveProject = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButtonBuildProject = New System.Windows.Forms.ToolBarButton()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.OpenFileDialogProj = New System.Windows.Forms.OpenFileDialog()
        Me.SaveFileDialogMain = New System.Windows.Forms.SaveFileDialog()
        Me.PanelMainRight = New System.Windows.Forms.Panel()
        Me.PanelProjExplorer = New System.Windows.Forms.Panel()
        Me.PanelProjExplorerBody = New System.Windows.Forms.Panel()
        Me.PanelProjExplorerBottom = New System.Windows.Forms.Panel()
        Me.TreeViewProjExplorer = New System.Windows.Forms.TreeView()
        Me.ImageListProjExplorer = New System.Windows.Forms.ImageList(Me.components)
        Me.PanelProjExplorerToolbar = New System.Windows.Forms.Panel()
        Me.PanelProjExplorerHead = New System.Windows.Forms.Panel()
        Me.LabelProjExplorerTitle = New System.Windows.Forms.Label()
        Me.btnCloseProjExplorer = New System.Windows.Forms.Button()
        Me.SplitterProperties = New System.Windows.Forms.Splitter()
        Me.PanelProperties = New System.Windows.Forms.Panel()
        Me.ContextMenuProperty = New System.Windows.Forms.ContextMenu()
        Me.mnuPropFloat = New System.Windows.Forms.MenuItem()
        Me.mnuPropHide = New System.Windows.Forms.MenuItem()
        Me.SplitterPropertiesDesc = New System.Windows.Forms.Splitter()
        Me.PanelPropertiesBody = New System.Windows.Forms.Panel()
        Me.ListViewProperties = New System.Windows.Forms.ListView()
        Me.CHPropertyName = New System.Windows.Forms.ColumnHeader()
        Me.CHPropertyValue = New System.Windows.Forms.ColumnHeader()
        Me.TextBoxProperty = New System.Windows.Forms.TextBox()
        Me.GroupBoxPropertyDesc = New System.Windows.Forms.GroupBox()
        Me.LabelPropertyDesc = New System.Windows.Forms.Label()
        Me.PanelPropertiesToolbar = New System.Windows.Forms.Panel()
        Me.PanelPropertiesHead = New System.Windows.Forms.Panel()
        Me.LabelPropertiesTitle = New System.Windows.Forms.Label()
        Me.btnCloseProperties = New System.Windows.Forms.Button()
        Me.ContextMenuProjExplorer = New System.Windows.Forms.ContextMenu()
        Me.mnuPEFloat = New System.Windows.Forms.MenuItem()
        Me.mnuPEHide = New System.Windows.Forms.MenuItem()
        Me.ContextMenuModuleNode = New System.Windows.Forms.ContextMenu()
        Me.mnuMNFullPath = New System.Windows.Forms.MenuItem()
        Me.mnuMNOpen = New System.Windows.Forms.MenuItem()
        Me.mnuMNClose = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.mnuMNExclude = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.mnuMNDelete = New System.Windows.Forms.MenuItem()
        Me.mnuMNRename = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        Me.mnuMNProperties = New System.Windows.Forms.MenuItem()
        Me.SplitterMainRight = New System.Windows.Forms.Splitter()
        Me.OpenFileDialogModule = New System.Windows.Forms.OpenFileDialog()
        Me.ContextMenuProjNode = New System.Windows.Forms.ContextMenu()
        Me.mnuPNFullPath = New System.Windows.Forms.MenuItem()
        Me.mnuPNCloseAllMod = New System.Windows.Forms.MenuItem()
        Me.mnuPNCloseProj = New System.Windows.Forms.MenuItem()
        Me.MenuItem18 = New System.Windows.Forms.MenuItem()
        Me.mnuPNAddNewMod = New System.Windows.Forms.MenuItem()
        Me.mnuPNAddOldMod = New System.Windows.Forms.MenuItem()
        Me.MenuItem19 = New System.Windows.Forms.MenuItem()
        Me.mnuPNSaveProj = New System.Windows.Forms.MenuItem()
        Me.mnuPNRename = New System.Windows.Forms.MenuItem()
        Me.MenuItem20 = New System.Windows.Forms.MenuItem()
        Me.mnuPNProperties = New System.Windows.Forms.MenuItem()
        Me.SaveFileDialogModule = New System.Windows.Forms.SaveFileDialog()
        Me.PanelMainRight.SuspendLayout()
        Me.PanelProjExplorer.SuspendLayout()
        Me.PanelProjExplorerBody.SuspendLayout()
        Me.PanelProjExplorerHead.SuspendLayout()
        Me.PanelProperties.SuspendLayout()
        Me.PanelPropertiesBody.SuspendLayout()
        Me.GroupBoxPropertyDesc.SuspendLayout()
        Me.PanelPropertiesHead.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusBarMain
        '
        Me.StatusBarMain.Location = New System.Drawing.Point(0, 493)
        Me.StatusBarMain.Name = "StatusBarMain"
        Me.StatusBarMain.Size = New System.Drawing.Size(824, 25)
        Me.StatusBarMain.TabIndex = 0
        '
        'MainMenu
        '
        Me.MainMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuEdit, Me.mnuView, Me.mnuBuild, Me.mnuTools, Me.mnuWindow, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNew, Me.mnuOpen, Me.MenuItem2, Me.mnuAddNewMod, Me.mnuAddOldMod, Me.MenuItem8, Me.mnuCloseMod, Me.mnuCloseAllMod, Me.mnuCloseProj, Me.mnuFileSeparator, Me.mnuSaveModule, Me.mnuSaveModuleAs, Me.mnuSaveProj, Me.MenuItem3, Me.mnuRecentProj, Me.MenuItem10, Me.mnuExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuNew
        '
        Me.mnuNew.Index = 0
        Me.mnuNew.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNewProj})
        Me.mnuNew.Text = "&New"
        '
        'mnuNewProj
        '
        Me.mnuNewProj.Index = 0
        Me.mnuNewProj.Shortcut = System.Windows.Forms.Shortcut.CtrlN
        Me.mnuNewProj.Text = "Project ..."
        '
        'mnuOpen
        '
        Me.mnuOpen.Index = 1
        Me.mnuOpen.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuOpenProj})
        Me.mnuOpen.Text = "&Open"
        '
        'mnuOpenProj
        '
        Me.mnuOpenProj.Index = 0
        Me.mnuOpenProj.Shortcut = System.Windows.Forms.Shortcut.CtrlO
        Me.mnuOpenProj.Text = "Project ..."
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.Text = "-"
        '
        'mnuAddNewMod
        '
        Me.mnuAddNewMod.Enabled = False
        Me.mnuAddNewMod.Index = 3
        Me.mnuAddNewMod.Shortcut = System.Windows.Forms.Shortcut.CtrlShiftN
        Me.mnuAddNewMod.Text = "Add Ne&w Module..."
        '
        'mnuAddOldMod
        '
        Me.mnuAddOldMod.Enabled = False
        Me.mnuAddOldMod.Index = 4
        Me.mnuAddOldMod.Shortcut = System.Windows.Forms.Shortcut.CtrlShiftO
        Me.mnuAddOldMod.Text = "Add Existin&g Module..."
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 5
        Me.MenuItem8.Text = "-"
        '
        'mnuCloseMod
        '
        Me.mnuCloseMod.Enabled = False
        Me.mnuCloseMod.Index = 6
        Me.mnuCloseMod.Text = "&Close Module"
        '
        'mnuCloseAllMod
        '
        Me.mnuCloseAllMod.Enabled = False
        Me.mnuCloseAllMod.Index = 7
        Me.mnuCloseAllMod.Text = "Close &All Modules"
        '
        'mnuCloseProj
        '
        Me.mnuCloseProj.Enabled = False
        Me.mnuCloseProj.Index = 8
        Me.mnuCloseProj.Text = "Close &Project"
        '
        'mnuFileSeparator
        '
        Me.mnuFileSeparator.Index = 9
        Me.mnuFileSeparator.Text = "-"
        '
        'mnuSaveModule
        '
        Me.mnuSaveModule.Enabled = False
        Me.mnuSaveModule.Index = 10
        Me.mnuSaveModule.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.mnuSaveModule.Text = "&Save Module"
        '
        'mnuSaveModuleAs
        '
        Me.mnuSaveModuleAs.Enabled = False
        Me.mnuSaveModuleAs.Index = 11
        Me.mnuSaveModuleAs.Shortcut = System.Windows.Forms.Shortcut.CtrlA
        Me.mnuSaveModuleAs.Text = "Save Module &As ..."
        '
        'mnuSaveProj
        '
        Me.mnuSaveProj.Enabled = False
        Me.mnuSaveProj.Index = 12
        Me.mnuSaveProj.Shortcut = System.Windows.Forms.Shortcut.CtrlShiftS
        Me.mnuSaveProj.Text = "Save Projec&t"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 13
        Me.MenuItem3.Text = "-"
        '
        'mnuRecentProj
        '
        Me.mnuRecentProj.Index = 14
        Me.mnuRecentProj.Text = "Recent Pro&jects"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 15
        Me.MenuItem10.Text = "-"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 16
        Me.mnuExit.Text = "E&xit"
        '
        'mnuEdit
        '
        Me.mnuEdit.Index = 1
        Me.mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuEditUndo, Me.mnuEditRedo, Me.mnuEditSeparator, Me.mnuEditCut, Me.mnuEditCopy, Me.mnuEditPaste, Me.mnuEditDelete})
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuEditUndo
        '
        Me.mnuEditUndo.Enabled = False
        Me.mnuEditUndo.Index = 0
        Me.mnuEditUndo.Text = "&Undo"
        '
        'mnuEditRedo
        '
        Me.mnuEditRedo.Enabled = False
        Me.mnuEditRedo.Index = 1
        Me.mnuEditRedo.Text = "&Redo"
        '
        'mnuEditSeparator
        '
        Me.mnuEditSeparator.Index = 2
        Me.mnuEditSeparator.Text = "-"
        '
        'mnuEditCut
        '
        Me.mnuEditCut.Enabled = False
        Me.mnuEditCut.Index = 3
        Me.mnuEditCut.Text = "Cu&t"
        '
        'mnuEditCopy
        '
        Me.mnuEditCopy.Enabled = False
        Me.mnuEditCopy.Index = 4
        Me.mnuEditCopy.Text = "&Copy"
        '
        'mnuEditPaste
        '
        Me.mnuEditPaste.Enabled = False
        Me.mnuEditPaste.Index = 5
        Me.mnuEditPaste.Text = "&Paste"
        '
        'mnuEditDelete
        '
        Me.mnuEditDelete.Enabled = False
        Me.mnuEditDelete.Index = 6
        Me.mnuEditDelete.Text = "&Delete"
        '
        'mnuView
        '
        Me.mnuView.Index = 2
        Me.mnuView.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuViewStep, Me.mnuViewDisp, Me.mnuViewInstr, Me.MenuItem1, Me.mnuViewProjExplorer, Me.mnuViewProperties})
        Me.mnuView.Text = "&View"
        '
        'mnuViewStep
        '
        Me.mnuViewStep.Enabled = False
        Me.mnuViewStep.Index = 0
        Me.mnuViewStep.Text = "Step Explorer"
        '
        'mnuViewDisp
        '
        Me.mnuViewDisp.Enabled = False
        Me.mnuViewDisp.Index = 1
        Me.mnuViewDisp.Text = "Display Explorer"
        '
        'mnuViewInstr
        '
        Me.mnuViewInstr.Enabled = False
        Me.mnuViewInstr.Index = 2
        Me.mnuViewInstr.Text = "Option Explorer"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 3
        Me.MenuItem1.Text = "-"
        '
        'mnuViewProjExplorer
        '
        Me.mnuViewProjExplorer.Index = 4
        Me.mnuViewProjExplorer.Shortcut = System.Windows.Forms.Shortcut.F3
        Me.mnuViewProjExplorer.Text = "Project Ex&plorer"
        '
        'mnuViewProperties
        '
        Me.mnuViewProperties.Index = 5
        Me.mnuViewProperties.Shortcut = System.Windows.Forms.Shortcut.F4
        Me.mnuViewProperties.Text = "Properties &Window"
        '
        'mnuBuild
        '
        Me.mnuBuild.Index = 3
        Me.mnuBuild.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuBuildProject, Me.mnuBuildModule})
        Me.mnuBuild.Text = "&Build"
        '
        'mnuBuildProject
        '
        Me.mnuBuildProject.Enabled = False
        Me.mnuBuildProject.Index = 0
        Me.mnuBuildProject.Shortcut = System.Windows.Forms.Shortcut.F5
        Me.mnuBuildProject.Text = "&Build Project"
        '
        'mnuBuildModule
        '
        Me.mnuBuildModule.Enabled = False
        Me.mnuBuildModule.Index = 1
        Me.mnuBuildModule.Shortcut = System.Windows.Forms.Shortcut.F6
        Me.mnuBuildModule.Text = "B&uild Module"
        '
        'mnuTools
        '
        Me.mnuTools.Index = 4
        Me.mnuTools.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuOptions})
        Me.mnuTools.Text = "&Tools"
        '
        'mnuOptions
        '
        Me.mnuOptions.Index = 0
        Me.mnuOptions.Text = "&Options"
        '
        'mnuWindow
        '
        Me.mnuWindow.Index = 5
        Me.mnuWindow.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuWinArrange, Me.mnuWinCascade, Me.mnuWinTileH, Me.mnuWinTileV, Me.mnuWinSeparator})
        Me.mnuWindow.Text = "&Window"
        '
        'mnuWinArrange
        '
        Me.mnuWinArrange.Enabled = False
        Me.mnuWinArrange.Index = 0
        Me.mnuWinArrange.Text = "Arrange Icons"
        '
        'mnuWinCascade
        '
        Me.mnuWinCascade.Enabled = False
        Me.mnuWinCascade.Index = 1
        Me.mnuWinCascade.Text = "Cascade"
        '
        'mnuWinTileH
        '
        Me.mnuWinTileH.Enabled = False
        Me.mnuWinTileH.Index = 2
        Me.mnuWinTileH.Text = "Tile Horizontally"
        '
        'mnuWinTileV
        '
        Me.mnuWinTileV.Enabled = False
        Me.mnuWinTileV.Index = 3
        Me.mnuWinTileV.Text = "Tile Vertically"
        '
        'mnuWinSeparator
        '
        Me.mnuWinSeparator.Index = 4
        Me.mnuWinSeparator.Text = "-"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 6
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelpContent, Me.mnuHelpIndex, Me.mnuHelpSeparator, Me.mnuAbout})
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpContent
        '
        Me.mnuHelpContent.Index = 0
        Me.mnuHelpContent.Text = "&Content"
        '
        'mnuHelpIndex
        '
        Me.mnuHelpIndex.Index = 1
        Me.mnuHelpIndex.Text = "&Index"
        '
        'mnuHelpSeparator
        '
        Me.mnuHelpSeparator.Index = 2
        Me.mnuHelpSeparator.Text = "-"
        '
        'mnuAbout
        '
        Me.mnuAbout.Index = 3
        Me.mnuAbout.Text = "&About ..."
        '
        'ImageListTooBar
        '
        Me.ImageListTooBar.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageListTooBar.ImageSize = New System.Drawing.Size(20, 20)
        Me.ImageListTooBar.ImageStream = CType(resources.GetObject("ImageListTooBar.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListTooBar.TransparentColor = System.Drawing.Color.Transparent
        '
        'OFDialogGenFW
        '
        Me.OFDialogGenFW.DefaultExt = "txt"
        '
        'ToolBarMain
        '
        Me.ToolBarMain.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButtonNewProject, Me.ToolBarButtonOpenProject, Me.ToolBarButtonCloseProject, Me.ToolBarButtonSeparator1, Me.ToolBarButtonNewModule, Me.ToolBarButtonExistModule, Me.ToolBarButtonDeleteModule, Me.ToolBarButtonSeparator2, Me.ToolBarButtonSaveProject, Me.ToolBarButtonBuildProject})
        Me.ToolBarMain.ButtonSize = New System.Drawing.Size(24, 24)
        Me.ToolBarMain.DropDownArrows = True
        Me.ToolBarMain.ImageList = Me.ImageListTooBar
        Me.ToolBarMain.Name = "ToolBarMain"
        Me.ToolBarMain.ShowToolTips = True
        Me.ToolBarMain.Size = New System.Drawing.Size(824, 29)
        Me.ToolBarMain.TabIndex = 2
        '
        'ToolBarButtonNewProject
        '
        Me.ToolBarButtonNewProject.ImageIndex = 8
        Me.ToolBarButtonNewProject.ToolTipText = "New Project"
        '
        'ToolBarButtonOpenProject
        '
        Me.ToolBarButtonOpenProject.ImageIndex = 1
        Me.ToolBarButtonOpenProject.ToolTipText = "Open Project"
        '
        'ToolBarButtonCloseProject
        '
        Me.ToolBarButtonCloseProject.ImageIndex = 11
        Me.ToolBarButtonCloseProject.ToolTipText = "Close Project"
        '
        'ToolBarButtonSeparator1
        '
        Me.ToolBarButtonSeparator1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'ToolBarButtonNewModule
        '
        Me.ToolBarButtonNewModule.ImageIndex = 0
        Me.ToolBarButtonNewModule.ToolTipText = "Add New Module"
        '
        'ToolBarButtonExistModule
        '
        Me.ToolBarButtonExistModule.ImageIndex = 9
        Me.ToolBarButtonExistModule.ToolTipText = "Add Existing Module"
        '
        'ToolBarButtonDeleteModule
        '
        Me.ToolBarButtonDeleteModule.ImageIndex = 10
        Me.ToolBarButtonDeleteModule.ToolTipText = "Delete Module"
        '
        'ToolBarButtonSeparator2
        '
        Me.ToolBarButtonSeparator2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'ToolBarButtonSaveProject
        '
        Me.ToolBarButtonSaveProject.ImageIndex = 2
        Me.ToolBarButtonSaveProject.ToolTipText = "Save Project"
        '
        'ToolBarButtonBuildProject
        '
        Me.ToolBarButtonBuildProject.ImageIndex = 12
        Me.ToolBarButtonBuildProject.ToolTipText = "Build Project"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = -1
        Me.MenuItem4.Text = ""
        '
        'OpenFileDialogProj
        '
        Me.OpenFileDialogProj.DefaultExt = "xcproj"
        Me.OpenFileDialogProj.Filter = """XC Project files (*.xcproj)|*.xcproj"""
        '
        'SaveFileDialogMain
        '
        Me.SaveFileDialogMain.DefaultExt = "xcproj"
        Me.SaveFileDialogMain.FileName = "project 1"
        Me.SaveFileDialogMain.Filter = """XC Project files (*.xcproj)|*.xcproj"""
        Me.SaveFileDialogMain.Title = "Save File"
        '
        'PanelMainRight
        '
        Me.PanelMainRight.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelProjExplorer, Me.SplitterProperties, Me.PanelProperties})
        Me.PanelMainRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.PanelMainRight.Location = New System.Drawing.Point(574, 29)
        Me.PanelMainRight.Name = "PanelMainRight"
        Me.PanelMainRight.Size = New System.Drawing.Size(250, 464)
        Me.PanelMainRight.TabIndex = 7
        '
        'PanelProjExplorer
        '
        Me.PanelProjExplorer.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelProjExplorerBody, Me.PanelProjExplorerHead})
        Me.PanelProjExplorer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelProjExplorer.Name = "PanelProjExplorer"
        Me.PanelProjExplorer.Size = New System.Drawing.Size(250, 61)
        Me.PanelProjExplorer.TabIndex = 4
        '
        'PanelProjExplorerBody
        '
        Me.PanelProjExplorerBody.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelProjExplorerBottom, Me.TreeViewProjExplorer, Me.PanelProjExplorerToolbar})
        Me.PanelProjExplorerBody.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelProjExplorerBody.Location = New System.Drawing.Point(0, 25)
        Me.PanelProjExplorerBody.Name = "PanelProjExplorerBody"
        Me.PanelProjExplorerBody.Size = New System.Drawing.Size(250, 36)
        Me.PanelProjExplorerBody.TabIndex = 2
        '
        'PanelProjExplorerBottom
        '
        Me.PanelProjExplorerBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelProjExplorerBottom.Location = New System.Drawing.Point(0, 31)
        Me.PanelProjExplorerBottom.Name = "PanelProjExplorerBottom"
        Me.PanelProjExplorerBottom.Size = New System.Drawing.Size(250, 5)
        Me.PanelProjExplorerBottom.TabIndex = 2
        '
        'TreeViewProjExplorer
        '
        Me.TreeViewProjExplorer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeViewProjExplorer.ImageList = Me.ImageListProjExplorer
        Me.TreeViewProjExplorer.LabelEdit = True
        Me.TreeViewProjExplorer.Location = New System.Drawing.Point(0, 25)
        Me.TreeViewProjExplorer.Name = "TreeViewProjExplorer"
        Me.TreeViewProjExplorer.Size = New System.Drawing.Size(250, 11)
        Me.TreeViewProjExplorer.TabIndex = 1
        '
        'ImageListProjExplorer
        '
        Me.ImageListProjExplorer.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageListProjExplorer.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageListProjExplorer.ImageStream = CType(resources.GetObject("ImageListProjExplorer.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListProjExplorer.TransparentColor = System.Drawing.Color.Transparent
        '
        'PanelProjExplorerToolbar
        '
        Me.PanelProjExplorerToolbar.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelProjExplorerToolbar.Name = "PanelProjExplorerToolbar"
        Me.PanelProjExplorerToolbar.Size = New System.Drawing.Size(250, 25)
        Me.PanelProjExplorerToolbar.TabIndex = 0
        '
        'PanelProjExplorerHead
        '
        Me.PanelProjExplorerHead.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.PanelProjExplorerHead.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelProjExplorerTitle, Me.btnCloseProjExplorer})
        Me.PanelProjExplorerHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelProjExplorerHead.Name = "PanelProjExplorerHead"
        Me.PanelProjExplorerHead.Size = New System.Drawing.Size(250, 25)
        Me.PanelProjExplorerHead.TabIndex = 1
        '
        'LabelProjExplorerTitle
        '
        Me.LabelProjExplorerTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelProjExplorerTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelProjExplorerTitle.Name = "LabelProjExplorerTitle"
        Me.LabelProjExplorerTitle.Size = New System.Drawing.Size(227, 25)
        Me.LabelProjExplorerTitle.TabIndex = 1
        Me.LabelProjExplorerTitle.Text = "Project Explorer"
        Me.LabelProjExplorerTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCloseProjExplorer
        '
        Me.btnCloseProjExplorer.BackColor = System.Drawing.SystemColors.Control
        Me.btnCloseProjExplorer.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseProjExplorer.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseProjExplorer.Location = New System.Drawing.Point(227, 0)
        Me.btnCloseProjExplorer.Name = "btnCloseProjExplorer"
        Me.btnCloseProjExplorer.Size = New System.Drawing.Size(23, 25)
        Me.btnCloseProjExplorer.TabIndex = 0
        Me.btnCloseProjExplorer.Text = "X"
        '
        'SplitterProperties
        '
        Me.SplitterProperties.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.SplitterProperties.Location = New System.Drawing.Point(0, 61)
        Me.SplitterProperties.Name = "SplitterProperties"
        Me.SplitterProperties.Size = New System.Drawing.Size(250, 3)
        Me.SplitterProperties.TabIndex = 3
        Me.SplitterProperties.TabStop = False
        '
        'PanelProperties
        '
        Me.PanelProperties.ContextMenu = Me.ContextMenuProperty
        Me.PanelProperties.Controls.AddRange(New System.Windows.Forms.Control() {Me.SplitterPropertiesDesc, Me.PanelPropertiesBody, Me.GroupBoxPropertyDesc, Me.PanelPropertiesToolbar, Me.PanelPropertiesHead})
        Me.PanelProperties.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelProperties.Location = New System.Drawing.Point(0, 64)
        Me.PanelProperties.Name = "PanelProperties"
        Me.PanelProperties.Size = New System.Drawing.Size(250, 400)
        Me.PanelProperties.TabIndex = 2
        '
        'ContextMenuProperty
        '
        Me.ContextMenuProperty.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPropFloat, Me.mnuPropHide})
        '
        'mnuPropFloat
        '
        Me.mnuPropFloat.Index = 0
        Me.mnuPropFloat.Text = "&Float"
        '
        'mnuPropHide
        '
        Me.mnuPropHide.Index = 1
        Me.mnuPropHide.Text = "&Hide"
        '
        'SplitterPropertiesDesc
        '
        Me.SplitterPropertiesDesc.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.SplitterPropertiesDesc.Location = New System.Drawing.Point(0, 325)
        Me.SplitterPropertiesDesc.Name = "SplitterPropertiesDesc"
        Me.SplitterPropertiesDesc.Size = New System.Drawing.Size(250, 3)
        Me.SplitterPropertiesDesc.TabIndex = 4
        Me.SplitterPropertiesDesc.TabStop = False
        '
        'PanelPropertiesBody
        '
        Me.PanelPropertiesBody.Controls.AddRange(New System.Windows.Forms.Control() {Me.ListViewProperties, Me.TextBoxProperty})
        Me.PanelPropertiesBody.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelPropertiesBody.Location = New System.Drawing.Point(0, 50)
        Me.PanelPropertiesBody.Name = "PanelPropertiesBody"
        Me.PanelPropertiesBody.Size = New System.Drawing.Size(250, 278)
        Me.PanelPropertiesBody.TabIndex = 3
        '
        'ListViewProperties
        '
        Me.ListViewProperties.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.CHPropertyName, Me.CHPropertyValue})
        Me.ListViewProperties.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListViewProperties.FullRowSelect = True
        Me.ListViewProperties.GridLines = True
        Me.ListViewProperties.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.ListViewProperties.HideSelection = False
        Me.ListViewProperties.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ListViewProperties.LabelWrap = False
        Me.ListViewProperties.MultiSelect = False
        Me.ListViewProperties.Name = "ListViewProperties"
        Me.ListViewProperties.Size = New System.Drawing.Size(250, 278)
        Me.ListViewProperties.TabIndex = 0
        Me.ListViewProperties.View = System.Windows.Forms.View.Details
        '
        'CHPropertyName
        '
        Me.CHPropertyName.Text = "Attribute"
        Me.CHPropertyName.Width = 100
        '
        'CHPropertyValue
        '
        Me.CHPropertyValue.Text = "Value"
        Me.CHPropertyValue.Width = 150
        '
        'TextBoxProperty
        '
        Me.TextBoxProperty.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxProperty.Location = New System.Drawing.Point(40, 160)
        Me.TextBoxProperty.Name = "TextBoxProperty"
        Me.TextBoxProperty.TabIndex = 10
        Me.TextBoxProperty.Text = ""
        Me.TextBoxProperty.Visible = False
        '
        'GroupBoxPropertyDesc
        '
        Me.GroupBoxPropertyDesc.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelPropertyDesc})
        Me.GroupBoxPropertyDesc.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBoxPropertyDesc.Location = New System.Drawing.Point(0, 328)
        Me.GroupBoxPropertyDesc.Name = "GroupBoxPropertyDesc"
        Me.GroupBoxPropertyDesc.Size = New System.Drawing.Size(250, 72)
        Me.GroupBoxPropertyDesc.TabIndex = 2
        Me.GroupBoxPropertyDesc.TabStop = False
        '
        'LabelPropertyDesc
        '
        Me.LabelPropertyDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelPropertyDesc.Location = New System.Drawing.Point(3, 18)
        Me.LabelPropertyDesc.Name = "LabelPropertyDesc"
        Me.LabelPropertyDesc.Size = New System.Drawing.Size(244, 51)
        Me.LabelPropertyDesc.TabIndex = 0
        '
        'PanelPropertiesToolbar
        '
        Me.PanelPropertiesToolbar.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelPropertiesToolbar.Location = New System.Drawing.Point(0, 25)
        Me.PanelPropertiesToolbar.Name = "PanelPropertiesToolbar"
        Me.PanelPropertiesToolbar.Size = New System.Drawing.Size(250, 25)
        Me.PanelPropertiesToolbar.TabIndex = 1
        '
        'PanelPropertiesHead
        '
        Me.PanelPropertiesHead.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.PanelPropertiesHead.Controls.AddRange(New System.Windows.Forms.Control() {Me.LabelPropertiesTitle, Me.btnCloseProperties})
        Me.PanelPropertiesHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelPropertiesHead.Name = "PanelPropertiesHead"
        Me.PanelPropertiesHead.Size = New System.Drawing.Size(250, 25)
        Me.PanelPropertiesHead.TabIndex = 0
        '
        'LabelPropertiesTitle
        '
        Me.LabelPropertiesTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelPropertiesTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelPropertiesTitle.Name = "LabelPropertiesTitle"
        Me.LabelPropertiesTitle.Size = New System.Drawing.Size(227, 25)
        Me.LabelPropertiesTitle.TabIndex = 2
        Me.LabelPropertiesTitle.Text = "Properties"
        Me.LabelPropertiesTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCloseProperties
        '
        Me.btnCloseProperties.BackColor = System.Drawing.SystemColors.Control
        Me.btnCloseProperties.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCloseProperties.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseProperties.Location = New System.Drawing.Point(227, 0)
        Me.btnCloseProperties.Name = "btnCloseProperties"
        Me.btnCloseProperties.Size = New System.Drawing.Size(23, 25)
        Me.btnCloseProperties.TabIndex = 1
        Me.btnCloseProperties.Text = "X"
        '
        'ContextMenuProjExplorer
        '
        Me.ContextMenuProjExplorer.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPEFloat, Me.mnuPEHide})
        '
        'mnuPEFloat
        '
        Me.mnuPEFloat.Index = 0
        Me.mnuPEFloat.Text = "&Float"
        '
        'mnuPEHide
        '
        Me.mnuPEHide.Index = 1
        Me.mnuPEHide.Text = "&Hide"
        '
        'ContextMenuModuleNode
        '
        Me.ContextMenuModuleNode.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuMNFullPath, Me.mnuMNOpen, Me.mnuMNClose, Me.MenuItem6, Me.mnuMNExclude, Me.MenuItem7, Me.mnuMNDelete, Me.mnuMNRename, Me.MenuItem9, Me.mnuMNProperties})
        '
        'mnuMNFullPath
        '
        Me.mnuMNFullPath.Index = 0
        Me.mnuMNFullPath.Text = "FullPath"
        Me.mnuMNFullPath.Visible = False
        '
        'mnuMNOpen
        '
        Me.mnuMNOpen.Index = 1
        Me.mnuMNOpen.Text = "&Open"
        '
        'mnuMNClose
        '
        Me.mnuMNClose.Index = 2
        Me.mnuMNClose.Text = "&Close"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 3
        Me.MenuItem6.Text = "-"
        '
        'mnuMNExclude
        '
        Me.mnuMNExclude.Index = 4
        Me.mnuMNExclude.Text = "Exclude from pro&ject"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 5
        Me.MenuItem7.Text = "-"
        '
        'mnuMNDelete
        '
        Me.mnuMNDelete.Index = 6
        Me.mnuMNDelete.Text = "&Delete"
        '
        'mnuMNRename
        '
        Me.mnuMNRename.Index = 7
        Me.mnuMNRename.Text = "Rena&me"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 8
        Me.MenuItem9.Text = "-"
        '
        'mnuMNProperties
        '
        Me.mnuMNProperties.Index = 9
        Me.mnuMNProperties.Text = "P&roperties"
        '
        'SplitterMainRight
        '
        Me.SplitterMainRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.SplitterMainRight.Location = New System.Drawing.Point(571, 29)
        Me.SplitterMainRight.Name = "SplitterMainRight"
        Me.SplitterMainRight.Size = New System.Drawing.Size(3, 464)
        Me.SplitterMainRight.TabIndex = 8
        Me.SplitterMainRight.TabStop = False
        '
        'OpenFileDialogModule
        '
        Me.OpenFileDialogModule.DefaultExt = "xc"
        Me.OpenFileDialogModule.Filter = "XC Module files (*.xc)|*.xc"
        Me.OpenFileDialogModule.Multiselect = True
        '
        'ContextMenuProjNode
        '
        Me.ContextMenuProjNode.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPNFullPath, Me.mnuPNCloseAllMod, Me.mnuPNCloseProj, Me.MenuItem18, Me.mnuPNAddNewMod, Me.mnuPNAddOldMod, Me.MenuItem19, Me.mnuPNSaveProj, Me.mnuPNRename, Me.MenuItem20, Me.mnuPNProperties})
        '
        'mnuPNFullPath
        '
        Me.mnuPNFullPath.Index = 0
        Me.mnuPNFullPath.Text = "FullPath"
        Me.mnuPNFullPath.Visible = False
        '
        'mnuPNCloseAllMod
        '
        Me.mnuPNCloseAllMod.Index = 1
        Me.mnuPNCloseAllMod.Text = "Close &All Modules"
        '
        'mnuPNCloseProj
        '
        Me.mnuPNCloseProj.Index = 2
        Me.mnuPNCloseProj.Text = "&Close Project"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 3
        Me.MenuItem18.Text = "-"
        '
        'mnuPNAddNewMod
        '
        Me.mnuPNAddNewMod.Index = 4
        Me.mnuPNAddNewMod.Text = "Add &New Module..."
        '
        'mnuPNAddOldMod
        '
        Me.mnuPNAddOldMod.Index = 5
        Me.mnuPNAddOldMod.Text = "Add Existin&g Module..."
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 6
        Me.MenuItem19.Text = "-"
        '
        'mnuPNSaveProj
        '
        Me.mnuPNSaveProj.Index = 7
        Me.mnuPNSaveProj.Text = "&Save Project"
        '
        'mnuPNRename
        '
        Me.mnuPNRename.Index = 8
        Me.mnuPNRename.Text = "Rena&me Project"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 9
        Me.MenuItem20.Text = "-"
        '
        'mnuPNProperties
        '
        Me.mnuPNProperties.Index = 10
        Me.mnuPNProperties.Text = "P&roperties"
        '
        'SaveFileDialogModule
        '
        Me.SaveFileDialogModule.DefaultExt = "xc"
        Me.SaveFileDialogModule.Filter = "XC Module files (*.xc)|*.xc"
        '
        'FormMain
        '
        Me.AllowDrop = True
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(824, 518)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.SplitterMainRight, Me.PanelMainRight, Me.ToolBarMain, Me.StatusBarMain})
        Me.IsMdiContainer = True
        Me.Menu = Me.MainMenu
        Me.Name = "FormMain"
        Me.Text = "XinCoder"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.PanelMainRight.ResumeLayout(False)
        Me.PanelProjExplorer.ResumeLayout(False)
        Me.PanelProjExplorerBody.ResumeLayout(False)
        Me.PanelProjExplorerHead.ResumeLayout(False)
        Me.PanelProperties.ResumeLayout(False)
        Me.PanelPropertiesBody.ResumeLayout(False)
        Me.GroupBoxPropertyDesc.ResumeLayout(False)
        Me.PanelPropertiesHead.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    ' Use by General IMS Website Project. 
    ' Such a project allows only one framework module.
    Private curFrameworkModuleName As String
    Private curFrameworkModuleNode As TreeNode
    ' Used by modules when creating output files, these are used to
    ' create links to the framework.
    Friend frameworkParams As New clsWebsiteFrameworkParameters()

    Private frmNewProject As FormNewProject
    Private frmNewModule As FormNewModule
    Private frmAbout As FormAbout
    Private frmOption As FormOption

    ' Keep track of openning modules by remembering Module File FullPath in Upper case.
    Friend openingModules As New ArrayList()
    Friend curProjectName As String
    Friend curProjectPath As String
    Private curProjNode As TreeNode
    Private curProjectType As ModuleMain.xcProjectType

    ' Use to edit property in Properties window.
    Dim propertyItem As ListViewItem ' current selected ListViewItem in Properties window
    Dim propertyUnderEditing As Boolean = False

    Private allInfoSaved As Boolean

    Private LabelTitleDefaultForeColor As Color
    Private LabelTitleDefaultBackColor As Color

    ' A module is internal if it locates inside the project folder (or subfolder), 
    ' and external if it locates outside the project folder.
    Const IS_INTERNAL_MODULE As String = "internal"
    Const IS_EXTERNAL_MODULE As String = "external"

    Private recentProjects As New ArrayList() ' Store recent projects
    Private MAX_RECENT_PROJECT_COUNT As Integer


    Private Sub FormMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'If SetupManager.readConfig = False Then
        'Me.Close()
        'Exit Sub
        'End If

        Me.setMainStatusBar("Thank you for using XinCoder")
        'Call Me.showFrmChildBase()
        'Call Me.createNewProjPrompt()
        Call Me.getTitleDefaultColor()
        Me.allInfoSaved = True
        Me.setEnableStatus_projMenu(False)

        Me.getRecentProjects()
        Me.showRecentProjects()
    End Sub


    ' Read recent projects from file
    Private Sub getRecentProjects()
        Dim projects As String()
        If IOManager.fileExists(SetupManager.RECENT_PROJECTS_FILE_PATH) Then
            IOManager.getFileAsLines(SetupManager.RECENT_PROJECTS_FILE_PATH, projects)
        End If
        recentProjects = MyUtil.stringArrayToArrayList(projects)
    End Sub


    ' Show recentProjects to File menu
    Private Sub showRecentProjects()
        Me.mnuRecentProj.MenuItems.Clear()

        Dim i, maxI As Integer
        maxI = recentProjects.Count
        For i = 1 To maxI
            Dim mnu As New MenuItem("&" & i & ". " & CType(recentProjects(i - 1), String))
            AddHandler mnu.Click, AddressOf openRecentProj
            Me.mnuRecentProj.MenuItems.Add(mnu)
        Next
    End Sub

    Private Sub openRecentProj(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim index As Integer = CType(sender, MenuItem).Index
        Me.openProjectFile(CType(Me.recentProjects(index), String))
    End Sub

    ' Add to recentProjects array when a new project is added or openned.
    Private Sub addRecentProjects(ByVal FullPath As String)
        Dim index As Integer = _
        MyUtil.arrayListContainsString(Me.recentProjects, FullPath, False)
        If index > -1 Then
            Me.recentProjects.RemoveAt(index)
        End If

        Me.recentProjects.Insert(0, FullPath)
        Dim countDiff As Integer = Me.recentProjects.Count - SetupManager.MaxRecentProjectsCount
        If countDiff > 0 Then
            Me.recentProjects.RemoveRange(4, countDiff)
        End If
    End Sub

    ' Write to file
    Private Sub writeRecentProjects()
        Dim str As String = MyUtil.arrayListToString(Me.recentProjects, vbCrLf)
        str = MyUtil.removeEndString(str, vbCrLf)
        IOManager.SaveTextToFile(str, SetupManager.RECENT_PROJECTS_FILE_PATH)
    End Sub


    Private Sub getTitleDefaultColor()
        Me.LabelTitleDefaultForeColor = Me.LabelProjExplorerTitle.ForeColor
        Me.LabelTitleDefaultBackColor = Me.LabelProjExplorerTitle.BackColor
    End Sub


    Friend Sub setCurProject(ByVal projName As String, ByVal projPath As String)
        Me.curProjectName = projName
        Me.curProjectPath = projPath
    End Sub


    Private Sub createNewProjPrompt()
        If frmNewProject Is Nothing Then
            frmNewProject = New FormNewProject()
        End If
        frmNewProject.ShowDialog(Me)
    End Sub


    ' Return false if a module with the given path is currently open or already exists.
    ' Return true only if such a module is not open and does not exist.
    Private Function validateNewModule(ByVal name As String, ByVal path As String, _
            Optional ByRef errInfo As xcValidateNewFileStatus = xcValidateNewFileStatus.noError) As Boolean

        If Me.openingModules.Contains(UCase(path)) Then
            errInfo = ModuleMain.xcValidateNewFileStatus.isOpening
            Return False
        ElseIf IOManager.fileExists(path) Then
            errInfo = ModuleMain.xcValidateNewFileStatus.alreadyExists
            Return False
        End If

        Return True
    End Function


    Friend Function createNewModule(ByVal modName As String, ByVal modPath As String, _
        ByVal modType As ModuleMain.xcModuleType, _
        Optional ByRef errInfo As xcValidateNewFileStatus = xcValidateNewFileStatus.noError)

        If validateNewModule(modName, modPath, errInfo) = False Then
            If errInfo = ModuleMain.xcValidateNewFileStatus.isOpening Then
                MsgBox("Module " & modPath & " is currently openned. ")
            ElseIf errInfo = ModuleMain.xcValidateNewFileStatus.alreadyExists Then
                MsgBox("Module " & modPath & " already exists. Please choose another name.")
            Else
                MsgBox("CreateNewModule Error: " & errInfo)
            End If
            Return False
        End If

        ' Validate if the module type is right.
        If Me.curProjectType = ModuleMain.xcProjectType.xcGeneralIMSWebsiteProject Then
            If Me.curFrameworkModuleName = "" Then
                ' No framework module yet
                If modType <> ModuleMain.xcModuleType.xcWebSiteFramework Then
                    MsgBox("For a General Website Project, a Framework module must be created before other modules.")
                    Return False
                End If
            Else    ' A framwork already exists.
                If modType = ModuleMain.xcModuleType.xcWebSiteFramework Then
                    MsgBox("A Framework module already exists. A General IMS Website Project can have only one Framework module.")
                    Return False
                End If
            End If
        End If

        Me.openModule(modType, modPath, modName)
        Me.saveProject(curProjectName, curProjectPath, curProjectType)
        Return True
    End Function


    Friend Sub openModule(ByVal modType As xcModuleType, _
                          ByVal modPath As String, ByVal modName As String, _
                                Optional ByVal openWindow As Boolean = True)

        Dim newModule As FormChildBase
        'Dim frmGenFW As New frmFramework()
        'frmGenFW.Show()

        'MsgBox("module type: " & projType.ToString)
        If modType = ModuleMain.xcModuleType.xcWebSiteFramework Then
            newModule = New FormGenFW()
            'newProject = New FormChildBase()
        ElseIf modType = ModuleMain.xcModuleType.xcLogin Then
            newModule = New FormGenLogin()
        ElseIf modType = ModuleMain.xcModuleType.xcCalendar Then
            newModule = New FormGenCalendar()
        ElseIf modType = ModuleMain.xcModuleType.xcDiscBoard Then
            newModule = New FormGenBBS()
        ElseIf modType = ModuleMain.xcModuleType.xcIMS Then
            newModule = New FormGenIMS()
        ElseIf modType = ModuleMain.xcModuleType.xcSearch Then
            newModule = New FormGenSearch()
        ElseIf modType = ModuleMain.xcModuleType.xcUpload Then
            newModule = New FormGenUpload()
        Else
            'newModule = New FormChildBase()
            MsgBox("This module type is not implemented yet.")
            Exit Sub
        End If

        ' Add node to Project Explorer if it is not in the project explorer.
        Dim modNode As TreeNode = getNodeInProjectExplorer(modPath)
        ' Set current Framework Module Node
        If modNode Is Nothing Then
            modNode = addModNodeToProjectExplorer(modPath, modName)

            If modType = ModuleMain.xcModuleType.xcWebSiteFramework Then
                Me.curFrameworkModuleName = modName
                Me.curFrameworkModuleNode = modNode
            End If
        End If
        Me.TreeViewProjExplorer.SelectedNode = modNode

        ' Used when first time load project, 
        ' keep the window status for each module to be the same as last time.
        If Not openWindow Then Exit Sub

        newModule.initProjParam(Me, modType, modPath, modName)
        newModule.Show()

        ' Add to openingModule list.
        Me.openingModules.Add(UCase(modPath))

        ' Enable module-related main menu items.
        Me.setEnableStatus_modMenu(True)
    End Sub


    ' Create a module node, add it to project explorer and return it.
    Private Function addModNodeToProjectExplorer(ByVal modPath As String, _
        ByVal modName As String) As TreeNode

        Dim moduleFileExt As String = SetupManager.MODULE_FILE_EXTENSION
        Dim modNode As TreeNode = New TreeNode(modName & moduleFileExt)

        modNode.ImageIndex = 1
        modNode.SelectedImageIndex = 1
        modNode.Tag = modPath
        Me.curProjNode.Nodes.Add(modNode)
        Me.curProjNode.Expand()

        Return modNode
    End Function


    ' Determine if a node with given path exists in project explorer
    Private Function getNodeInProjectExplorer(ByVal modPath As String) As TreeNode
        Dim node As TreeNode
        For Each node In Me.curProjNode.Nodes
            If UCase(CType(node.Tag, String)) = UCase(modPath) Then
                Return node
            End If
        Next

        Return Nothing
    End Function


    ' For testing FormChildBase class only.
    Private Sub showFrmChildBase()
        Dim frmChildBase As New FormChildBase()
        frmChildBase.initProjParam(Me, ModuleMain.xcModuleType.xcTemplateBase, _
                                   "frmChildBase", SetupManager.OUTPUT_PATH)
        frmChildBase.Show()
        Me.openingModules.Add(UCase(SetupManager.OUTPUT_PATH))
        Me.setEnableStatus_modMenu(True)
    End Sub


    ' Enable or disable menu items related to module.
    Private Sub setEnableStatus_modMenu(ByVal status As Boolean)
        Me.resetModuleRelatedMenu()
        ' File menu
        'Me.mnuSaveModule.Enabled = status
        'Me.mnuSaveModuleAs.Enabled = status

        ' Close menu
        'Me.mnuCloseMod.Enabled = status
        Me.mnuCloseAllMod.Enabled = status

        ' View menu
        Me.mnuViewStep.Enabled = status
        Me.mnuViewDisp.Enabled = status
        Me.mnuViewInstr.Enabled = status

        ' Window menu
        Me.mnuWinArrange.Enabled = status
        Me.mnuWinCascade.Enabled = status
        Me.mnuWinTileH.Enabled = status
        Me.mnuWinTileV.Enabled = status
    End Sub

    Private Sub setEnableStatus_singleModMenu(ByVal status As Boolean)
        ' File Menu
        Me.mnuSaveModule.Enabled = status
        Me.mnuSaveModuleAs.Enabled = status

        ' Close menu
        Me.mnuCloseMod.Enabled = status

        ' Build Menu
        Me.mnuBuildModule.Enabled = status
    End Sub


    ' Return false if a project with the given path is currently open or already exists.
    ' Return true otherwise.
    Private Function validateNewProject(ByVal projName As String, ByVal projPath As String, _
            Optional ByRef errInfo As xcValidateNewFileStatus = xcValidateNewFileStatus.noError) As Boolean
        'MsgBox("curProj: " & Me.curProjectPath & vbCrLf & "new Proj: " & projPath)
        If UCase(Me.curProjectPath) = UCase(projPath) Then
            errInfo = ModuleMain.xcValidateNewFileStatus.isOpening
            Return False
        ElseIf IOManager.folderExists(projPath) Then
            errInfo = ModuleMain.xcValidateNewFileStatus.alreadyExists
            Return False
        End If

        Return True
    End Function


    Friend Function createNewProject(ByVal projectName As String, ByVal projectPath As String, _
        ByVal projectType As ModuleMain.xcProjectType) As Boolean

        Dim errinfo As ModuleMain.xcValidateNewFileStatus

        If validateNewProject(projectName, projectPath, errInfo) = False Then
            If errInfo = ModuleMain.xcValidateNewFileStatus.isOpening Then
                MessageBox.Show("Project " & projectName & " in folder " & projectPath & _
                       " is currently openned. ", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            ElseIf errInfo = ModuleMain.xcValidateNewFileStatus.alreadyExists Then
                'MsgBox("Folder " & projPath & " already exists. Please choose another project name.")
                'Return False
                If MessageBox.Show("Folder " & projectPath & " already exists. Do you want to overwrite this folder?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then Return False
                If MessageBox.Show("Current contents in " & projectPath & " will be deleted. Do you want to proceed?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then Return False
                IOManager.DeleteFolder(projectPath, True)
            End If
        End If

        If Me.curProjectPath <> "" Then
            Me.saveProject(Me.curProjectName, Me.curProjectPath, Me.curProjectType)
            Me.closeCurrentProject()
        End If

        ' Assign new curProjectName and curProjectPath variables.
        Me.curProjectName = projectName
        Me.curProjectPath = projectPath
        Me.curProjectType = projectType

        ' Create folders 
        IOManager.CreateFolder(curProjectPath) ' project folder, contains project files
        IOManager.CreateFolder(getProjectOutputPath())  ' project output folder
        ' Create project file
        Me.saveProject(curProjectName, curProjectPath, Me.curProjectType)

        ' Add node to Project Explorer
        Me.addProjNodeToProjectExplorer(Me.curProjectName, Me.curProjectPath)

        Me.setEnableStatus_projMenu(True)

        Me.addRecentProjects(Me.curProjectPath & Me.curProjectName & SetupManager.PROJ_FILE_EXTENSION)
        Me.showRecentProjects()

        ' If is a Website project, create a Framework module.
        If projectType = ModuleMain.xcProjectType.xcGeneralIMSWebsiteProject Then
            Me.createNewModule(xcUtil.WebsiteFrameworkModuleName, _
                Me.getCurFrameworkModulePath(), _
                ModuleMain.xcModuleType.xcWebSiteFramework)
        End If

        Return True
    End Function


    Private Function getCurFrameworkModulePath() As String
        Return curProjectPath & Me.curFrameworkModuleName & SetupManager.MODULE_FILE_EXTENSION
    End Function


    Friend Function getProjectOutputPath() As String
        Dim outputPath As String = Me.curProjectPath & SetupManager.OUTPUT_FOLDER
        If Not outputPath.EndsWith("\") Then
            outputPath += "\"
        End If
        ''''' tmp output path
        outputPath = "c:\inetpub\wwwroot\xcOutput\"
        '''''
        Return outputPath
    End Function

    ' Note: projPath always ends with "\"
    Private Sub saveProject(ByVal projName As String, _
        ByVal projPath As String, ByVal projType As String)

        ' Update framework parameters if the Framework module is open.
        Me.updateFrameworkParams()

        Dim projFileExt As String = SetupManager.PROJ_FILE_EXTENSION
        Dim projFile As String = projPath & projName & projFileExt

        Dim str As String = ""
        str = str & "CreateDate=" & Now() & vbCrLf
        str = str & "ProjectName=" & projName & vbCrLf
        str = str & "ProjectType=" & projType & vbCrLf
        str = str & vbCrLf
        'str = str & "ProjectFrameworkModuleName=" & Me.curFrameworkModuleName & vbCrLf
        str = str & vbCrLf
        str = str & Me.frameworkParams.writeParams("Project")
        str = str & vbCrLf

        ' Save module information
        If Not Me.curProjNode Is Nothing Then
            Dim node As TreeNode
            For Each node In Me.curProjNode.Nodes
                str = str & "ModuleName=" & IOManager.formatFileName(node.Text, _
                      SetupManager.MODULE_FILE_EXTENSION, False) & vbCrLf
                'str = str & "ModuleFullPath=" & CType(node.Tag, String) & vbCrLf

                Dim modPath As String = CType(node.Tag, String)
                If UCase(modPath).StartsWith(UCase(Me.curProjectPath)) Then
                    str = str + "ModuleSource=" & Me.IS_INTERNAL_MODULE & vbCrLf
                    str = str + "RelativePath=" & modPath.Substring(Len(Me.curProjectPath)) & vbCrLf
                Else
                    str = str + "ModuleSource=" & Me.IS_EXTERNAL_MODULE & vbCrLf
                    str = str + "AbsolutePath=" & modPath & vbCrLf
                End If

                If Me.openingModules.Contains(UCase(modPath)) Then
                    str = str + "ModuleWindowIsOpen=True" & vbCrLf
                Else
                    str = str + "ModuleWindowIsOpen=False" & vbCrLf
                End If

                str = str & vbCrLf
            Next
        End If

        IOManager.SaveTextToFile(str, projFile)

        ' save opening project modules information
        Dim mdiChild As FormChildBase
        For Each mdiChild In Me.MdiChildren
            mdiChild.saveModule()
        Next

        ' Make the interface obvious and friendly.
        Me.showWaitCursor()
    End Sub


    ' Change the cursor to let the user know Saving activity is happening.
    Private Sub showWaitCursor()
        Me.Cursor = Cursors.WaitCursor
        Threading.Thread.Sleep(200) ' sleep 0.2 second to show the WaitCursor.
        Me.Cursor = Cursors.Default
    End Sub


    ' set main menu status related to project
    Private Sub setEnableStatus_projMenu(ByVal status As Boolean)
        Me.mnuAddNewMod.Enabled = status
        Me.mnuAddOldMod.Enabled = status
        Me.mnuCloseProj.Enabled = status
        Me.mnuSaveProj.Enabled = status
        Me.mnuBuildProject.Enabled = status
        Me.ToolBarButtonNewModule.Enabled = status
        Me.ToolBarButtonExistModule.Enabled = status
        Me.ToolBarButtonSaveProject.Enabled = status
        Me.ToolBarButtonDeleteModule.Enabled = status
        Me.ToolBarButtonBuildProject.Enabled = status
        Me.ToolBarButtonCloseProject.Enabled = status
    End Sub


    ' Bring the selected module window to the front.
    Friend Sub mnuWindowDynamicMenuItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWindow.Click
        Dim mnuItem As MenuItem = CType(sender, MenuItem)

        Dim mdiChild As FormChildBase
        For Each mdiChild In Me.MdiChildren
            If mdiChild.windowMenuItem Is mnuItem Then
                mdiChild.Focus()
                mdiChild.BringToFront()
                Exit For
            End If
        Next
    End Sub


    ' Event when a child is activated.
    Private Sub FormMain_MdiChildActivate(ByVal sender As Object, ByVal e As _
               System.EventArgs) Handles MyBase.MdiChildActivate
        Dim activeChild As FormChildBase = CType(Me.ActiveMdiChild, FormChildBase)

        ' Check corresponding item in Window menu
        Dim mnuItem As MenuItem
        For Each mnuItem In Me.mnuWindow.MenuItems
            mnuItem.Checked = False
        Next
        If (Not Me.ActiveMdiChild Is Nothing) Then
            activeChild.windowMenuItem.Checked = True
        End If

        ' Change the corresponding selected node in project explorer.
        If (Not Me.curProjNode Is Nothing) And (Not activeChild Is Nothing) Then
            Dim node As TreeNode
            For Each node In Me.curProjNode.Nodes
                If UCase(node.Text) = UCase(activeChild.moduleName & _
                                      SetupManager.MODULE_FILE_EXTENSION) Then
                    Me.TreeViewProjExplorer.SelectedNode = node
                    Me.resetProjectExplorerNodesBackColor()
                End If
            Next
        End If
    End Sub


    Public Sub setMainStatusBar(ByVal msg As String)
        Me.StatusBarMain.Text = "  " & msg
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Menu Click functions
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' File menu

    ' File -> New -> Project
    Private Sub mnuNewProj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewProj.Click
        Me.createNewProjPrompt()
    End Sub

    ' File -> Open -> Project
    Private Sub mnuOpenProj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpenProj.Click
        Me.openProjPrompt()
    End Sub

    ' File -> Add New Module...
    Private Sub mnuAddNewMod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddNewMod.Click
        addNewModPrompt()
    End Sub

    Private Sub addNewModPrompt()
        If Me.frmNewModule Is Nothing Then
            Me.frmNewModule = New FormNewModule()
        End If
        Me.frmNewModule.ShowDialog(Me)
    End Sub

    ' File -> Add Existing Module...
    Private Sub mnuAddOldMod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddOldMod.Click
        addExistingModPrompt()
    End Sub

    Private Sub addExistingModPrompt()
        Me.OpenFileDialogModule.FileName = ""
        'Me.OpenFileDialogModule.InitialDirectory = SetupManager.OUTPUT_PATH
        Me.OpenFileDialogModule.InitialDirectory = Me.curProjectPath
        Me.OpenFileDialogModule.ShowDialog(Me)
    End Sub

    Private Sub OpenFileDialogModule_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialogModule.FileOk
        Dim filePath As String
        For Each filePath In Me.OpenFileDialogModule.FileNames
            If filePath.EndsWith(SetupManager.MODULE_FILE_EXTENSION) Then
                Me.openModuleFile(filePath, "", True)
            End If
        Next
    End Sub

    ' File -> Close Module
    Private Sub mnuClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCloseMod.Click
        closeMdiChild(Me.ActiveMdiChild)
    End Sub

    ' File -> Close All Modules
    Private Sub mnuCloseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCloseAllMod.Click
        closeAllMdiChild()
    End Sub

    Private Sub closeAllMdiChild()
        Dim mdiChild As Form
        For Each mdiChild In Me.MdiChildren
            closeMdiChild(mdiChild)
        Next
    End Sub

    ' Close mdiChild, used by FormMain.vb and FormChildBase.vb
    Friend Sub closeMdiChild(ByRef mdiChild As FormChildBase)
        If Not mdiChild Is Nothing Then
            ' Change the index number of dynamic items in mnuWindow.
            Dim index As Integer = Me.mnuWindow.MenuItems.IndexOf(mdiChild.windowMenuItem)
            Dim i
            For i = index + 1 To Me.mnuWindow.MenuItems.Count - 1
                Dim str As String = Me.mnuWindow.MenuItems.Item(i).Text
                str = str.Substring(1)
                str = CStr(i - 5) & str
                Me.mnuWindow.MenuItems.Item(i).Text = str
            Next

            Me.mnuWindow.MenuItems.Remove(mdiChild.windowMenuItem)
            Me.openingModules.Remove(UCase(mdiChild.moduleFullPath)) '''

            If mdiChild.infoSaved = False Then
                If MessageBox.Show("Do you want to save changes to this module?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    mdiChild.saveModule()
                End If
            End If

            mdiChild.Close()
        End If

        If Me.MdiChildren.Length = 0 Then
            'Me.mnuCloseMod.Text = "&Close Module"
            'Me.mnuSaveModule.Text = "&Save Module"
            'Me.mnuSaveModuleAs.Text = "Save Module &As ..."
            'Me.mnuBuildModule.Text = "B&uild Module"
            Me.setMenuText("")
            Me.setEnableStatus_modMenu(False)
            Me.setEnableStatus_singleModMenu(False)
        End If

    End Sub


    ' if moduleName is not empty, add it to menu text
    Private Sub setMenuText(ByVal moduleName As String)
        If moduleName = "" Then
            Me.mnuCloseMod.Text = "&Close Module"
            Me.mnuSaveModule.Text = "&Save Module"
            Me.mnuSaveModuleAs.Text = "Save Module &As ..."
            Me.mnuBuildModule.Text = "B&uild Module"
        Else
            Me.mnuCloseMod.Text = "&Close Module " & moduleName
            Me.mnuSaveModule.Text = "&Save Module " & moduleName
            Me.mnuSaveModuleAs.Text = "Save Module " & moduleName & " &As ..."
            Me.mnuBuildModule.Text = "B&uild Module " & moduleName
        End If
    End Sub


    ' File -> Close Project
    Private Sub mnuCloseProj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCloseProj.Click
        closeCurrentProject()
    End Sub

    Private Sub closeCurrentProject()
        'If Me.allInfoSaved = False Then
        If MessageBox.Show("Do you want to save all the work on current project?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Me.saveProject(Me.curProjectName, Me.curProjectPath, Me.curProjectType)
        End If
        'End If
        Me.closeAllMdiChild()
        Me.curProjectName = ""
        Me.curProjectPath = ""
        If Not Me.curProjNode Is Nothing Then Me.curProjNode.Nodes.Clear()
        Me.curProjNode = Nothing
        Me.TreeViewProjExplorer.Nodes.Clear()

        Me.setEnableStatus_projMenu(False)
        Me.refreshPropertyWindow()
    End Sub

    ' File -> Exit
    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub


    ' Functions for menu View

    Private Sub mnuViewStep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewStep.Click
        If Not Me.ActiveMdiChild Is Nothing Then
            CType(Me.ActiveMdiChild, FormChildBase).viewStepExplorer()
        End If
    End Sub


    Private Sub mnuViewDisp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewDisp.Click
        If Not Me.ActiveMdiChild Is Nothing Then
            CType(Me.ActiveMdiChild, FormChildBase).viewDispExplorer()
        End If
    End Sub


    Private Sub mnuViewInstr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewInstr.Click
        If Not Me.ActiveMdiChild Is Nothing Then
            CType(Me.ActiveMdiChild, FormChildBase).viewInstrExplorer()
        End If
    End Sub


    ' Functions for menu Help

    Private Sub mnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
        If frmAbout Is Nothing Then
            frmAbout = New FormAbout()
        End If
        frmAbout.ShowDialog(Me)
    End Sub


    Private Sub mnuHelpContent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpContent.Click
        Help.ShowHelp(Me, SetupManager.HELP_FILE)
    End Sub


    Private Sub mnuHelpIndex_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpIndex.Click
        Help.ShowHelpIndex(Me, SetupManager.HELP_FILE)
    End Sub


    ' Functions for menu Window

    Private Sub mnuWinArrange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWinArrange.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub mnuWinCascade_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWinCascade.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub mnuWinTileH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWinTileH.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub mnuWinTileV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWinTileV.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub


    ' Tools menu

    Private Sub mnuOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOptions.Click
        If frmOption Is Nothing Then
            frmOption = New FormOption()
        End If
        frmOption.ShowDialog()
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions for Tool bar buttons
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub ToolBarMain_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBarMain.ButtonClick
        If e.Button Is Me.ToolBarButtonNewProject Then
            Me.createNewProjPrompt()
        ElseIf e.Button Is Me.ToolBarButtonOpenProject Then
            Me.openProjPrompt()
        ElseIf e.Button Is Me.ToolBarButtonCloseProject Then
            Me.closeCurrentProject()

        ElseIf e.Button Is Me.ToolBarButtonNewModule Then
            Me.addNewModPrompt()
        ElseIf e.Button Is Me.ToolBarButtonExistModule Then
            addExistingModPrompt()
        ElseIf e.Button Is Me.ToolBarButtonDeleteModule Then
            Me.deleteModuleFile()

        ElseIf e.Button Is Me.ToolBarButtonSaveProject Then
            'Me.saveProjPrompt()
            Me.saveProject(Me.curProjectName, Me.curProjectPath, Me.curProjectType)
        ElseIf e.Button Is Me.ToolBarButtonBuildProject Then
            Me.buildProject()
        End If
    End Sub


    Private Sub openProjPrompt()
        Dim projExt As String = SetupManager.PROJ_FILE_EXTENSION
        Me.OpenFileDialogProj.DefaultExt = projExt
        Me.OpenFileDialogProj.AddExtension = True

        '"XC Project files (*.xcproj|*.xcproj|All files (*.*)|*.*"
        Me.OpenFileDialogProj.Filter = "XC Project files (*" & projExt & ")|*" & projExt & "|All files (*.*)|*.*"
        Me.OpenFileDialogProj.FilterIndex = 0

        Me.OpenFileDialogProj.InitialDirectory = SetupManager.OUTPUT_PATH
        Me.OpenFileDialogProj.ShowDialog()
    End Sub


    Private Sub saveProjPrompt()
        Me.SaveFileDialogMain.FileName = Me.curProjectName
        Me.SaveFileDialogMain.InitialDirectory = Me.curProjectPath
        Me.SaveFileDialogMain.ShowDialog()
    End Sub


    Private Sub OpenFileDialogProj_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialogProj.FileOk
        Dim projName As String = Me.OpenFileDialogProj.FileName
        If projName.EndsWith(SetupManager.PROJ_FILE_EXTENSION) Then
            openProjectFile(Me.OpenFileDialogProj.FileName)
        End If
    End Sub


    Private Sub SaveFileDialogMain_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialogMain.FileOk
        Dim curChild As FormChildBase = CType(Me.ActiveMdiChild, FormChildBase)
    End Sub


    Private Function addProjNodeToProjectExplorer(ByVal projName As String, _
        ByVal projPath As String) As TreeNode

        curProjNode = New TreeNode(projName)
        curProjNode.ImageIndex = 0
        curProjNode.SelectedImageIndex = 0
        curProjNode.Tag = projPath
        curProjNode.Expand()

        Me.TreeViewProjExplorer.Nodes.Clear()
        Me.TreeViewProjExplorer.Nodes.Add(curProjNode)

        Return curProjNode
    End Function


    Private Sub openProjectFile(ByVal projFilePath As String)
        Dim projExt As String = SetupManager.PROJ_FILE_EXTENSION

        'MsgBox("projfile to open: " & projFilePath & vbCrLf & "cur proj: " & Me.curProjectPath & Me.curProjectName & ".xcproj")
        If UCase(projFilePath) = UCase(Me.curProjectPath & _
            Me.curProjectName & projExt) Then
            MsgBox("This project is currently open.")
            Exit Sub
        End If

        'MsgBox("projName=" & projectName & ", projpath=" & projectPath)
        If Me.curProjectPath <> "" Then Me.closeCurrentProject()

        ' Read project file
        Dim errInfo As String
        Dim lines() As String

        If IOManager.getFileAsLines(projFilePath, lines, errInfo) = False Then
            MsgBox("Error: cannot open project " & projFilePath & vbCrLf & errInfo, MsgBoxStyle.Critical)
            Exit Sub
        End If

        Dim projNameIsRead As Boolean = False
        Dim projTypeIsRead As Boolean = False
        'Dim projFWNameIsRead As Boolean = False

        Dim i As Integer
        Dim maxI As Integer = lines.Length - 1

        Dim line, lcase_line As String
        For i = 0 To maxI
            line = lines(i)
            lcase_line = LCase(line)
            If lcase_line.StartsWith("projectname") Then
                Me.curProjectName = MyUtil.getStrAfterDelimiter(line, "=")
                projNameIsRead = True
            ElseIf lcase_line.StartsWith("projecttype") Then
                Me.curProjectType = CType(MyUtil.getStrAfterDelimiter(line, "="), ModuleMain.xcProjectType)
                projTypeIsRead = True
                'ElseIf lcase_line.StartsWith("projectframeworkmodulename") Then
                '    Me.curFrameworkModuleName = MyUtil.getStrAfterDelimiter(line, "=")
                '    projFWNameIsRead = True
            End If
            'If projNameIsRead And projTypeIsRead And projFWNameIsRead Then Exit For
            If projNameIsRead And projTypeIsRead Then Exit For
        Next

        ' Get framework parameters from the file.
        Me.frameworkParams.readParams(lines)

        ' Finish read file

        Me.curProjectPath = IOManager.getFolder(projFilePath)

        ' Add project node to Project Explorer
        Me.addProjNodeToProjectExplorer(Me.curProjectName, Me.curProjectPath)

        Me.setEnableStatus_projMenu(True)

        ' Open modules
        ' The advantage of using IS_INTERNAL_MODULE + relativePath and
        ' IS_EXTERNAL_MODULE + absolutePath is that you can move the project folder
        ' around and still can open the project modules without problem.
        Dim modSource As String
        Dim pathLine As String
        Dim moduleWindowIsOpen As Boolean
        For i = 0 To maxI
            line = LCase(lines(i))

            If line.StartsWith("modulesource") Then
                modSource = LCase(MyUtil.getStrAfterDelimiter(line, "="))
                pathLine = lines(i + 1)
                moduleWindowIsOpen = CType(MyUtil.getStrAfterDelimiter(lines(i + 2), "="), Boolean)

                If modSource = Me.IS_INTERNAL_MODULE Then
                    Me.openModuleFile(Me.curProjectPath & _
                       MyUtil.getStrAfterDelimiter(pathLine, "="), "", moduleWindowIsOpen)
                ElseIf modSource = Me.IS_EXTERNAL_MODULE Then  ' external
                    Me.openModuleFile(MyUtil.getStrAfterDelimiter(pathLine, "="), "", moduleWindowIsOpen)
                End If
                i = i + 2
            End If
        Next

        Me.addRecentProjects(Me.curProjectPath & Me.curProjectName & SetupManager.PROJ_FILE_EXTENSION)
        Me.showRecentProjects()
    End Sub


    ' Return the mdiChild with given FilePath in current mdiChildren collection
    Private Function getMdiChild(ByVal FilePath As String) As FormChildBase
        Dim mdiChild As FormChildBase
        For Each mdiChild In Me.MdiChildren
            If UCase(mdiChild.moduleFullPath) = UCase(FilePath) Then
                Return mdiChild
            End If
        Next
        Return Nothing
    End Function


    ' Parameters:
    ' - openWindow 
    ' used when first open a project, to remember whether the module is open last time.
    ' - Filename 
    ' used after a module is renamed, the new name needs to be passed to the module.
    Private Sub openModuleFile(ByVal FilePath As String, _
        ByVal newModName As String, ByVal openWindow As Boolean)
        If Me.openingModules.Contains(UCase(FilePath)) Then
            'MsgBox("This module is currently open")
            Me.getMdiChild(FilePath).BringToFront()
            Exit Sub
        End If

        ' Read file
        Dim createDate As Date
        Dim moduleName As String
        Dim modulePath As String = FilePath
        Dim moduleType As ModuleMain.xcModuleType

        ' Read module file
        Dim errInfo As String
        Dim lines() As String

        If IOManager.getFileAsLines(FilePath, lines, errInfo) = False Then
            'MsgBox("Error: cannot open module " & FilePath & vbCrLf & errInfo, MsgBoxStyle.Critical)
            MsgBox("Error: " & vbCrLf & errInfo, MsgBoxStyle.Critical)
            Exit Sub
        End If

        Dim i As Integer
        Dim maxI As Integer = lines.Length - 1

        Dim line, lcase_line As String
        For i = 0 To maxI
            line = lines(i)
            lcase_line = LCase(line)
            If lcase_line.StartsWith("createdate") Then
                createDate = CType(MyUtil.getStrAfterDelimiter(line, "="), Date)
            ElseIf lcase_line.StartsWith("modulename") Then
                moduleName = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("moduletype") Then
                moduleType = CType(MyUtil.getStrAfterDelimiter(line, "="), ModuleMain.xcModuleType)
            End If
        Next
        ' Finish read file

        If newModName <> "" Then moduleName = newModName ' happens only after rename a module
        Me.openModule(moduleType, modulePath, moduleName, openWindow)
    End Sub


    Private Sub mnuViewProjExplorer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewProjExplorer.Click
        Me.PanelMainRight.Visible = True
        Me.PanelProjExplorer.Visible = True
        Me.PanelProperties.Dock = DockStyle.Bottom
        Me.setCurExplorer(Me.LabelProjExplorerTitle)
    End Sub


    Private Sub mnuViewProperties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewProperties.Click
        Me.PanelMainRight.Visible = True
        Me.PanelProperties.Visible = True
        Me.setCurExplorer(Me.LabelPropertiesTitle)
    End Sub

    ' Close Project Explorer
    Private Sub btnCloseProjExplorer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseProjExplorer.Click
        Me.closeProjExplorer()
    End Sub

    Private Sub closeProjExplorer()
        Me.PanelProjExplorer.Visible = False
        Me.PanelProperties.Dock = DockStyle.Fill
        If Me.PanelProjExplorer.Visible = False And Me.PanelProperties.Visible = False Then
            Me.PanelMainRight.Visible = False
        End If
    End Sub

    ' Close Properties window.
    Private Sub btnCloseProperties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseProperties.Click
        Me.closePropertiesWindow()
    End Sub

    Private Sub closePropertiesWindow()
        Me.PanelProperties.Visible = False
        If Me.PanelProjExplorer.Visible = False And Me.PanelProperties.Visible = False Then
            Me.PanelMainRight.Visible = False
        End If
    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions to Edit Property in Properties window.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private prevTextBoxPropertyValue As String
    Private Sub ListViewProperties_MouseUp(ByVal sender As System.Object, ByVal e As MouseEventArgs) Handles ListViewProperties.MouseUp
        'MsgBox("clicked: " & e.X & ", " & e.Y) ' & vbCrLf & " - " & _
        propertyItem = ListViewProperties.GetItemAt(e.X, e.Y)
        Dim x, y, w, h As Integer

        If Not propertyItem Is Nothing Then
            ' Full Path cannot be editted.
            'If LCase(propertyItem.SubItems.Item(0).Text) <> "full path" Then
            If CType(propertyItem.Tag, Boolean) = True Then
                x = propertyItem.GetBounds(ItemBoundsPortion.Entire).X
                y = propertyItem.GetBounds(ItemBoundsPortion.Entire).Y
                w = propertyItem.GetBounds(ItemBoundsPortion.Entire).Width
                h = propertyItem.GetBounds(ItemBoundsPortion.Entire).Height
                'MsgBox(Me.ListViewProperties.GetItemAt(e.X, e.Y).SubItems.Item(1).Text())
                'MsgBox("X=" & Me.ListViewProperties.GetItemAt(e.X, e.Y).GetBounds(ItemBoundsPortion.ItemOnly).X & ", Y=" & Me.ListViewProperties.GetItemAt(e.X, e.Y).GetBounds(ItemBoundsPortion.ItemOnly).Y & ", W=" & Me.ListViewProperties.GetItemAt(e.X, e.Y).GetBounds(ItemBoundsPortion.ItemOnly).Width & ", H=" & Me.ListViewProperties.GetItemAt(e.X, e.Y).GetBounds(ItemBoundsPortion.ItemOnly).Height)
                Me.TextBoxProperty.Text = Me.ListViewProperties.GetItemAt(e.X, e.Y).SubItems.Item(1).Text()
                prevTextBoxPropertyValue = Me.TextBoxProperty.Text

                Me.TextBoxProperty.Left = x + Me.ListViewProperties.Columns.Item(0).Width + 3 '+ Me.ListViewProperties.Left
                Me.TextBoxProperty.Top = y + 2 '+ Me.ListViewProperties.Top
                Me.TextBoxProperty.Width = w - Me.ListViewProperties.Columns.Item(0).Width - 1
                'Me.TextBoxProperty.Height = h ' Height is decided by font size
                Me.TextBoxProperty.Visible = True
                Me.TextBoxProperty.BorderStyle = BorderStyle.None
                Me.TextBoxProperty.BringToFront()
                propertyUnderEditing = True
            End If
        Else
            Me.refreshPropertyWindow()
        End If

        Me.setCurExplorer(Me.LabelPropertiesTitle)
    End Sub


    Private Sub TextBoxProperty_enter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxProperty.KeyPress
        'If e.KeyCode = Keys.Enter Then  ' This produces a sound on Enter
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            endEditProperty()
        End If
    End Sub


    Private Sub ListViewProperties_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewProperties.SelectedIndexChanged
        endEditProperty()
    End Sub


    Private Sub endEditProperty()
        If propertyUnderEditing = False Or propertyItem Is Nothing Then
            Me.TextBoxProperty.Visible = False
            Exit Sub
        End If

        ' Get the node to which this property belongs to.
        Dim node As TreeNode = Me.TreeViewProjExplorer.SelectedNode
        If node Is Nothing Then Exit Sub

        Dim propertyName As String = propertyItem.SubItems.Item(0).Text
        Dim newPropertyValue As String = Trim(Me.TextBoxProperty.Text)

        Dim editResult As Boolean

        If node Is Me.curProjNode And propertyName = Me.PROPERTY_TITLE_PROJ_FILE Then
            editResult = editProjectName(node, propertyItem, newPropertyValue)
        ElseIf propertyName = Me.PROPERTY_TITLE_FILE_NAME Then
            editResult = editModuleName(node, propertyItem, newPropertyValue)
        End If

        If editResult = True Then
            Me.TextBoxProperty.Visible = False
            Me.propertyUnderEditing = False
        Else

        End If
    End Sub

    Private Function editProjectName(ByRef node As TreeNode, _
            ByRef propertyItem As ListViewItem, ByVal newValue As String) As Boolean
        Dim oldProjName As String = node.Text
        Dim newProjName As String = newValue
        Dim oldProjPath As String = CType(node.Tag, String)
        Dim newProjPath As String = oldProjPath

        If newProjName = oldProjName Then Return True
        If newProjName = "" Then
            MsgBox("Project name cannot be empty")
            Return False
        End If
        If Not newProjName.EndsWith(SetupManager.PROJ_FILE_EXTENSION) Then
            MsgBox("Project name must end with " & SetupManager.PROJ_FILE_EXTENSION)
            Return False
        End If

        newProjName = MyUtil.removeEndString(newValue, SetupManager.PROJ_FILE_EXTENSION)
        Me.renameProject(node, oldProjPath, oldProjName, newProjName)
        Return True
    End Function

    Private Function editModuleName(ByRef node As TreeNode, _
            ByRef propertyItem As ListViewItem, ByVal newValue As String) As Boolean
        Dim oldModuleName As String = node.Text
        Dim newModuleName As String = newValue
        Dim oldModulePath As String = CType(node.Tag, String)
        Dim newModulePath As String = IOManager.getFolder(oldModulePath) & newModuleName

        If newModuleName = oldModuleName Then Return True
        If newModuleName = "" Then
            MsgBox("Module name cannot be empty")
            Return False
        End If
        If Not newModuleName.EndsWith(SetupManager.MODULE_FILE_EXTENSION) Then
            MsgBox("Module name must end with " & SetupManager.MODULE_FILE_EXTENSION)
            Return False
        End If
        If IOManager.fileExists(newModulePath) Then
            MsgBox("File " & newModulePath & " already exists.")
            Return False
        End If

        Me.renameModule(node, oldModuleName, oldModulePath, newModuleName, newModulePath, True)
        Return True
    End Function


    Private Function validatePropertyValue_Filename(ByRef node As TreeNode, _
        ByVal fileFolder As String, ByVal value As String) As Boolean
        If value = "" Then
            MsgBox("You must enter a name.")
            Return False
        End If

        If node Is Me.curProjNode Then
            If Not value.EndsWith(SetupManager.PROJ_FILE_EXTENSION) Then
                MsgBox("Project name must end with " & SetupManager.PROJ_FILE_EXTENSION)
                Return False
            End If
        Else
            If Not value.EndsWith(SetupManager.MODULE_FILE_EXTENSION) Then
                MsgBox("Module name must end with " & SetupManager.MODULE_FILE_EXTENSION)
                Return False
            ElseIf IOManager.fileExists(fileFolder & value) Then
                MsgBox("File " & fileFolder & value & " already exists.")
                Return False
            End If
        End If

        Return True
    End Function


    Private Sub ListViewProperties_resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewProperties.Resize
        Me.ListViewProperties.Columns.Item(1).Width = Me.ListViewProperties.Width - Me.ListViewProperties.Columns.Item(0).Width - 5
        If propertyUnderEditing = True Then
            Me.TextBoxProperty.Left = propertyItem.GetBounds(ItemBoundsPortion.Entire).X + Me.ListViewProperties.Columns.Item(0).Width + 3
            Me.TextBoxProperty.Width = propertyItem.GetBounds(ItemBoundsPortion.Entire).Width - Me.ListViewProperties.Columns.Item(0).Width - 1
        End If
    End Sub


    Private Sub resumeTitleColor()
        Me.LabelProjExplorerTitle.ForeColor = Me.LabelTitleDefaultForeColor
        Me.LabelProjExplorerTitle.BackColor = Me.LabelTitleDefaultBackColor

        Me.LabelPropertiesTitle.ForeColor = Me.LabelTitleDefaultForeColor
        Me.LabelPropertiesTitle.BackColor = Me.LabelTitleDefaultBackColor
    End Sub


    Private Sub LabelProjExplorerTitle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelProjExplorerTitle.Click
        Me.setCurExplorer(Me.LabelProjExplorerTitle)
    End Sub

    ' Set current explorer's title color to be highlighted.
    Private Sub setCurExplorer(ByRef focusedPanelTitle As Label)
        Me.resumeTitleColor()
        focusedPanelTitle.ForeColor = Color.White
        focusedPanelTitle.BackColor = Color.Blue
    End Sub


    Private Overloads Sub TreeViewProjExplorer_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeViewProjExplorer.AfterSelect
        Me.TextBoxProperty.Visible = False
        Me.resetProjectExplorerNodesBackColor()
        Me.refreshPropertyWindow()
        Me.resetModuleRelatedMenu()
    End Sub


    ' enable relevant menus for module only when this module's window is open
    Private Sub resetModuleRelatedMenu()
        Dim node As TreeNode = Me.TreeViewProjExplorer.SelectedNode
        If (Not node Is Nothing) And (Not node Is Me.curProjNode) Then
            Dim moduleFullPath As String = CType(node.Tag, String)
            Dim moduleName As String = IOManager.getFile(moduleFullPath)
            If Me.getMdiChild(moduleFullPath) Is Nothing Then
                Me.setMenuText("")
                setEnableStatus_singleModMenu(False)
            Else
                Me.setMenuText(moduleName)
                setEnableStatus_singleModMenu(True)
            End If
        End If
    End Sub


    ' Double click on a node of Project explorer opens corresponding file.
    Private Overloads Sub TreeViewProjExplorer_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TreeViewProjExplorer.DoubleClick
        If Me.TreeViewProjExplorer.SelectedNode Is Me.curProjNode Then Exit Sub

        Dim modName As String = Me.TreeViewProjExplorer.SelectedNode.Text
        modName = IOManager.formatFileName(modName, SetupManager.MODULE_FILE_EXTENSION, False)
        Me.openModuleFile(CType(Me.TreeViewProjExplorer.SelectedNode.Tag, String), _
                          modName, True)
    End Sub


    ' Constants used by Property window.
    Private PROPERTY_TITLE_PROJ_FILE As String = "Project File"
    Private PROPERTY_TITLE_PROJ_FOLDER As String = "Project Folder"
    Private PROPERTY_TITLE_PROJ_TYPE As String = "Project Type"
    Private PROPERTY_TITLE_FILE_NAME As String = "File Name"
    Private PROPERTY_TITLE_FILE_PATH As String = "File Path"

    Private Sub refreshPropertyWindow()
        Me.ListViewProperties.Items.Clear()
        Dim node As TreeNode = Me.TreeViewProjExplorer.SelectedNode

        If Not node Is Nothing Then
            If node Is Me.curProjNode Then
                Me.addProperty(PROPERTY_TITLE_PROJ_FILE, node.Text & SetupManager.PROJ_FILE_EXTENSION, True)
                Me.addProperty(PROPERTY_TITLE_PROJ_FOLDER, CType(node.Tag, String), False)
                Me.addProperty(PROPERTY_TITLE_PROJ_TYPE, Me.getProjectTypeName(Me.curProjectType), False)
            Else    ' is a module node
                If node Is Me.curFrameworkModuleNode Then
                    ' Special handling for framework module
                    Me.addProperty(PROPERTY_TITLE_FILE_NAME, node.Text, False)
                Else
                    Me.addProperty(PROPERTY_TITLE_FILE_NAME, node.Text, True)
                End If
                Me.addProperty(PROPERTY_TITLE_FILE_PATH, CType(node.Tag, String), False)
            End If
        End If
    End Sub


    ' Return the type name of a project's type.
    Private Function getProjectTypeName(ByVal projType As ModuleMain.xcProjectType)
        If projType = ModuleMain.xcProjectType.xcGeneralIMSWebsiteProject Then
            Return "General Website"
        End If
    End Function


    Private Sub addProperty(ByVal Field As String, ByVal value As String, _
                            ByVal editable As Boolean)
        Dim str() As String = {Field, value}
        Dim lvi As New ListViewItem(str)
        If editable Then
            lvi.Tag = True
        Else
            lvi.Tag = False
            lvi.ForeColor = Color.Gray
        End If
        Me.ListViewProperties.Items.Add(lvi)
    End Sub


    ' Function to show context menu right/left upon clicking on Project Explorer
    Private Sub TreeViewProjExplorer_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TreeViewProjExplorer.MouseUp
        resumeTitleColor()
        Me.LabelProjExplorerTitle.ForeColor = Color.White
        Me.LabelProjExplorerTitle.BackColor = Color.Blue

        ' Show menu only if Right Mouse button is clicked
        If e.Button = MouseButtons.Right Then
            'MsgBox("right click")
            ' Point where mouse is clicked
            Dim p As Point = New Point(e.X, e.Y)

            ' Go to the node that the user clicked
            Dim node As TreeNode = Me.TreeViewProjExplorer.GetNodeAt(p)
            If Not node Is Nothing Then
                ' Highlight the node that the user clicked.
                ' The node is highlighted until the Menu is displayed on the screen
                TreeViewProjExplorer.SelectedNode = node

                ' Find the appropriate ContextMenu based on the highlighted node
                If node Is Me.curProjNode Then
                    Me.mnuPNFullPath.Text = CType(node.Tag, String)
                    Me.ContextMenuProjNode.Show(TreeViewProjExplorer, p)
                Else    ' is a module node
                    Me.mnuMNFullPath.Text = CType(node.Tag, String)
                    If Me.getMdiChild(CType(node.Tag, String)) Is Nothing Then 'window closed
                        Me.mnuMNClose.Enabled = False
                        Me.mnuMNOpen.Enabled = True
                    Else    ' window is open
                        Me.mnuMNClose.Enabled = True
                        Me.mnuMNOpen.Enabled = False
                    End If

                    ' Special handling for Framework module
                    If node Is Me.curFrameworkModuleNode Then
                        Me.mnuMNRename.Enabled = False
                        Me.mnuMNDelete.Enabled = False
                        Me.mnuMNExclude.Enabled = False
                    Else
                        Me.mnuMNRename.Enabled = True
                        Me.mnuMNDelete.Enabled = True
                        Me.mnuMNExclude.Enabled = True
                    End If

                    Me.ContextMenuModuleNode.Show(TreeViewProjExplorer, p)
                End If
            Else
                Me.ContextMenuProjExplorer.Show(TreeViewProjExplorer, p)
            End If
        End If

    End Sub


    Private Sub resetProjectExplorerNodesBackColor()
        If Not Me.curProjNode Is Nothing Then
            If Me.curProjNode.IsSelected Then
                Me.curProjNode.BackColor = Color.LightGray
            Else
                Me.curProjNode.BackColor = Color.White
            End If

            Dim node As TreeNode
            For Each node In Me.curProjNode.Nodes
                If node.IsSelected Then
                    node.BackColor = Color.LightGray
                Else
                    node.BackColor = Color.White
                End If
            Next
        End If
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions for Context menu ContextMenuModuleNode
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub mnuMNOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMNOpen.Click
        'Dim fileFullPath = CType(Me.TreeViewProjExplorer.SelectedNode.Tag, String)
        'Me.openModuleFile(fileFullPath)
        Dim modName As String = Me.TreeViewProjExplorer.SelectedNode.Text
        modName = IOManager.formatFileName(modName, SetupManager.MODULE_FILE_EXTENSION, False)
        Me.openModuleFile(Me.mnuMNFullPath.Text, modName, True)
    End Sub


    Private Sub mnuMNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMNClose.Click
        Me.closeMdiChild(Me.getMdiChild(Me.mnuMNFullPath.Text))
    End Sub

    Private Sub mnuMNExclude_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMNExclude.Click
        Me.exlcludeModuleFromProj(Me.mnuMNFullPath.Text, Me.TreeViewProjExplorer.SelectedNode)
    End Sub

    Private Sub exlcludeModuleFromProj(ByVal path As String, ByRef node As TreeNode)
        Me.closeMdiChild(Me.getMdiChild(path))
        Me.curProjNode.Nodes.Remove(node)
        Me.saveProject(Me.curProjectName, Me.curProjectPath, Me.curProjectType)
    End Sub

    Private Sub mnuMNDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMNDelete.Click
        'Dim filePath As String = Me.mnuMNFullPath.Text
        Me.deleteModuleFile()
    End Sub

    Private Sub deleteModuleFile()
        Dim node As TreeNode = Me.TreeViewProjExplorer.SelectedNode
        If node Is Nothing Or node Is Me.curProjNode Then Exit Sub

        Dim filePath As String = CType(node.Tag, String)
        If MessageBox.Show(node.Text & " will be deleted permanently.", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = DialogResult.OK Then
            Me.exlcludeModuleFromProj(filePath, node)
            IOManager.DeleteFile(filePath)
        End If
    End Sub

    Private Sub mnuMNRename_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMNRename.Click
        Me.TreeViewProjExplorer.SelectedNode.BeginEdit()
    End Sub

    Private Sub mnuMNProperties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMNProperties.Click
        Me.PanelProperties.Show()
        Me.setCurExplorer(Me.LabelPropertiesTitle)
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Functions for Context menu ContextMenuModuleNode
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub mnuPNCloseAllMod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPNCloseAllMod.Click
        Me.closeAllMdiChild()
    End Sub


    Private Sub mnuPNCloseProj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPNCloseProj.Click
        Me.closeCurrentProject()
    End Sub

    Private Sub mnuPNAddNewMod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPNAddNewMod.Click
        addNewModPrompt()
    End Sub

    Private Sub mnuPNAddOldMod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPNAddOldMod.Click
        Me.addExistingModPrompt()
    End Sub

    Private Sub mnuPNSaveProj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPNSaveProj.Click
        Me.saveProject(Me.curProjectName, Me.curProjectPath, Me.curProjectType)
    End Sub

    Private Sub mnuPNRename_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPNRename.Click
        Me.curProjNode.BeginEdit()
    End Sub

    Private Sub mnuPNProperties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPNProperties.Click
        Me.PanelProperties.Show()
        Me.setCurExplorer(Me.LabelPropertiesTitle)
    End Sub


    ' DOn't allow change Framework node name, since one website project has
    ' only one framework node and the framework module name is not used
    ' in code generation, so there is no point to change its name.
    Private Sub TreeViewProjExplorer_BeforeLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TreeViewProjExplorer.BeforeLabelEdit
        e.CancelEdit = (e.Node Is Me.curFrameworkModuleNode)
    End Sub

    Private Sub TreeViewProjExplorer_AfterLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TreeViewProjExplorer.AfterLabelEdit
        'MsgBox("sender: " & sender.GetType.ToString & vbCrLf & "e: " & e.GetType.ToString)
        If e.Node Is Me.curProjNode Then
            ' Get parent folder of the project folder
            'Dim parentFolder As String = IOManager.getFolder(MyUtil.removeEndString(CType(e.Node.Tag, String), "\"))

            If e.Label = "" Or e.Label = e.Node.Text Then
                e.CancelEdit = True ' Name is not changed, do nothing 
            Else
                If Not IOManager.isValidFileName(e.Label, True) Then
                    e.CancelEdit = True
                    e.Node.BeginEdit()
                Else    ' Only change project name and project file name, not path.
                    e.CancelEdit = False
                    Dim oldProjName As String = Me.curProjectName
                    Dim oldProjPath As String = Me.curProjectPath
                    Dim newProjName As String = e.Label

                    renameProject(e.Node, oldProjPath, oldProjName, newProjName)

                    'IO.File.Delete(oldProjPath & oldProjName & SetupManager.PROJ_FILE_EXTENSION)
                    'Me.curProjectName = newProjName
                    'Me.saveProject(Me.curProjectName, Me.curProjectPath)
                    'e.Node.Text = newProjName
                    'Me.refreshPropertyWindow()
                End If
            End If
        Else ' a module node
            Dim fileFolder As String = IOManager.getFolder(CType(e.Node.Tag, String))
            'MsgBox("e.label = " & e.Label & vbCrLf & "e.node.text = " & e.Node.Text & _
            '        vbCrLf & "new path = " & fileFolder & e.Label)

            If e.Label = "" Or e.Label = e.Node.Text Then
                e.CancelEdit = True ' Name is not changed, do nothing 
            Else
                If Not IOManager.isValidFileName(e.Label, True) Then
                    e.CancelEdit = True
                    e.Node.BeginEdit()
                ElseIf IOManager.fileExists(fileFolder & e.Label) Then
                    MsgBox("File " & fileFolder & e.Label & " already exists.")
                    e.CancelEdit = True
                    e.Node.BeginEdit()
                ElseIf Not e.Label.EndsWith(SetupManager.MODULE_FILE_EXTENSION) Then
                    MsgBox("Module file must end with extension " & SetupManager.MODULE_FILE_EXTENSION)
                    e.CancelEdit = True
                    e.Node.BeginEdit()
                Else
                    e.CancelEdit = False
                    Dim oldFileName = e.Node.Text
                    Dim oldFilePath = CType(e.Node.Tag, String)
                    Dim newFileName = e.Label
                    Dim newFilePath = fileFolder & e.Label

                    Me.renameModule(e.Node, oldFileName, oldFilePath, newFileName, newFilePath, True)
                End If
            End If
        End If

    End Sub


    Private Sub renameProject(ByRef node As TreeNode, ByVal oldProjPath As String, _
            ByVal oldProjName As String, ByVal newProjName As String)
        IO.File.Delete(oldProjPath & oldProjName & SetupManager.PROJ_FILE_EXTENSION)
        Me.curProjectName = newProjName
        Me.saveProject(Me.curProjectName, Me.curProjectPath, Me.curProjectType)
        node.Text = newProjName

        Me.refreshPropertyWindow()

        ' Update the list of recentProjects
        Me.recentProjects.RemoveAt(0)
        Me.recentProjects.Insert(0, Me.curProjectPath & Me.curProjectName & SetupManager.PROJ_FILE_EXTENSION)
        Me.showRecentProjects()
    End Sub

    ' rename a module and corresponding actions.
    ' The input newFileName is a valid one
    ' - rmOldFile - whether remove the file with old name
    Private Function renameModule(ByRef node As TreeNode, _
            ByVal oldFileName As String, ByVal oldFilePath As String, _
            ByVal newFileName As String, ByVal newFilePath As String, ByVal rmOldFile As Boolean) As Boolean

        If rmOldFile Then ' Used by Rename
            IO.File.Move(oldFilePath, newFilePath)
        Else ' used by SaveAs
            ' This is not need because Module.SaveModule will do this.
            'IO.File.Copy(oldFilePath, newFilePath)
        End If

        If Me.openingModules.Contains(UCase(oldFilePath)) Then
            Me.openingModules.Remove(UCase(oldFilePath))
            Me.openingModules.Add(UCase(newFilePath))
        End If
        Dim mdiChild As FormChildBase = Me.getMdiChild(oldFilePath)
        If Not mdiChild Is Nothing Then
            mdiChild.rename(IOManager.formatFileName(newFileName, _
            SetupManager.MODULE_FILE_EXTENSION, False), newFilePath)
            mdiChild.saveModule()

            'Me.mnuCloseMod.Text = "&Close Module " & newFileName
            'Me.mnuSaveModule.Text = "&Save Module " & newFileName
            'Me.mnuSaveModuleAs.Text = "Save Module " & newFileName & " &As ..."
            'Me.mnuBuildModule.Text = "B&uild Module " & newFileName
            Me.setMenuText(newFileName)
        End If
        node.Tag = newFilePath
        node.Text = newFileName

        Me.saveProject(Me.curProjectName, Me.curProjectPath, Me.curProjectType)
        Me.refreshPropertyWindow()
    End Function


    Private Sub TextBoxProperty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxProperty.TextChanged
        Dim caretPos As Integer = Me.TextBoxProperty.SelectionStart ' current caret postion
        If Not IOManager.isValidFileName(Me.TextBoxProperty.Text, True) Then
            Me.TextBoxProperty.Text = Me.prevTextBoxPropertyValue
            Me.TextBoxProperty.SelectionStart = caretPos - 1
        Else
            Me.prevTextBoxPropertyValue = Me.TextBoxProperty.Text
        End If
    End Sub

    Private Sub mnuSaveProj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveProj.Click
        Me.saveProject(Me.curProjectName, Me.curProjectPath, Me.curProjectType)
    End Sub

    Private Sub LabelPropertiesTitle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelPropertiesTitle.Click
        Me.setCurExplorer(Me.LabelPropertiesTitle)
    End Sub

    Private Sub mnuSaveModule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveModule.Click
        If Not Me.ActiveMdiChild Is Nothing Then
            CType(Me.ActiveMdiChild, FormChildBase).saveModule()
        End If
        Me.showWaitCursor()
    End Sub

    ' Functions for menu File -> Save Module As
    Private Sub mnuSaveModuleAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveModuleAs.Click
        Me.saveModuleAsPrompt()
    End Sub

    Private Sub saveModuleAsPrompt()
        Me.SaveFileDialogModule.FileName = CType(Me.ActiveMdiChild, FormChildBase).moduleName
        Me.SaveFileDialogModule.InitialDirectory = CType(Me.ActiveMdiChild, FormChildBase).moduleFullPath
        Me.SaveFileDialogModule.ShowDialog(Me)
    End Sub


    Private Sub SaveFileDialogModule_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialogModule.FileOk
        Dim fileFullPath As String = Me.SaveFileDialogModule.FileName
        'MsgBox("Folder: " & IOManager.getFolder(fileFullPath) & vbCrLf & _
        '       "File: " & IOManager.getFile(fileFullPath))
        Dim node As TreeNode = Me.TreeViewProjExplorer.SelectedNode

        Dim oldFileName = node.Text
        Dim oldFilePath = CType(node.Tag, String)
        Dim newFileName = IOManager.getFile(fileFullPath)
        Dim newFilePath = fileFullPath

        If Not Me.getNodeInProjectExplorer(newFilePath) Is Nothing Then
            MsgBox("Module " & newFileName & " is in current project and cannot be overwritten.")
            Exit Sub
        End If

        Me.renameModule(node, oldFileName, oldFilePath, newFileName, newFilePath, False)
    End Sub

    Private Sub LabelPropertyDesc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelPropertyDesc.Click
        Me.setCurExplorer(Me.LabelPropertiesTitle)
    End Sub

    Private Sub GroupBoxPropertyDesc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxPropertyDesc.Click
        Me.setCurExplorer(Me.LabelPropertiesTitle)
    End Sub

    Private Sub mnuPropFloat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPropFloat.Click

    End Sub

    Private Sub mnuPropHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPropHide.Click
        Me.closePropertiesWindow()
    End Sub

    Private Sub mnuPEHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPEHide.Click
        Me.closeProjExplorer()
    End Sub

    Private Sub mnuPEFloat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPEFloat.Click

    End Sub


    ' Allow open project or module by dragdrop onto the form.
    Private Sub FormMain_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragDrop
        e.Effect = DragDropEffects.All
        Dim path As String
        Dim s() As String = e.Data.GetData("FileDrop", False)
        Dim i As Integer
        For i = 0 To s.Length - 1
            path = s(i)
            'MsgBox(path)
            If path.EndsWith(SetupManager.PROJ_FILE_EXTENSION) Then
                Me.openProjectFile(path)
            ElseIf path.EndsWith(SetupManager.MODULE_FILE_EXTENSION) And _
                (Not Me.curProjNode Is Nothing) Then
                Me.openModuleFile(path, "", True)
            End If
        Next
    End Sub


    Private Sub FormMain_DragOver(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragOver
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub mnuBuildProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuBuildProject.Click
        Me.buildProject()
    End Sub


    Private Sub buildProject()
        ' Save file before building project
        Me.saveProject(Me.curProjectName, Me.curProjectPath, Me.curProjectType)

        If Me.curProjectType = ModuleMain.xcProjectType.xcGeneralIMSWebsiteProject Then
            Dim curFrameworkModule As FormChildBase
            curFrameworkModule = Me.getMdiChild(Me.getCurFrameworkModulePath())

            ' Always build the framwork module first
            If Not curFrameworkModule Is Nothing Then
                curFrameworkModule.buildModule()
            End If

            Dim mdiChild As FormChildBase
            For Each mdiChild In Me.MdiChildren
                If Not mdiChild Is curFrameworkModule Then mdiChild.buildModule()
            Next
        Else

        end if
    End Sub


    Private Sub mnuBuildModule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuBuildModule.Click
        If Not Me.ActiveMdiChild Is Nothing Then
            CType(Me.ActiveMdiChild, FormChildBase).buildModule()
        End If
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call this each time saving project.
    ' Before building project the project will be saved first,
    ' so it guarantees the framwork params are always up-to-date.
    '
    ' Framework params can be changed only when the framework module is open.
    ' So update framework params only when the framework module is open.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub updateFrameworkParams()
        If Me.curProjectType = ModuleMain.xcProjectType.xcGeneralIMSWebsiteProject Then

            Dim curFrameworkModule As FormGenFW
            curFrameworkModule = Me.getMdiChild(Me.getCurFrameworkModulePath())

            ' If framework module is open, read from the framework form
            If Not curFrameworkModule Is Nothing Then
                curFrameworkModule.getFrameworkParams(frameworkParams)
            End If
        End If
    End Sub
End Class
