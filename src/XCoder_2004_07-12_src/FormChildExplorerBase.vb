Public Class FormChildExplorerBase
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
    Friend WithEvents mnuDock As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHide As System.Windows.Forms.MenuItem
    Friend WithEvents mnuBorderStyle As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStandard As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolbox As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenuExplorer As System.Windows.Forms.ContextMenu
    Public WithEvents PanelExplorer As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.PanelExplorer = New System.Windows.Forms.Panel()
        Me.ContextMenuExplorer = New System.Windows.Forms.ContextMenu()
        Me.mnuDock = New System.Windows.Forms.MenuItem()
        Me.mnuHide = New System.Windows.Forms.MenuItem()
        Me.mnuBorderStyle = New System.Windows.Forms.MenuItem()
        Me.mnuStandard = New System.Windows.Forms.MenuItem()
        Me.mnuToolbox = New System.Windows.Forms.MenuItem()
        Me.SuspendLayout()
        '
        'PanelExplorer
        '
        Me.PanelExplorer.ContextMenu = Me.ContextMenuExplorer
        Me.PanelExplorer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelExplorer.Name = "PanelExplorer"
        Me.PanelExplorer.Size = New System.Drawing.Size(292, 260)
        Me.PanelExplorer.TabIndex = 0
        '
        'ContextMenuExplorer
        '
        Me.ContextMenuExplorer.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuDock, Me.mnuHide, Me.mnuBorderStyle})
        '
        'mnuDock
        '
        Me.mnuDock.Index = 0
        Me.mnuDock.Text = "Dock"
        '
        'mnuHide
        '
        Me.mnuHide.Index = 1
        Me.mnuHide.Text = "Hide"
        '
        'mnuBorderStyle
        '
        Me.mnuBorderStyle.Index = 2
        Me.mnuBorderStyle.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuStandard, Me.mnuToolbox})
        Me.mnuBorderStyle.Text = "Border Style"
        '
        'mnuStandard
        '
        Me.mnuStandard.Index = 0
        Me.mnuStandard.Text = "Stardard"
        '
        'mnuToolbox
        '
        Me.mnuToolbox.Enabled = False
        Me.mnuToolbox.Index = 1
        Me.mnuToolbox.Text = "Toobox"
        '
        'FormChildExplorerBase
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(292, 260)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelExplorer})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "FormChildExplorerBase"
        Me.ShowInTaskbar = False
        Me.Text = "FormChildExplorerBase"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Friend isFloating As Boolean
    Friend myParent As FormChildBase

    Friend Sub init(ByRef f As FormChildBase)
        Me.Owner = f
        myParent = f
        Me.initComponent()
    End Sub

    Private Sub dockMe()
        Me.isFloating = False
        Me.Hide()
        Me.showPanel()
    End Sub

    Friend Sub floatMe()
        Me.isFloating = True
        Me.Show()
        Me.BringToFront()
    End Sub

    ' Change functions of Alt+F4 and close button "X" on upper right corner
    ' from closing window to docking the window in FormChildBase.
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Dim SC_CLOSE As Integer = 61536
        Dim WM_SYSCOMMAND As Integer = 274
        ' determine whether it's system information, 
        ' and what to do with button "X" and alt + f4
        If m.Msg = WM_SYSCOMMAND AndAlso m.WParam.ToInt32 = SC_CLOSE Then
            Me.dockMe()
            Exit Sub
        End If
        MyBase.WndProc(m)
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Context menu functions
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub mnuDock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDock.Click
        Me.dockMe()
    End Sub

    Private Sub mnuHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHide.Click
        Me.Hide()
        Me.isFloating = False
        Me.hidePanel()
    End Sub

    Private Sub mnuStandard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStandard.Click
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.mnuStandard.Enabled = False
        Me.mnuToolbox.Enabled = True
    End Sub

    Private Sub mnuToolbox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolbox.Click
        Me.FormBorderStyle = FormBorderStyle.SizableToolWindow
        Me.mnuStandard.Enabled = True
        Me.mnuToolbox.Enabled = False
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Sub different for each explorer
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Protected Overridable Sub hidePanel()
        ' To be overriden
    End Sub

    Protected Overridable Sub showPanel()
        ' To be overriden
    End Sub

    Protected Overridable Sub initComponent()
        ' To be overriden
    End Sub

End Class
