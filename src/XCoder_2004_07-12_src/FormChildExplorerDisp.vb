

Public Class FormChildExplorerDisp
    Inherits XinCoder.FormChildExplorerBase

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
    Friend WithEvents TabControlDisplay As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents AxWebBrowser1 As AxSHDocVw.AxWebBrowser
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormChildExplorerDisp))
        Me.TabControlDisplay = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.AxWebBrowser1 = New AxSHDocVw.AxWebBrowser()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.PanelExplorer.SuspendLayout()
        Me.TabControlDisplay.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelExplorer
        '
        Me.PanelExplorer.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControlDisplay})
        Me.PanelExplorer.Size = New System.Drawing.Size(592, 447)
        Me.PanelExplorer.Visible = True
        '
        'TabControlDisplay
        '
        Me.TabControlDisplay.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.TabControlDisplay.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPage1, Me.TabPage2})
        Me.TabControlDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControlDisplay.ItemSize = New System.Drawing.Size(78, 21)
        Me.TabControlDisplay.Multiline = True
        Me.TabControlDisplay.Name = "TabControlDisplay"
        Me.TabControlDisplay.SelectedIndex = 0
        Me.TabControlDisplay.ShowToolTips = True
        Me.TabControlDisplay.Size = New System.Drawing.Size(592, 447)
        Me.TabControlDisplay.TabIndex = 56
        '
        'TabPage1
        '
        Me.TabPage1.Controls.AddRange(New System.Windows.Forms.Control() {Me.AxWebBrowser1})
        Me.TabPage1.Location = New System.Drawing.Point(4, 4)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(584, 418)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Quick View"
        '
        'AxWebBrowser1
        '
        Me.AxWebBrowser1.ContainingControl = Me
        Me.AxWebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AxWebBrowser1.Enabled = True
        Me.AxWebBrowser1.OcxState = CType(resources.GetObject("AxWebBrowser1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxWebBrowser1.Size = New System.Drawing.Size(584, 418)
        Me.AxWebBrowser1.TabIndex = 47
        '
        'TabPage2
        '
        Me.TabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.RichTextBox1})
        Me.TabPage2.Location = New System.Drawing.Point(4, 4)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(584, 418)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Source"
        Me.TabPage2.Visible = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(584, 418)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.WordWrap = False
        '
        'FormChildExplorerDisp
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(592, 447)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelExplorer})
        Me.Name = "FormChildExplorerDisp"
        Me.Text = " Display Explorer"
        Me.PanelExplorer.ResumeLayout(False)
        Me.TabControlDisplay.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Protected Overrides Sub hidePanel()
        myParent.hidePanelDisp()
    End Sub

    Protected Overrides Sub showPanel()
        myParent.PanelInstr.Dock = DockStyle.Bottom
        myParent.setOnFocusColor(myParent.LabelDispTitle)
        myParent.PanelDisplay.Show()
    End Sub

    Protected Overrides Sub initComponent()
        ' To be overriden
    End Sub

End Class
