

Public Class FormChildExplorerStep
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
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.PanelExplorer.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelExplorer
        '
        Me.PanelExplorer.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnNext, Me.btnBack})
        Me.PanelExplorer.Size = New System.Drawing.Size(242, 567)
        Me.PanelExplorer.Visible = True
        '
        'btnNext
        '
        Me.btnNext.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNext.BackColor = System.Drawing.SystemColors.Control
        Me.btnNext.Location = New System.Drawing.Point(95, 708)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(96, 32)
        Me.btnNext.TabIndex = 6
        Me.btnNext.Text = "Next"
        '
        'btnBack
        '
        Me.btnBack.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnBack.BackColor = System.Drawing.SystemColors.Control
        Me.btnBack.Location = New System.Drawing.Point(7, 708)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(88, 32)
        Me.btnBack.TabIndex = 5
        Me.btnBack.Text = "Back"
        '
        'FormChildExplorerStep
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(242, 567)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelExplorer})
        Me.Name = "FormChildExplorerStep"
        Me.Text = " Step Explorer"
        Me.PanelExplorer.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Protected Overrides Sub hidePanel()
        myParent.hidePanelStep()
    End Sub

    Protected Overrides Sub showPanel()
        myParent.setOnFocusColor(myParent.LabelStepTitle)
        myParent.PanelStep.Show()
    End Sub

    Protected Overrides Sub initComponent()
        ' To be overriden
    End Sub

End Class
