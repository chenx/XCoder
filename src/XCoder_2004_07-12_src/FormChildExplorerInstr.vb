

Public Class FormChildExplorerInstr
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'PanelExplorer
        '
        Me.PanelExplorer.Size = New System.Drawing.Size(592, 447)
        Me.PanelExplorer.Visible = True
        '
        'FormChildExplorerInstr
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(592, 447)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelExplorer})
        Me.Name = "FormChildExplorerInstr"
        Me.Text = " Instruction Explorer"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Protected Overrides Sub hidePanel()
        myParent.hidePanelInstr()
    End Sub

    Protected Overrides Sub showPanel()        
        myParent.setOnFocusColor(myParent.LabelInstrTitle)
        myParent.PanelInstr.Show()
    End Sub

    Protected Overrides Sub initComponent()
        ' To be overriden
    End Sub

End Class
