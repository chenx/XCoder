Public Class FormTblAttr
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBoxWidth As System.Windows.Forms.TextBox
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxWidth = New System.Windows.Forms.TextBox()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Table Width:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBoxWidth
        '
        Me.TextBoxWidth.Location = New System.Drawing.Point(128, 24)
        Me.TextBoxWidth.Name = "TextBoxWidth"
        Me.TextBoxWidth.Size = New System.Drawing.Size(144, 22)
        Me.TextBoxWidth.TabIndex = 1
        Me.TextBoxWidth.Text = ""
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(144, 200)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(96, 23)
        Me.btnOk.TabIndex = 2
        Me.btnOk.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(264, 200)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        '
        'FormTblAttr
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(512, 260)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCancel, Me.btnOk, Me.TextBoxWidth, Me.Label1})
        Me.Name = "FormTblAttr"
        Me.Text = "Default Layout Table Attributes"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private _metadataForm As FormMetadataWizard
    Private _tbWidth As TextBox

    Public Sub init(ByRef metadataForm As FormMetadataWizard, ByRef textboxWidth As TextBox)
        Me._metadataForm = metadataForm
        Me._tbWidth = textboxWidth
    End Sub

    Private Sub TextBoxWidth_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxWidth.KeyPress
        If e.KeyChar() = Chr(13) Then
            e.Handled = True
            Me.previewMetadata()
        End If
    End Sub

    Private Sub previewMetadata()
        Me._tbWidth.Text = Me.TextBoxWidth.Text
        Me._metadataForm.previewMetadata()
    End Sub
End Class
