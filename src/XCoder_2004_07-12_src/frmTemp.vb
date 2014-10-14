Public Class frmTemp
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
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPageInput As System.Windows.Forms.TabPage
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TabPageSelect As System.Windows.Forms.TabPage
    Friend WithEvents TabPageTextarea As System.Windows.Forms.TabPage
    Friend WithEvents TabPageRadio As System.Windows.Forms.TabPage
    Friend WithEvents TabPageCheckbox As System.Windows.Forms.TabPage
    Friend WithEvents TabPageComment As System.Windows.Forms.TabPage
    Friend WithEvents btnFinish As System.Windows.Forms.Button
    Friend WithEvents btnNextField As System.Windows.Forms.Button
    Friend WithEvents ComboBoxFieldType As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TabControl2 = New System.Windows.Forms.TabControl()
        Me.TabPageInput = New System.Windows.Forms.TabPage()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TabPageSelect = New System.Windows.Forms.TabPage()
        Me.TabPageTextarea = New System.Windows.Forms.TabPage()
        Me.TabPageRadio = New System.Windows.Forms.TabPage()
        Me.TabPageCheckbox = New System.Windows.Forms.TabPage()
        Me.TabPageComment = New System.Windows.Forms.TabPage()
        Me.btnFinish = New System.Windows.Forms.Button()
        Me.btnNextField = New System.Windows.Forms.Button()
        Me.ComboBoxFieldType = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.TabControl2.SuspendLayout()
        Me.TabPageInput.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl2
        '
        Me.TabControl2.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPageInput, Me.TabPageSelect, Me.TabPageTextarea, Me.TabPageRadio, Me.TabPageCheckbox, Me.TabPageComment})
        Me.TabControl2.Location = New System.Drawing.Point(-8, 200)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(760, 128)
        Me.TabControl2.TabIndex = 31
        '
        'TabPageInput
        '
        Me.TabPageInput.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label4, Me.Label3, Me.Label2, Me.TextBox6, Me.TextBox5, Me.TextBox4, Me.TextBox2, Me.TextBox3, Me.TextBox1, Me.Label7, Me.Label6, Me.Label5})
        Me.TabPageInput.Location = New System.Drawing.Point(4, 25)
        Me.TabPageInput.Name = "TabPageInput"
        Me.TabPageInput.Size = New System.Drawing.Size(752, 99)
        Me.TabPageInput.TabIndex = 0
        Me.TabPageInput.Text = "Text/Password/Hidden"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(216, 32)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "HTML Name"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(128, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Type"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Label"
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(504, 56)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(90, 22)
        Me.TextBox6.TabIndex = 19
        Me.TextBox6.Text = ""
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(408, 56)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(90, 22)
        Me.TextBox5.TabIndex = 18
        Me.TextBox5.Text = ""
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(312, 56)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(90, 22)
        Me.TextBox4.TabIndex = 17
        Me.TextBox4.Text = ""
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(120, 56)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(90, 22)
        Me.TextBox2.TabIndex = 15
        Me.TextBox2.Text = ""
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(216, 56)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(90, 22)
        Me.TextBox3.TabIndex = 16
        Me.TextBox3.Text = ""
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(24, 56)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(90, 22)
        Me.TextBox1.TabIndex = 14
        Me.TextBox1.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(504, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 16)
        Me.Label7.TabIndex = 26
        Me.Label7.Text = "DB Name"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(408, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 16)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Maxlength"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(312, 32)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 16)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "Size"
        '
        'TabPageSelect
        '
        Me.TabPageSelect.Location = New System.Drawing.Point(4, 25)
        Me.TabPageSelect.Name = "TabPageSelect"
        Me.TabPageSelect.Size = New System.Drawing.Size(752, 99)
        Me.TabPageSelect.TabIndex = 1
        Me.TabPageSelect.Text = "Select"
        Me.TabPageSelect.Visible = False
        '
        'TabPageTextarea
        '
        Me.TabPageTextarea.Location = New System.Drawing.Point(4, 25)
        Me.TabPageTextarea.Name = "TabPageTextarea"
        Me.TabPageTextarea.Size = New System.Drawing.Size(752, 99)
        Me.TabPageTextarea.TabIndex = 2
        Me.TabPageTextarea.Text = "Textarea"
        Me.TabPageTextarea.Visible = False
        '
        'TabPageRadio
        '
        Me.TabPageRadio.Location = New System.Drawing.Point(4, 25)
        Me.TabPageRadio.Name = "TabPageRadio"
        Me.TabPageRadio.Size = New System.Drawing.Size(752, 99)
        Me.TabPageRadio.TabIndex = 3
        Me.TabPageRadio.Text = "Radio"
        Me.TabPageRadio.Visible = False
        '
        'TabPageCheckbox
        '
        Me.TabPageCheckbox.Location = New System.Drawing.Point(4, 25)
        Me.TabPageCheckbox.Name = "TabPageCheckbox"
        Me.TabPageCheckbox.Size = New System.Drawing.Size(752, 99)
        Me.TabPageCheckbox.TabIndex = 4
        Me.TabPageCheckbox.Text = "Checkbox"
        Me.TabPageCheckbox.Visible = False
        '
        'TabPageComment
        '
        Me.TabPageComment.Location = New System.Drawing.Point(4, 25)
        Me.TabPageComment.Name = "TabPageComment"
        Me.TabPageComment.Size = New System.Drawing.Size(752, 99)
        Me.TabPageComment.TabIndex = 5
        Me.TabPageComment.Text = "HTML Content"
        Me.TabPageComment.Visible = False
        '
        'btnFinish
        '
        Me.btnFinish.Location = New System.Drawing.Point(488, 88)
        Me.btnFinish.Name = "btnFinish"
        Me.btnFinish.Size = New System.Drawing.Size(168, 23)
        Me.btnFinish.TabIndex = 34
        Me.btnFinish.Text = "Finish"
        '
        'btnNextField
        '
        Me.btnNextField.Location = New System.Drawing.Point(312, 88)
        Me.btnNextField.Name = "btnNextField"
        Me.btnNextField.Size = New System.Drawing.Size(168, 23)
        Me.btnNextField.TabIndex = 33
        Me.btnNextField.Text = "Next Field"
        '
        'ComboBoxFieldType
        '
        Me.ComboBoxFieldType.Items.AddRange(New Object() {"Input(Text/Password/Hidden)", "Select", "Textarea", "Radio", "Checkbox", "HTML Content"})
        Me.ComboBoxFieldType.Location = New System.Drawing.Point(40, 88)
        Me.ComboBoxFieldType.Name = "ComboBoxFieldType"
        Me.ComboBoxFieldType.Size = New System.Drawing.Size(216, 24)
        Me.ComboBoxFieldType.TabIndex = 32
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(456, 432)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(168, 23)
        Me.Button1.TabIndex = 43
        Me.Button1.Text = "Finish"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(280, 432)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(168, 23)
        Me.Button2.TabIndex = 42
        Me.Button2.Text = "Next Field"
        '
        'ComboBox1
        '
        Me.ComboBox1.Items.AddRange(New Object() {"Text Box", "Password Box", "Hidden Field", "Select", "Textarea", "Radio", "Checkbox", "HTML Content"})
        Me.ComboBox1.Location = New System.Drawing.Point(8, 432)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(216, 24)
        Me.ComboBox1.TabIndex = 41
        '
        'RadioButton2
        '
        Me.RadioButton2.Location = New System.Drawing.Point(624, 400)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(128, 24)
        Me.RadioButton2.TabIndex = 40
        Me.RadioButton2.Text = "Custom Layout"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 400)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(400, 24)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "This wizard walks through the process of creating a metadata file."
        '
        'RadioButton1
        '
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(488, 400)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(128, 24)
        Me.RadioButton1.TabIndex = 39
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Default Layout"
        '
        'frmTemp
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(744, 528)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.Button2, Me.ComboBox1, Me.RadioButton2, Me.Label1, Me.RadioButton1, Me.btnFinish, Me.btnNextField, Me.ComboBoxFieldType, Me.TabControl2})
        Me.Name = "frmTemp"
        Me.Text = "frmTemp"
        Me.TabControl2.ResumeLayout(False)
        Me.TabPageInput.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmTemp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
