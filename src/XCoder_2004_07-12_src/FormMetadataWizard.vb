Imports System.Text.RegularExpressions

Public Class FormMetadataWizard
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TabPageSelect As System.Windows.Forms.TabPage
    Friend WithEvents TabPageTextarea As System.Windows.Forms.TabPage
    Friend WithEvents TabPageRadio As System.Windows.Forms.TabPage
    Friend WithEvents TabPageCheckbox As System.Windows.Forms.TabPage
    Friend WithEvents TabPageComment As System.Windows.Forms.TabPage
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents TabControlField As System.Windows.Forms.TabControl
    Friend WithEvents TextBoxSelSize As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxSelWidth As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxSelType As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxSelName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxSelLabel As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxSelOptions As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxSelDBName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTADBName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTACols As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTARows As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTAType As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTAName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTALabel As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxTAWrap As System.Windows.Forms.CheckBox
    Friend WithEvents TextBoxRadioOptions As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRadioDBName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRadioType As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRadioName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRadioLabel As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCBText As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCBDBName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCBValue As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCBType As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCBName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCBLabel As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxComment As System.Windows.Forms.TextBox
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents TabControlView As System.Windows.Forms.TabControl
    Friend WithEvents TabPagePreview As System.Windows.Forms.TabPage

    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnFinish As System.Windows.Forms.Button
    Friend WithEvents AxWebBrowser1 As AxSHDocVw.AxWebBrowser
    Friend WithEvents ButtonPreview As System.Windows.Forms.Button
    Friend WithEvents TabPageText As System.Windows.Forms.TabPage
    Friend WithEvents TabPagePassword As System.Windows.Forms.TabPage
    Friend WithEvents TabPageHidden As System.Windows.Forms.TabPage
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents btnInsert As System.Windows.Forms.Button
    Friend WithEvents TextBoxTextDBName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTextMaxLength As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTextSize As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTextType As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTextName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTextLabel As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxPassDBName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxPassMaxLength As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxPassSize As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxPassType As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxPassName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxPassLabel As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxHiddenDBName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxHiddenMaxLength As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxHiddenSize As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxHiddenType As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxHiddenName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxHiddenLabel As System.Windows.Forms.TextBox
    Friend WithEvents TabPageEditMetadata As System.Windows.Forms.TabPage
    Friend WithEvents TabPageEdit As System.Windows.Forms.TabPage
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ListBoxEdit As System.Windows.Forms.ListBox
    Friend WithEvents btnUp As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnSaveFreeEdit As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBoxTableAttr As System.Windows.Forms.GroupBox
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents RadioButtonLayoutDefault As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonLayoutCustom As System.Windows.Forms.RadioButton
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFinishEdit As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents TabPageHR As System.Windows.Forms.TabPage
    Friend WithEvents TabPageBlankLine As System.Windows.Forms.TabPage
    Friend WithEvents TabPageNewTable As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents TextBoxHRType As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxHRLabel As System.Windows.Forms.TextBox
    Friend WithEvents Label50 As System.Windows.Forms.Label
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Friend WithEvents TextBoxLineType As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxLineNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents TextBoxNewTableType As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxNewTableCol1Width As System.Windows.Forms.TextBox
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTableCol1Width As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTableWidth As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTableBorder As System.Windows.Forms.TextBox
    Friend WithEvents Label55 As System.Windows.Forms.Label
    Friend WithEvents Label56 As System.Windows.Forms.Label
    Friend WithEvents TextBoxTablePadding As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTableSpacing As System.Windows.Forms.TextBox
    Friend WithEvents Label57 As System.Windows.Forms.Label
    Friend WithEvents TextBoxTableBgColor As System.Windows.Forms.TextBox
    Friend WithEvents Label58 As System.Windows.Forms.Label
    Friend WithEvents ButtonTblAttr As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormMetadataWizard))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.GroupBoxTableAttr = New System.Windows.Forms.GroupBox()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.TextBoxTableBgColor = New System.Windows.Forms.TextBox()
        Me.Label55 = New System.Windows.Forms.Label()
        Me.Label56 = New System.Windows.Forms.Label()
        Me.TextBoxTablePadding = New System.Windows.Forms.TextBox()
        Me.TextBoxTableSpacing = New System.Windows.Forms.TextBox()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.TextBoxTableCol1Width = New System.Windows.Forms.TextBox()
        Me.TextBoxTableWidth = New System.Windows.Forms.TextBox()
        Me.TextBoxTableBorder = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label58 = New System.Windows.Forms.Label()
        Me.RadioButtonLayoutDefault = New System.Windows.Forms.RadioButton()
        Me.RadioButtonLayoutCustom = New System.Windows.Forms.RadioButton()
        Me.ButtonPreview = New System.Windows.Forms.Button()
        Me.btnFinish = New System.Windows.Forms.Button()
        Me.btnInsert = New System.Windows.Forms.Button()
        Me.TabControlField = New System.Windows.Forms.TabControl()
        Me.TabPageText = New System.Windows.Forms.TabPage()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBoxTextDBName = New System.Windows.Forms.TextBox()
        Me.TextBoxTextMaxLength = New System.Windows.Forms.TextBox()
        Me.TextBoxTextSize = New System.Windows.Forms.TextBox()
        Me.TextBoxTextType = New System.Windows.Forms.TextBox()
        Me.TextBoxTextName = New System.Windows.Forms.TextBox()
        Me.TextBoxTextLabel = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TabPagePassword = New System.Windows.Forms.TabPage()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.TextBoxPassDBName = New System.Windows.Forms.TextBox()
        Me.TextBoxPassMaxLength = New System.Windows.Forms.TextBox()
        Me.TextBoxPassSize = New System.Windows.Forms.TextBox()
        Me.TextBoxPassType = New System.Windows.Forms.TextBox()
        Me.TextBoxPassName = New System.Windows.Forms.TextBox()
        Me.TextBoxPassLabel = New System.Windows.Forms.TextBox()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.TabPageHidden = New System.Windows.Forms.TabPage()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.TextBoxHiddenDBName = New System.Windows.Forms.TextBox()
        Me.TextBoxHiddenMaxLength = New System.Windows.Forms.TextBox()
        Me.TextBoxHiddenSize = New System.Windows.Forms.TextBox()
        Me.TextBoxHiddenType = New System.Windows.Forms.TextBox()
        Me.TextBoxHiddenName = New System.Windows.Forms.TextBox()
        Me.TextBoxHiddenLabel = New System.Windows.Forms.TextBox()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.TabPageTextarea = New System.Windows.Forms.TabPage()
        Me.CheckBoxTAWrap = New System.Windows.Forms.CheckBox()
        Me.TextBoxTADBName = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.TextBoxTACols = New System.Windows.Forms.TextBox()
        Me.TextBoxTARows = New System.Windows.Forms.TextBox()
        Me.TextBoxTAType = New System.Windows.Forms.TextBox()
        Me.TextBoxTAName = New System.Windows.Forms.TextBox()
        Me.TextBoxTALabel = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.TabPageSelect = New System.Windows.Forms.TabPage()
        Me.TextBoxSelOptions = New System.Windows.Forms.TextBox()
        Me.TextBoxSelDBName = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TextBoxSelSize = New System.Windows.Forms.TextBox()
        Me.TextBoxSelWidth = New System.Windows.Forms.TextBox()
        Me.TextBoxSelType = New System.Windows.Forms.TextBox()
        Me.TextBoxSelName = New System.Windows.Forms.TextBox()
        Me.TextBoxSelLabel = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.TabPageRadio = New System.Windows.Forms.TabPage()
        Me.TextBoxRadioOptions = New System.Windows.Forms.TextBox()
        Me.TextBoxRadioDBName = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.TextBoxRadioType = New System.Windows.Forms.TextBox()
        Me.TextBoxRadioName = New System.Windows.Forms.TextBox()
        Me.TextBoxRadioLabel = New System.Windows.Forms.TextBox()
        Me.TabPageCheckbox = New System.Windows.Forms.TabPage()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.TextBoxCBText = New System.Windows.Forms.TextBox()
        Me.TextBoxCBDBName = New System.Windows.Forms.TextBox()
        Me.TextBoxCBValue = New System.Windows.Forms.TextBox()
        Me.TextBoxCBType = New System.Windows.Forms.TextBox()
        Me.TextBoxCBName = New System.Windows.Forms.TextBox()
        Me.TextBoxCBLabel = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.TabPageComment = New System.Windows.Forms.TabPage()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBoxComment = New System.Windows.Forms.TextBox()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.TabPageHR = New System.Windows.Forms.TabPage()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label49 = New System.Windows.Forms.Label()
        Me.TextBoxHRType = New System.Windows.Forms.TextBox()
        Me.TextBoxHRLabel = New System.Windows.Forms.TextBox()
        Me.TabPageBlankLine = New System.Windows.Forms.TabPage()
        Me.Label50 = New System.Windows.Forms.Label()
        Me.Label51 = New System.Windows.Forms.Label()
        Me.TextBoxLineType = New System.Windows.Forms.TextBox()
        Me.TextBoxLineNumber = New System.Windows.Forms.TextBox()
        Me.TabPageNewTable = New System.Windows.Forms.TabPage()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.TextBoxNewTableType = New System.Windows.Forms.TextBox()
        Me.TextBoxNewTableCol1Width = New System.Windows.Forms.TextBox()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.TabControlView = New System.Windows.Forms.TabControl()
        Me.TabPagePreview = New System.Windows.Forms.TabPage()
        Me.AxWebBrowser1 = New AxSHDocVw.AxWebBrowser()
        Me.TabPageEdit = New System.Windows.Forms.TabPage()
        Me.ListBoxEdit = New System.Windows.Forms.ListBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnFinishEdit = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnDown = New System.Windows.Forms.Button()
        Me.btnUp = New System.Windows.Forms.Button()
        Me.TabPageEditMetadata = New System.Windows.Forms.TabPage()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSaveFreeEdit = New System.Windows.Forms.Button()
        Me.ButtonTblAttr = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBoxTableAttr.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabControlField.SuspendLayout()
        Me.TabPageText.SuspendLayout()
        Me.TabPagePassword.SuspendLayout()
        Me.TabPageHidden.SuspendLayout()
        Me.TabPageTextarea.SuspendLayout()
        Me.TabPageSelect.SuspendLayout()
        Me.TabPageRadio.SuspendLayout()
        Me.TabPageCheckbox.SuspendLayout()
        Me.TabPageComment.SuspendLayout()
        Me.TabPageHR.SuspendLayout()
        Me.TabPageBlankLine.SuspendLayout()
        Me.TabPageNewTable.SuspendLayout()
        Me.TabControlView.SuspendLayout()
        Me.TabPagePreview.SuspendLayout()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPageEdit.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.TabPageEditMetadata.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1, Me.TabControlField})
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(792, 232)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Add Field"
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.ButtonTblAttr, Me.btnEdit, Me.GroupBoxTableAttr, Me.GroupBox2, Me.ButtonPreview, Me.btnFinish, Me.btnInsert})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 18)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(786, 123)
        Me.Panel1.TabIndex = 33
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(120, 88)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(100, 23)
        Me.btnEdit.TabIndex = 47
        Me.btnEdit.Text = "Edit Field"
        '
        'GroupBoxTableAttr
        '
        Me.GroupBoxTableAttr.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label57, Me.TextBoxTableBgColor, Me.Label55, Me.Label56, Me.TextBoxTablePadding, Me.TextBoxTableSpacing, Me.Label48, Me.Label47, Me.Label46, Me.TextBoxTableCol1Width, Me.TextBoxTableWidth, Me.TextBoxTableBorder})
        Me.GroupBoxTableAttr.Location = New System.Drawing.Point(440, 8)
        Me.GroupBoxTableAttr.Name = "GroupBoxTableAttr"
        Me.GroupBoxTableAttr.Size = New System.Drawing.Size(336, 104)
        Me.GroupBoxTableAttr.TabIndex = 46
        Me.GroupBoxTableAttr.TabStop = False
        Me.GroupBoxTableAttr.Text = "Default Layout Table Attribute"
        '
        'Label57
        '
        Me.Label57.Location = New System.Drawing.Point(8, 72)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(56, 16)
        Me.Label57.TabIndex = 13
        Me.Label57.Text = "BGColor"
        '
        'TextBoxTableBgColor
        '
        Me.TextBoxTableBgColor.Location = New System.Drawing.Point(80, 72)
        Me.TextBoxTableBgColor.Name = "TextBoxTableBgColor"
        Me.TextBoxTableBgColor.Size = New System.Drawing.Size(48, 22)
        Me.TextBoxTableBgColor.TabIndex = 12
        Me.TextBoxTableBgColor.Text = ""
        '
        'Label55
        '
        Me.Label55.Location = New System.Drawing.Point(176, 24)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(96, 16)
        Me.Label55.TabIndex = 11
        Me.Label55.Text = "Cellpadding"
        '
        'Label56
        '
        Me.Label56.Location = New System.Drawing.Point(176, 48)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(96, 16)
        Me.Label56.TabIndex = 10
        Me.Label56.Text = "Cellspacing"
        '
        'TextBoxTablePadding
        '
        Me.TextBoxTablePadding.Location = New System.Drawing.Point(272, 24)
        Me.TextBoxTablePadding.Name = "TextBoxTablePadding"
        Me.TextBoxTablePadding.Size = New System.Drawing.Size(48, 22)
        Me.TextBoxTablePadding.TabIndex = 9
        Me.TextBoxTablePadding.Text = ""
        '
        'TextBoxTableSpacing
        '
        Me.TextBoxTableSpacing.Location = New System.Drawing.Point(272, 48)
        Me.TextBoxTableSpacing.Name = "TextBoxTableSpacing"
        Me.TextBoxTableSpacing.Size = New System.Drawing.Size(48, 22)
        Me.TextBoxTableSpacing.TabIndex = 8
        Me.TextBoxTableSpacing.Text = ""
        '
        'Label48
        '
        Me.Label48.Location = New System.Drawing.Point(144, 72)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(120, 16)
        Me.Label48.TabIndex = 7
        Me.Label48.Text = "First Column Width"
        '
        'Label47
        '
        Me.Label47.Location = New System.Drawing.Point(8, 24)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(48, 16)
        Me.Label47.TabIndex = 6
        Me.Label47.Text = "Width"
        '
        'Label46
        '
        Me.Label46.Location = New System.Drawing.Point(8, 48)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(48, 16)
        Me.Label46.TabIndex = 5
        Me.Label46.Text = "Border"
        '
        'TextBoxTableCol1Width
        '
        Me.TextBoxTableCol1Width.Location = New System.Drawing.Point(272, 72)
        Me.TextBoxTableCol1Width.Name = "TextBoxTableCol1Width"
        Me.TextBoxTableCol1Width.Size = New System.Drawing.Size(48, 22)
        Me.TextBoxTableCol1Width.TabIndex = 2
        Me.TextBoxTableCol1Width.Text = ""
        '
        'TextBoxTableWidth
        '
        Me.TextBoxTableWidth.Location = New System.Drawing.Point(80, 24)
        Me.TextBoxTableWidth.Name = "TextBoxTableWidth"
        Me.TextBoxTableWidth.Size = New System.Drawing.Size(48, 22)
        Me.TextBoxTableWidth.TabIndex = 1
        Me.TextBoxTableWidth.Text = ""
        '
        'TextBoxTableBorder
        '
        Me.TextBoxTableBorder.Location = New System.Drawing.Point(80, 48)
        Me.TextBoxTableBorder.Name = "TextBoxTableBorder"
        Me.TextBoxTableBorder.Size = New System.Drawing.Size(48, 22)
        Me.TextBoxTableBorder.TabIndex = 0
        Me.TextBoxTableBorder.Text = "0"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label58, Me.RadioButtonLayoutDefault, Me.RadioButtonLayoutCustom})
        Me.GroupBox2.Location = New System.Drawing.Point(256, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(176, 104)
        Me.GroupBox2.TabIndex = 45
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Layout"
        '
        'Label58
        '
        Me.Label58.Location = New System.Drawing.Point(16, 48)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(144, 23)
        Me.Label58.TabIndex = 41
        Me.Label58.Text = "(Table of Two columns)"
        '
        'RadioButtonLayoutDefault
        '
        Me.RadioButtonLayoutDefault.Checked = True
        Me.RadioButtonLayoutDefault.Location = New System.Drawing.Point(16, 24)
        Me.RadioButtonLayoutDefault.Name = "RadioButtonLayoutDefault"
        Me.RadioButtonLayoutDefault.Size = New System.Drawing.Size(112, 24)
        Me.RadioButtonLayoutDefault.TabIndex = 39
        Me.RadioButtonLayoutDefault.TabStop = True
        Me.RadioButtonLayoutDefault.Text = "Default Layout"
        '
        'RadioButtonLayoutCustom
        '
        Me.RadioButtonLayoutCustom.Location = New System.Drawing.Point(16, 72)
        Me.RadioButtonLayoutCustom.Name = "RadioButtonLayoutCustom"
        Me.RadioButtonLayoutCustom.Size = New System.Drawing.Size(120, 24)
        Me.RadioButtonLayoutCustom.TabIndex = 40
        Me.RadioButtonLayoutCustom.Text = "Custom Layout"
        '
        'ButtonPreview
        '
        Me.ButtonPreview.Location = New System.Drawing.Point(16, 16)
        Me.ButtonPreview.Name = "ButtonPreview"
        Me.ButtonPreview.Size = New System.Drawing.Size(100, 23)
        Me.ButtonPreview.TabIndex = 44
        Me.ButtonPreview.Text = "Preview"
        '
        'btnFinish
        '
        Me.btnFinish.Location = New System.Drawing.Point(120, 16)
        Me.btnFinish.Name = "btnFinish"
        Me.btnFinish.Size = New System.Drawing.Size(100, 23)
        Me.btnFinish.TabIndex = 43
        Me.btnFinish.Text = "Finish"
        '
        'btnInsert
        '
        Me.btnInsert.Location = New System.Drawing.Point(16, 88)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(100, 23)
        Me.btnInsert.TabIndex = 42
        Me.btnInsert.Text = "Insert Field"
        '
        'TabControlField
        '
        Me.TabControlField.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPageText, Me.TabPagePassword, Me.TabPageHidden, Me.TabPageTextarea, Me.TabPageSelect, Me.TabPageRadio, Me.TabPageCheckbox, Me.TabPageComment, Me.TabPageHR, Me.TabPageBlankLine, Me.TabPageNewTable})
        Me.TabControlField.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TabControlField.Location = New System.Drawing.Point(3, 141)
        Me.TabControlField.Name = "TabControlField"
        Me.TabControlField.SelectedIndex = 0
        Me.TabControlField.Size = New System.Drawing.Size(786, 88)
        Me.TabControlField.TabIndex = 32
        '
        'TabPageText
        '
        Me.TabPageText.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label4, Me.Label3, Me.Label2, Me.TextBoxTextDBName, Me.TextBoxTextMaxLength, Me.TextBoxTextSize, Me.TextBoxTextType, Me.TextBoxTextName, Me.TextBoxTextLabel, Me.Label7, Me.Label6, Me.Label5})
        Me.TabPageText.Location = New System.Drawing.Point(4, 25)
        Me.TabPageText.Name = "TabPageText"
        Me.TabPageText.Size = New System.Drawing.Size(778, 59)
        Me.TabPageText.TabIndex = 0
        Me.TabPageText.Text = "Text box"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(200, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "HTML Name"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(112, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Type"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Label"
        '
        'TextBoxTextDBName
        '
        Me.TextBoxTextDBName.Location = New System.Drawing.Point(488, 32)
        Me.TextBoxTextDBName.Name = "TextBoxTextDBName"
        Me.TextBoxTextDBName.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxTextDBName.TabIndex = 19
        Me.TextBoxTextDBName.Text = ""
        '
        'TextBoxTextMaxLength
        '
        Me.TextBoxTextMaxLength.Location = New System.Drawing.Point(392, 32)
        Me.TextBoxTextMaxLength.Name = "TextBoxTextMaxLength"
        Me.TextBoxTextMaxLength.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxTextMaxLength.TabIndex = 18
        Me.TextBoxTextMaxLength.Text = "50"
        '
        'TextBoxTextSize
        '
        Me.TextBoxTextSize.Location = New System.Drawing.Point(296, 32)
        Me.TextBoxTextSize.Name = "TextBoxTextSize"
        Me.TextBoxTextSize.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxTextSize.TabIndex = 17
        Me.TextBoxTextSize.Text = "20"
        '
        'TextBoxTextType
        '
        Me.TextBoxTextType.Enabled = False
        Me.TextBoxTextType.ForeColor = System.Drawing.Color.Black
        Me.TextBoxTextType.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxTextType.Name = "TextBoxTextType"
        Me.TextBoxTextType.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxTextType.TabIndex = 15
        Me.TextBoxTextType.Text = "Text"
        '
        'TextBoxTextName
        '
        Me.TextBoxTextName.Location = New System.Drawing.Point(200, 32)
        Me.TextBoxTextName.Name = "TextBoxTextName"
        Me.TextBoxTextName.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxTextName.TabIndex = 16
        Me.TextBoxTextName.Text = ""
        '
        'TextBoxTextLabel
        '
        Me.TextBoxTextLabel.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxTextLabel.Name = "TextBoxTextLabel"
        Me.TextBoxTextLabel.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxTextLabel.TabIndex = 14
        Me.TextBoxTextLabel.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(488, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 26
        Me.Label7.Text = "Database Name"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(392, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 16)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Maxlength"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(296, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 16)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "Size"
        '
        'TabPagePassword
        '
        Me.TabPagePassword.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label34, Me.Label35, Me.Label36, Me.TextBoxPassDBName, Me.TextBoxPassMaxLength, Me.TextBoxPassSize, Me.TextBoxPassType, Me.TextBoxPassName, Me.TextBoxPassLabel, Me.Label37, Me.Label38, Me.Label39})
        Me.TabPagePassword.Location = New System.Drawing.Point(4, 25)
        Me.TabPagePassword.Name = "TabPagePassword"
        Me.TabPagePassword.Size = New System.Drawing.Size(778, 59)
        Me.TabPagePassword.TabIndex = 6
        Me.TabPagePassword.Text = "Password Box"
        '
        'Label34
        '
        Me.Label34.Location = New System.Drawing.Point(200, 8)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(88, 16)
        Me.Label34.TabIndex = 35
        Me.Label34.Text = "HTML Name"
        '
        'Label35
        '
        Me.Label35.Location = New System.Drawing.Point(112, 8)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(80, 16)
        Me.Label35.TabIndex = 34
        Me.Label35.Text = "Type"
        '
        'Label36
        '
        Me.Label36.Location = New System.Drawing.Point(8, 8)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(80, 16)
        Me.Label36.TabIndex = 33
        Me.Label36.Text = "Label"
        '
        'TextBoxPassDBName
        '
        Me.TextBoxPassDBName.Location = New System.Drawing.Point(488, 32)
        Me.TextBoxPassDBName.Name = "TextBoxPassDBName"
        Me.TextBoxPassDBName.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxPassDBName.TabIndex = 32
        Me.TextBoxPassDBName.Text = ""
        '
        'TextBoxPassMaxLength
        '
        Me.TextBoxPassMaxLength.Location = New System.Drawing.Point(392, 32)
        Me.TextBoxPassMaxLength.Name = "TextBoxPassMaxLength"
        Me.TextBoxPassMaxLength.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxPassMaxLength.TabIndex = 31
        Me.TextBoxPassMaxLength.Text = "50"
        '
        'TextBoxPassSize
        '
        Me.TextBoxPassSize.Location = New System.Drawing.Point(296, 32)
        Me.TextBoxPassSize.Name = "TextBoxPassSize"
        Me.TextBoxPassSize.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxPassSize.TabIndex = 30
        Me.TextBoxPassSize.Text = "20"
        '
        'TextBoxPassType
        '
        Me.TextBoxPassType.Enabled = False
        Me.TextBoxPassType.ForeColor = System.Drawing.Color.Black
        Me.TextBoxPassType.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxPassType.Name = "TextBoxPassType"
        Me.TextBoxPassType.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxPassType.TabIndex = 28
        Me.TextBoxPassType.Text = "Password"
        '
        'TextBoxPassName
        '
        Me.TextBoxPassName.Location = New System.Drawing.Point(200, 32)
        Me.TextBoxPassName.Name = "TextBoxPassName"
        Me.TextBoxPassName.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxPassName.TabIndex = 29
        Me.TextBoxPassName.Text = ""
        '
        'TextBoxPassLabel
        '
        Me.TextBoxPassLabel.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxPassLabel.Name = "TextBoxPassLabel"
        Me.TextBoxPassLabel.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxPassLabel.TabIndex = 27
        Me.TextBoxPassLabel.Text = ""
        '
        'Label37
        '
        Me.Label37.Location = New System.Drawing.Point(488, 8)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(136, 16)
        Me.Label37.TabIndex = 38
        Me.Label37.Text = "Database Name"
        '
        'Label38
        '
        Me.Label38.Location = New System.Drawing.Point(392, 8)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(72, 16)
        Me.Label38.TabIndex = 37
        Me.Label38.Text = "Maxlength"
        '
        'Label39
        '
        Me.Label39.Location = New System.Drawing.Point(296, 8)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(64, 16)
        Me.Label39.TabIndex = 36
        Me.Label39.Text = "Size"
        '
        'TabPageHidden
        '
        Me.TabPageHidden.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label40, Me.Label41, Me.Label42, Me.TextBoxHiddenDBName, Me.TextBoxHiddenMaxLength, Me.TextBoxHiddenSize, Me.TextBoxHiddenType, Me.TextBoxHiddenName, Me.TextBoxHiddenLabel, Me.Label43, Me.Label44, Me.Label45})
        Me.TabPageHidden.Location = New System.Drawing.Point(4, 25)
        Me.TabPageHidden.Name = "TabPageHidden"
        Me.TabPageHidden.Size = New System.Drawing.Size(778, 59)
        Me.TabPageHidden.TabIndex = 7
        Me.TabPageHidden.Text = "Hidden field"
        '
        'Label40
        '
        Me.Label40.Location = New System.Drawing.Point(200, 8)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(88, 16)
        Me.Label40.TabIndex = 35
        Me.Label40.Text = "HTML Name"
        '
        'Label41
        '
        Me.Label41.Location = New System.Drawing.Point(112, 8)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(80, 16)
        Me.Label41.TabIndex = 34
        Me.Label41.Text = "Type"
        '
        'Label42
        '
        Me.Label42.Location = New System.Drawing.Point(8, 8)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(80, 16)
        Me.Label42.TabIndex = 33
        Me.Label42.Text = "Label"
        '
        'TextBoxHiddenDBName
        '
        Me.TextBoxHiddenDBName.Location = New System.Drawing.Point(488, 32)
        Me.TextBoxHiddenDBName.Name = "TextBoxHiddenDBName"
        Me.TextBoxHiddenDBName.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxHiddenDBName.TabIndex = 32
        Me.TextBoxHiddenDBName.Text = ""
        '
        'TextBoxHiddenMaxLength
        '
        Me.TextBoxHiddenMaxLength.Location = New System.Drawing.Point(392, 32)
        Me.TextBoxHiddenMaxLength.Name = "TextBoxHiddenMaxLength"
        Me.TextBoxHiddenMaxLength.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxHiddenMaxLength.TabIndex = 31
        Me.TextBoxHiddenMaxLength.Text = "50"
        '
        'TextBoxHiddenSize
        '
        Me.TextBoxHiddenSize.Location = New System.Drawing.Point(296, 32)
        Me.TextBoxHiddenSize.Name = "TextBoxHiddenSize"
        Me.TextBoxHiddenSize.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxHiddenSize.TabIndex = 30
        Me.TextBoxHiddenSize.Text = "20"
        '
        'TextBoxHiddenType
        '
        Me.TextBoxHiddenType.Enabled = False
        Me.TextBoxHiddenType.ForeColor = System.Drawing.Color.Black
        Me.TextBoxHiddenType.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxHiddenType.Name = "TextBoxHiddenType"
        Me.TextBoxHiddenType.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxHiddenType.TabIndex = 28
        Me.TextBoxHiddenType.Text = "Hidden"
        '
        'TextBoxHiddenName
        '
        Me.TextBoxHiddenName.Location = New System.Drawing.Point(200, 32)
        Me.TextBoxHiddenName.Name = "TextBoxHiddenName"
        Me.TextBoxHiddenName.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxHiddenName.TabIndex = 29
        Me.TextBoxHiddenName.Text = ""
        '
        'TextBoxHiddenLabel
        '
        Me.TextBoxHiddenLabel.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxHiddenLabel.Name = "TextBoxHiddenLabel"
        Me.TextBoxHiddenLabel.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxHiddenLabel.TabIndex = 27
        Me.TextBoxHiddenLabel.Text = ""
        '
        'Label43
        '
        Me.Label43.Location = New System.Drawing.Point(488, 8)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(136, 16)
        Me.Label43.TabIndex = 38
        Me.Label43.Text = "Database Name"
        '
        'Label44
        '
        Me.Label44.Location = New System.Drawing.Point(392, 8)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(72, 16)
        Me.Label44.TabIndex = 37
        Me.Label44.Text = "Maxlength"
        '
        'Label45
        '
        Me.Label45.Location = New System.Drawing.Point(296, 8)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(64, 16)
        Me.Label45.TabIndex = 36
        Me.Label45.Text = "Size"
        '
        'TabPageTextarea
        '
        Me.TabPageTextarea.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckBoxTAWrap, Me.TextBoxTADBName, Me.Label11, Me.Label16, Me.Label17, Me.Label18, Me.Label19, Me.TextBoxTACols, Me.TextBoxTARows, Me.TextBoxTAType, Me.TextBoxTAName, Me.TextBoxTALabel, Me.Label20, Me.Label21})
        Me.TabPageTextarea.Location = New System.Drawing.Point(4, 25)
        Me.TabPageTextarea.Name = "TabPageTextarea"
        Me.TabPageTextarea.Size = New System.Drawing.Size(778, 59)
        Me.TabPageTextarea.TabIndex = 2
        Me.TabPageTextarea.Text = "Textarea"
        Me.TabPageTextarea.Visible = False
        '
        'CheckBoxTAWrap
        '
        Me.CheckBoxTAWrap.Location = New System.Drawing.Point(512, 32)
        Me.CheckBoxTAWrap.Name = "CheckBoxTAWrap"
        Me.CheckBoxTAWrap.TabIndex = 57
        Me.CheckBoxTAWrap.Text = "Wrap"
        '
        'TextBoxTADBName
        '
        Me.TextBoxTADBName.Location = New System.Drawing.Point(392, 32)
        Me.TextBoxTADBName.Name = "TextBoxTADBName"
        Me.TextBoxTADBName.Size = New System.Drawing.Size(112, 22)
        Me.TextBoxTADBName.TabIndex = 53
        Me.TextBoxTADBName.Text = ""
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(512, 8)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(88, 16)
        Me.Label11.TabIndex = 56
        Me.Label11.Text = "Wrap"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(392, 8)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(104, 16)
        Me.Label16.TabIndex = 55
        Me.Label16.Text = "Database Name"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(200, 8)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(88, 16)
        Me.Label17.TabIndex = 50
        Me.Label17.Text = "HTML Name"
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(112, 8)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(80, 16)
        Me.Label18.TabIndex = 49
        Me.Label18.Text = "Type"
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(8, 8)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(80, 16)
        Me.Label19.TabIndex = 48
        Me.Label19.Text = "Label"
        '
        'TextBoxTACols
        '
        Me.TextBoxTACols.Location = New System.Drawing.Point(344, 32)
        Me.TextBoxTACols.Name = "TextBoxTACols"
        Me.TextBoxTACols.Size = New System.Drawing.Size(40, 22)
        Me.TextBoxTACols.TabIndex = 47
        Me.TextBoxTACols.Text = "60"
        '
        'TextBoxTARows
        '
        Me.TextBoxTARows.Location = New System.Drawing.Point(296, 32)
        Me.TextBoxTARows.Name = "TextBoxTARows"
        Me.TextBoxTARows.Size = New System.Drawing.Size(40, 22)
        Me.TextBoxTARows.TabIndex = 46
        Me.TextBoxTARows.Text = "10"
        '
        'TextBoxTAType
        '
        Me.TextBoxTAType.Enabled = False
        Me.TextBoxTAType.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxTAType.Name = "TextBoxTAType"
        Me.TextBoxTAType.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxTAType.TabIndex = 44
        Me.TextBoxTAType.Text = "TextArea"
        '
        'TextBoxTAName
        '
        Me.TextBoxTAName.Location = New System.Drawing.Point(200, 32)
        Me.TextBoxTAName.Name = "TextBoxTAName"
        Me.TextBoxTAName.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxTAName.TabIndex = 45
        Me.TextBoxTAName.Text = ""
        '
        'TextBoxTALabel
        '
        Me.TextBoxTALabel.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxTALabel.Name = "TextBoxTALabel"
        Me.TextBoxTALabel.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxTALabel.TabIndex = 43
        Me.TextBoxTALabel.Text = ""
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(344, 8)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(40, 16)
        Me.Label20.TabIndex = 52
        Me.Label20.Text = "Cols"
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(296, 8)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(40, 16)
        Me.Label21.TabIndex = 51
        Me.Label21.Text = "Rows"
        '
        'TabPageSelect
        '
        Me.TabPageSelect.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBoxSelOptions, Me.TextBoxSelDBName, Me.Label14, Me.Label15, Me.Label8, Me.Label9, Me.Label10, Me.TextBoxSelSize, Me.TextBoxSelWidth, Me.TextBoxSelType, Me.TextBoxSelName, Me.TextBoxSelLabel, Me.Label12, Me.Label13})
        Me.TabPageSelect.Location = New System.Drawing.Point(4, 25)
        Me.TabPageSelect.Name = "TabPageSelect"
        Me.TabPageSelect.Size = New System.Drawing.Size(778, 59)
        Me.TabPageSelect.TabIndex = 1
        Me.TabPageSelect.Text = "Select"
        Me.TabPageSelect.Visible = False
        '
        'TextBoxSelOptions
        '
        Me.TextBoxSelOptions.Location = New System.Drawing.Point(512, 32)
        Me.TextBoxSelOptions.Name = "TextBoxSelOptions"
        Me.TextBoxSelOptions.Size = New System.Drawing.Size(256, 22)
        Me.TextBoxSelOptions.TabIndex = 40
        Me.TextBoxSelOptions.Text = ""
        '
        'TextBoxSelDBName
        '
        Me.TextBoxSelDBName.Location = New System.Drawing.Point(392, 32)
        Me.TextBoxSelDBName.Name = "TextBoxSelDBName"
        Me.TextBoxSelDBName.Size = New System.Drawing.Size(112, 22)
        Me.TextBoxSelDBName.TabIndex = 39
        Me.TextBoxSelDBName.Text = ""
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(512, 8)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(88, 16)
        Me.Label14.TabIndex = 42
        Me.Label14.Text = "Options"
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(392, 8)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(104, 16)
        Me.Label15.TabIndex = 41
        Me.Label15.Text = "Database Name"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(200, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(88, 16)
        Me.Label8.TabIndex = 35
        Me.Label8.Text = "HTML Name"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(112, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(80, 16)
        Me.Label9.TabIndex = 34
        Me.Label9.Text = "Type"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(80, 16)
        Me.Label10.TabIndex = 33
        Me.Label10.Text = "Label"
        '
        'TextBoxSelSize
        '
        Me.TextBoxSelSize.Location = New System.Drawing.Point(344, 32)
        Me.TextBoxSelSize.Name = "TextBoxSelSize"
        Me.TextBoxSelSize.Size = New System.Drawing.Size(40, 22)
        Me.TextBoxSelSize.TabIndex = 31
        Me.TextBoxSelSize.Text = "1"
        '
        'TextBoxSelWidth
        '
        Me.TextBoxSelWidth.Location = New System.Drawing.Point(296, 32)
        Me.TextBoxSelWidth.Name = "TextBoxSelWidth"
        Me.TextBoxSelWidth.Size = New System.Drawing.Size(40, 22)
        Me.TextBoxSelWidth.TabIndex = 30
        Me.TextBoxSelWidth.Text = "200"
        '
        'TextBoxSelType
        '
        Me.TextBoxSelType.Enabled = False
        Me.TextBoxSelType.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxSelType.Name = "TextBoxSelType"
        Me.TextBoxSelType.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxSelType.TabIndex = 28
        Me.TextBoxSelType.Text = "Select"
        '
        'TextBoxSelName
        '
        Me.TextBoxSelName.Location = New System.Drawing.Point(200, 32)
        Me.TextBoxSelName.Name = "TextBoxSelName"
        Me.TextBoxSelName.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxSelName.TabIndex = 29
        Me.TextBoxSelName.Text = ""
        '
        'TextBoxSelLabel
        '
        Me.TextBoxSelLabel.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxSelLabel.Name = "TextBoxSelLabel"
        Me.TextBoxSelLabel.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxSelLabel.TabIndex = 27
        Me.TextBoxSelLabel.Text = ""
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(344, 8)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(40, 16)
        Me.Label12.TabIndex = 37
        Me.Label12.Text = "Size"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(296, 8)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(40, 16)
        Me.Label13.TabIndex = 36
        Me.Label13.Text = "Width"
        '
        'TabPageRadio
        '
        Me.TabPageRadio.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBoxRadioOptions, Me.TextBoxRadioDBName, Me.Label22, Me.Label23, Me.Label24, Me.Label25, Me.Label26, Me.TextBoxRadioType, Me.TextBoxRadioName, Me.TextBoxRadioLabel})
        Me.TabPageRadio.Location = New System.Drawing.Point(4, 25)
        Me.TabPageRadio.Name = "TabPageRadio"
        Me.TabPageRadio.Size = New System.Drawing.Size(778, 59)
        Me.TabPageRadio.TabIndex = 3
        Me.TabPageRadio.Text = "Radio"
        Me.TabPageRadio.Visible = False
        '
        'TextBoxRadioOptions
        '
        Me.TextBoxRadioOptions.Location = New System.Drawing.Point(416, 32)
        Me.TextBoxRadioOptions.Name = "TextBoxRadioOptions"
        Me.TextBoxRadioOptions.Size = New System.Drawing.Size(352, 22)
        Me.TextBoxRadioOptions.TabIndex = 54
        Me.TextBoxRadioOptions.Text = ""
        '
        'TextBoxRadioDBName
        '
        Me.TextBoxRadioDBName.Location = New System.Drawing.Point(296, 32)
        Me.TextBoxRadioDBName.Name = "TextBoxRadioDBName"
        Me.TextBoxRadioDBName.Size = New System.Drawing.Size(112, 22)
        Me.TextBoxRadioDBName.TabIndex = 53
        Me.TextBoxRadioDBName.Text = ""
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(416, 8)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(88, 16)
        Me.Label22.TabIndex = 56
        Me.Label22.Text = "Options"
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(296, 8)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(104, 16)
        Me.Label23.TabIndex = 55
        Me.Label23.Text = "Database Name"
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(200, 8)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(88, 16)
        Me.Label24.TabIndex = 50
        Me.Label24.Text = "HTML Name"
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(112, 8)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(80, 16)
        Me.Label25.TabIndex = 49
        Me.Label25.Text = "Type"
        '
        'Label26
        '
        Me.Label26.Location = New System.Drawing.Point(8, 8)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(80, 16)
        Me.Label26.TabIndex = 48
        Me.Label26.Text = "Label"
        '
        'TextBoxRadioType
        '
        Me.TextBoxRadioType.Enabled = False
        Me.TextBoxRadioType.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxRadioType.Name = "TextBoxRadioType"
        Me.TextBoxRadioType.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxRadioType.TabIndex = 44
        Me.TextBoxRadioType.Text = "Radio"
        '
        'TextBoxRadioName
        '
        Me.TextBoxRadioName.Location = New System.Drawing.Point(200, 32)
        Me.TextBoxRadioName.Name = "TextBoxRadioName"
        Me.TextBoxRadioName.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxRadioName.TabIndex = 45
        Me.TextBoxRadioName.Text = ""
        '
        'TextBoxRadioLabel
        '
        Me.TextBoxRadioLabel.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxRadioLabel.Name = "TextBoxRadioLabel"
        Me.TextBoxRadioLabel.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxRadioLabel.TabIndex = 43
        Me.TextBoxRadioLabel.Text = ""
        '
        'TabPageCheckbox
        '
        Me.TabPageCheckbox.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label27, Me.Label28, Me.Label29, Me.TextBoxCBText, Me.TextBoxCBDBName, Me.TextBoxCBValue, Me.TextBoxCBType, Me.TextBoxCBName, Me.TextBoxCBLabel, Me.Label30, Me.Label31, Me.Label32})
        Me.TabPageCheckbox.Location = New System.Drawing.Point(4, 25)
        Me.TabPageCheckbox.Name = "TabPageCheckbox"
        Me.TabPageCheckbox.Size = New System.Drawing.Size(778, 59)
        Me.TabPageCheckbox.TabIndex = 4
        Me.TabPageCheckbox.Text = "Checkbox"
        Me.TabPageCheckbox.Visible = False
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(200, 8)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(88, 16)
        Me.Label27.TabIndex = 35
        Me.Label27.Text = "HTML Name"
        '
        'Label28
        '
        Me.Label28.Location = New System.Drawing.Point(112, 8)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(80, 16)
        Me.Label28.TabIndex = 34
        Me.Label28.Text = "Type"
        '
        'Label29
        '
        Me.Label29.Location = New System.Drawing.Point(8, 8)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(80, 16)
        Me.Label29.TabIndex = 33
        Me.Label29.Text = "Label"
        '
        'TextBoxCBText
        '
        Me.TextBoxCBText.Location = New System.Drawing.Point(512, 32)
        Me.TextBoxCBText.Name = "TextBoxCBText"
        Me.TextBoxCBText.Size = New System.Drawing.Size(256, 22)
        Me.TextBoxCBText.TabIndex = 32
        Me.TextBoxCBText.Text = ""
        '
        'TextBoxCBDBName
        '
        Me.TextBoxCBDBName.Location = New System.Drawing.Point(392, 32)
        Me.TextBoxCBDBName.Name = "TextBoxCBDBName"
        Me.TextBoxCBDBName.Size = New System.Drawing.Size(112, 22)
        Me.TextBoxCBDBName.TabIndex = 31
        Me.TextBoxCBDBName.Text = ""
        '
        'TextBoxCBValue
        '
        Me.TextBoxCBValue.Enabled = False
        Me.TextBoxCBValue.Location = New System.Drawing.Point(296, 32)
        Me.TextBoxCBValue.Name = "TextBoxCBValue"
        Me.TextBoxCBValue.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxCBValue.TabIndex = 30
        Me.TextBoxCBValue.Text = "Y"
        '
        'TextBoxCBType
        '
        Me.TextBoxCBType.Enabled = False
        Me.TextBoxCBType.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxCBType.Name = "TextBoxCBType"
        Me.TextBoxCBType.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxCBType.TabIndex = 28
        Me.TextBoxCBType.Text = "Checkbox"
        '
        'TextBoxCBName
        '
        Me.TextBoxCBName.Location = New System.Drawing.Point(200, 32)
        Me.TextBoxCBName.Name = "TextBoxCBName"
        Me.TextBoxCBName.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxCBName.TabIndex = 29
        Me.TextBoxCBName.Text = ""
        '
        'TextBoxCBLabel
        '
        Me.TextBoxCBLabel.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxCBLabel.Name = "TextBoxCBLabel"
        Me.TextBoxCBLabel.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxCBLabel.TabIndex = 27
        Me.TextBoxCBLabel.Text = ""
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(512, 8)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(136, 16)
        Me.Label30.TabIndex = 38
        Me.Label30.Text = "Text"
        '
        'Label31
        '
        Me.Label31.Location = New System.Drawing.Point(392, 8)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(104, 16)
        Me.Label31.TabIndex = 37
        Me.Label31.Text = "Database Name"
        '
        'Label32
        '
        Me.Label32.Location = New System.Drawing.Point(296, 8)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(64, 16)
        Me.Label32.TabIndex = 36
        Me.Label32.Text = "Value"
        '
        'TabPageComment
        '
        Me.TabPageComment.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label54, Me.TextBox1, Me.TextBoxComment, Me.Label33})
        Me.TabPageComment.Location = New System.Drawing.Point(4, 25)
        Me.TabPageComment.Name = "TabPageComment"
        Me.TabPageComment.Size = New System.Drawing.Size(778, 59)
        Me.TabPageComment.TabIndex = 5
        Me.TabPageComment.Text = "HTML Content"
        Me.TabPageComment.Visible = False
        '
        'Label54
        '
        Me.Label54.Location = New System.Drawing.Point(8, 8)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(80, 16)
        Me.Label54.TabIndex = 42
        Me.Label54.Text = "Type"
        '
        'TextBox1
        '
        Me.TextBox1.Enabled = False
        Me.TextBox1.ForeColor = System.Drawing.Color.Black
        Me.TextBox1.Location = New System.Drawing.Point(8, 32)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(90, 22)
        Me.TextBox1.TabIndex = 41
        Me.TextBox1.Text = "Comment"
        '
        'TextBoxComment
        '
        Me.TextBoxComment.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxComment.Name = "TextBoxComment"
        Me.TextBoxComment.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBoxComment.Size = New System.Drawing.Size(664, 22)
        Me.TextBoxComment.TabIndex = 39
        Me.TextBoxComment.Text = ""
        '
        'Label33
        '
        Me.Label33.Location = New System.Drawing.Point(104, 8)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(40, 16)
        Me.Label33.TabIndex = 40
        Me.Label33.Text = "Text"
        '
        'TabPageHR
        '
        Me.TabPageHR.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label1, Me.Label49, Me.TextBoxHRType, Me.TextBoxHRLabel})
        Me.TabPageHR.Location = New System.Drawing.Point(4, 25)
        Me.TabPageHR.Name = "TabPageHR"
        Me.TabPageHR.Size = New System.Drawing.Size(778, 59)
        Me.TabPageHR.TabIndex = 8
        Me.TabPageHR.Text = "Horizontal Bar"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(104, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 16)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "Type"
        '
        'Label49
        '
        Me.Label49.Location = New System.Drawing.Point(8, 8)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(80, 16)
        Me.Label49.TabIndex = 25
        Me.Label49.Text = "Label"
        '
        'TextBoxHRType
        '
        Me.TextBoxHRType.Enabled = False
        Me.TextBoxHRType.ForeColor = System.Drawing.Color.Black
        Me.TextBoxHRType.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxHRType.Name = "TextBoxHRType"
        Me.TextBoxHRType.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxHRType.TabIndex = 24
        Me.TextBoxHRType.Text = "hr"
        '
        'TextBoxHRLabel
        '
        Me.TextBoxHRLabel.Enabled = False
        Me.TextBoxHRLabel.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxHRLabel.Name = "TextBoxHRLabel"
        Me.TextBoxHRLabel.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxHRLabel.TabIndex = 23
        Me.TextBoxHRLabel.Text = "hr"
        '
        'TabPageBlankLine
        '
        Me.TabPageBlankLine.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label50, Me.Label51, Me.TextBoxLineType, Me.TextBoxLineNumber})
        Me.TabPageBlankLine.Location = New System.Drawing.Point(4, 25)
        Me.TabPageBlankLine.Name = "TabPageBlankLine"
        Me.TabPageBlankLine.Size = New System.Drawing.Size(778, 59)
        Me.TabPageBlankLine.TabIndex = 9
        Me.TabPageBlankLine.Text = "Blank Line"
        '
        'Label50
        '
        Me.Label50.Location = New System.Drawing.Point(104, 8)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(80, 16)
        Me.Label50.TabIndex = 26
        Me.Label50.Text = "Type"
        '
        'Label51
        '
        Me.Label51.Location = New System.Drawing.Point(8, 8)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(88, 16)
        Me.Label51.TabIndex = 25
        Me.Label51.Text = "Line Number"
        '
        'TextBoxLineType
        '
        Me.TextBoxLineType.Enabled = False
        Me.TextBoxLineType.ForeColor = System.Drawing.Color.Black
        Me.TextBoxLineType.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxLineType.Name = "TextBoxLineType"
        Me.TextBoxLineType.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxLineType.TabIndex = 24
        Me.TextBoxLineType.Text = "BlankLine"
        '
        'TextBoxLineNumber
        '
        Me.TextBoxLineNumber.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxLineNumber.Name = "TextBoxLineNumber"
        Me.TextBoxLineNumber.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxLineNumber.TabIndex = 23
        Me.TextBoxLineNumber.Text = ""
        '
        'TabPageNewTable
        '
        Me.TabPageNewTable.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label52, Me.Label53, Me.TextBoxNewTableType, Me.TextBoxNewTableCol1Width})
        Me.TabPageNewTable.Location = New System.Drawing.Point(4, 25)
        Me.TabPageNewTable.Name = "TabPageNewTable"
        Me.TabPageNewTable.Size = New System.Drawing.Size(778, 59)
        Me.TabPageNewTable.TabIndex = 10
        Me.TabPageNewTable.Text = "New Table"
        '
        'Label52
        '
        Me.Label52.Location = New System.Drawing.Point(104, 8)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(80, 16)
        Me.Label52.TabIndex = 26
        Me.Label52.Text = "Type"
        '
        'Label53
        '
        Me.Label53.Location = New System.Drawing.Point(8, 8)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(104, 16)
        Me.Label53.TabIndex = 25
        Me.Label53.Text = "Column 1 Width"
        '
        'TextBoxNewTableType
        '
        Me.TextBoxNewTableType.Enabled = False
        Me.TextBoxNewTableType.ForeColor = System.Drawing.Color.Black
        Me.TextBoxNewTableType.Location = New System.Drawing.Point(104, 32)
        Me.TextBoxNewTableType.Name = "TextBoxNewTableType"
        Me.TextBoxNewTableType.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxNewTableType.TabIndex = 24
        Me.TextBoxNewTableType.Text = "NewTable"
        '
        'TextBoxNewTableCol1Width
        '
        Me.TextBoxNewTableCol1Width.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxNewTableCol1Width.Name = "TextBoxNewTableCol1Width"
        Me.TextBoxNewTableCol1Width.Size = New System.Drawing.Size(90, 22)
        Me.TextBoxNewTableCol1Width.TabIndex = 23
        Me.TextBoxNewTableCol1Width.Text = ""
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter1.Location = New System.Drawing.Point(0, 232)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(792, 3)
        Me.Splitter1.TabIndex = 10
        Me.Splitter1.TabStop = False
        '
        'TabControlView
        '
        Me.TabControlView.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPagePreview, Me.TabPageEdit, Me.TabPageEditMetadata})
        Me.TabControlView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControlView.Location = New System.Drawing.Point(0, 235)
        Me.TabControlView.Name = "TabControlView"
        Me.TabControlView.SelectedIndex = 0
        Me.TabControlView.Size = New System.Drawing.Size(792, 325)
        Me.TabControlView.TabIndex = 11
        '
        'TabPagePreview
        '
        Me.TabPagePreview.Controls.AddRange(New System.Windows.Forms.Control() {Me.AxWebBrowser1})
        Me.TabPagePreview.Location = New System.Drawing.Point(4, 25)
        Me.TabPagePreview.Name = "TabPagePreview"
        Me.TabPagePreview.Size = New System.Drawing.Size(784, 296)
        Me.TabPagePreview.TabIndex = 0
        Me.TabPagePreview.Text = "Preview"
        '
        'AxWebBrowser1
        '
        Me.AxWebBrowser1.ContainingControl = Me
        Me.AxWebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AxWebBrowser1.Enabled = True
        Me.AxWebBrowser1.OcxState = CType(resources.GetObject("AxWebBrowser1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxWebBrowser1.Size = New System.Drawing.Size(784, 296)
        Me.AxWebBrowser1.TabIndex = 0
        '
        'TabPageEdit
        '
        Me.TabPageEdit.Controls.AddRange(New System.Windows.Forms.Control() {Me.ListBoxEdit, Me.Panel2})
        Me.TabPageEdit.Location = New System.Drawing.Point(4, 25)
        Me.TabPageEdit.Name = "TabPageEdit"
        Me.TabPageEdit.Size = New System.Drawing.Size(784, 296)
        Me.TabPageEdit.TabIndex = 2
        Me.TabPageEdit.Text = "Edit Field"
        '
        'ListBoxEdit
        '
        Me.ListBoxEdit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBoxEdit.ItemHeight = 16
        Me.ListBoxEdit.Location = New System.Drawing.Point(0, 40)
        Me.ListBoxEdit.Name = "ListBoxEdit"
        Me.ListBoxEdit.ScrollAlwaysVisible = True
        Me.ListBoxEdit.Size = New System.Drawing.Size(784, 256)
        Me.ListBoxEdit.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnFinishEdit, Me.btnUpdate, Me.btnDelete, Me.btnDown, Me.btnUp})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(784, 40)
        Me.Panel2.TabIndex = 0
        '
        'btnFinishEdit
        '
        Me.btnFinishEdit.Enabled = False
        Me.btnFinishEdit.Location = New System.Drawing.Point(432, 8)
        Me.btnFinishEdit.Name = "btnFinishEdit"
        Me.btnFinishEdit.Size = New System.Drawing.Size(100, 24)
        Me.btnFinishEdit.TabIndex = 4
        Me.btnFinishEdit.Text = "Finish Edit"
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(328, 8)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(100, 24)
        Me.btnUpdate.TabIndex = 3
        Me.btnUpdate.Text = "Update"
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(224, 8)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(100, 24)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Delete"
        '
        'btnDown
        '
        Me.btnDown.Enabled = False
        Me.btnDown.Location = New System.Drawing.Point(120, 8)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(100, 24)
        Me.btnDown.TabIndex = 1
        Me.btnDown.Text = "Down"
        '
        'btnUp
        '
        Me.btnUp.Enabled = False
        Me.btnUp.Location = New System.Drawing.Point(16, 8)
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(100, 24)
        Me.btnUp.TabIndex = 0
        Me.btnUp.Text = "Up"
        '
        'TabPageEditMetadata
        '
        Me.TabPageEditMetadata.Controls.AddRange(New System.Windows.Forms.Control() {Me.RichTextBox1, Me.Panel3})
        Me.TabPageEditMetadata.Location = New System.Drawing.Point(4, 25)
        Me.TabPageEditMetadata.Name = "TabPageEditMetadata"
        Me.TabPageEditMetadata.Size = New System.Drawing.Size(784, 296)
        Me.TabPageEditMetadata.TabIndex = 1
        Me.TabPageEditMetadata.Text = "View Metadata"
        Me.TabPageEditMetadata.Visible = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Location = New System.Drawing.Point(0, 40)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(784, 256)
        Me.RichTextBox1.TabIndex = 2
        Me.RichTextBox1.Text = ""
        '
        'Panel3
        '
        Me.Panel3.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnSaveFreeEdit})
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(784, 40)
        Me.Panel3.TabIndex = 1
        Me.Panel3.Visible = False
        '
        'btnSaveFreeEdit
        '
        Me.btnSaveFreeEdit.Location = New System.Drawing.Point(16, 8)
        Me.btnSaveFreeEdit.Name = "btnSaveFreeEdit"
        Me.btnSaveFreeEdit.Size = New System.Drawing.Size(100, 23)
        Me.btnSaveFreeEdit.TabIndex = 46
        Me.btnSaveFreeEdit.Text = "Save"
        '
        'ButtonTblAttr
        '
        Me.ButtonTblAttr.Location = New System.Drawing.Point(120, 56)
        Me.ButtonTblAttr.Name = "ButtonTblAttr"
        Me.ButtonTblAttr.Size = New System.Drawing.Size(104, 23)
        Me.ButtonTblAttr.TabIndex = 48
        Me.ButtonTblAttr.Text = "More..."
        Me.ButtonTblAttr.Visible = False
        '
        'FormMetadataWizard
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(792, 560)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControlView, Me.Splitter1, Me.GroupBox1})
        Me.Name = "FormMetadataWizard"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Create Registration Form Wizard"
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.GroupBoxTableAttr.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.TabControlField.ResumeLayout(False)
        Me.TabPageText.ResumeLayout(False)
        Me.TabPagePassword.ResumeLayout(False)
        Me.TabPageHidden.ResumeLayout(False)
        Me.TabPageTextarea.ResumeLayout(False)
        Me.TabPageSelect.ResumeLayout(False)
        Me.TabPageRadio.ResumeLayout(False)
        Me.TabPageCheckbox.ResumeLayout(False)
        Me.TabPageComment.ResumeLayout(False)
        Me.TabPageHR.ResumeLayout(False)
        Me.TabPageBlankLine.ResumeLayout(False)
        Me.TabPageNewTable.ResumeLayout(False)
        Me.TabControlView.ResumeLayout(False)
        Me.TabPagePreview.ResumeLayout(False)
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPageEdit.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.TabPageEditMetadata.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private ownerForm As FormChildBase

    Dim tblAttr As New clsHTMLTableAttributes()
    Private clsMD As New clsMetadata()
    Private metadata As String
    Private htmlPreviewCode As String
    Private previewFolderPath As String
    Private previewFilePath As String

    ' Store lists to prevent duplicate names.
    Private htmlNames As New ArrayList()
    Private dbNames As New ArrayList()

    Private lbIndex As Integer

    Public Sub init(ByVal previewFolderPath As String, ByVal metadata As String, _
               ByVal tblAttributes As clsHTMLTableAttributes, _
               ByRef ownerForm As FormChildBase, _
               Optional ByVal isLoginForm As Boolean = False)

        Me.ownerForm = ownerForm

        Me.tblAttr = tblAttributes
        Me.setTableAttributes()

        Me.metadata = metadata
        clsMD.init(Me.metadata)
        Me.htmlNames = MyUtil.arrayToArrayList(clsMD.formArray)
        Me.dbNames = MyUtil.arrayToArrayList(clsMD.dbArray)
        'MsgBox(MyUtil.arrayListToString(Me.htmlNames, ",") & vbCrLf & MyUtil.arrayListToString(Me.dbNames, ","))

        Me.previewFolderPath = previewFolderPath
        Me.previewFilePath = Me.previewFolderPath & "metadataPreview.html"

        lbIndex = -1
        Me.previewMetadata()
    End Sub


    Private Sub FormMetadataWizard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub setTableAttributes()
        Me.TextBoxTableBgColor.Text = Me.tblAttr.bgColor
        Me.TextBoxTableBorder.Text = Me.tblAttr.border
        Me.TextBoxTableCol1Width.Text = Me.tblAttr.firstColWidth
        Me.TextBoxTablePadding.Text = Me.tblAttr.cellPadding
        Me.TextBoxTableSpacing.Text = Me.tblAttr.cellSpacing
        Me.TextBoxTableWidth.Text = Me.tblAttr.width
    End Sub

    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Me.insertFieldToMetadata()
    End Sub


    Private Sub ButtonPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPreview.Click
        Me.previewMetadata()
        Me.TabControlView.SelectedTab = Me.TabPagePreview
    End Sub


    Private Function insertFieldToMetadata() As Boolean
        If Me.TabControlField.SelectedTab Is Me.TabPageText Then
            If Me.insertText() = False Then Return False
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPagePassword Then
            If Me.insertPassword() = False Then Return False
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageHidden Then
            If Me.insertHidden() = False Then Return False
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageSelect Then
            If Me.insertSelect() = False Then Return False
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageTextarea Then
            If Me.insertTextArea() = False Then Return False
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageRadio Then
            If Me.insertRadio() = False Then Return False
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageCheckbox Then
            If Me.insertCheckbox() = False Then Return False
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageComment Then
            If Me.insertComment() = False Then Return False
        ElseIf Me.RadioButtonLayoutDefault.Checked = True Then
            If Me.TabControlField.SelectedTab Is Me.TabPageHR Then
                If Me.insertHR() = False Then Return False
            ElseIf Me.TabControlField.SelectedTab Is Me.TabPageNewTable Then
                If Me.insertNewTable() = False Then Return False
            ElseIf Me.TabControlField.SelectedTab Is Me.TabPageBlankLine Then
                If Me.insertBlankLine() = False Then Return False
            End If
        End If
        Me.previewMetadata()
    End Function



    Private Function insertText() As Boolean
        Dim label As String = Trim(Me.TextBoxTextLabel.Text)
        Dim type As String = "text"
        Dim htmlName As String = Trim(Me.TextBoxTextName.Text)
        Dim size As String = Trim(Me.TextBoxTextSize.Text)
        Dim maxLen As String = Trim(Me.TextBoxTextMaxLength.Text)
        Dim dbName As String = Trim(Me.TextBoxTextDBName.Text)

        If validateText(label, htmlName, size, maxLen, dbName) = False Then Return False

        Me.htmlNames.Add(htmlName)
        Me.dbNames.Add(dbName)

        Me.metadata &= label & "|" & type & "|" & htmlName & "|" & _
                       size & "|" & maxLen & "|" & dbName & "|" & vbCrLf
        Me.clearTextTab()
        Return True
    End Function

    Private Sub clearTextTab()
        Me.TextBoxTextLabel.Text = ""
        Me.TextBoxTextName.Text = ""
        Me.TextBoxTextDBName.Text = ""
    End Sub

    Private Function insertPassword() As Boolean
        Dim label As String = Trim(Me.TextBoxPassLabel.Text)
        Dim type As String = "Password"
        Dim htmlName As String = Trim(Me.TextBoxPassName.Text)
        Dim size As String = Trim(Me.TextBoxPassSize.Text)
        Dim maxLen As String = Trim(Me.TextBoxPassMaxLength.Text)
        Dim dbName As String = Trim(Me.TextBoxPassDBName.Text)

        If validatePassword(label, htmlName, size, maxLen, dbName) = False Then Return False

        Me.htmlNames.Add(htmlName)
        Me.dbNames.Add(dbName)

        Me.metadata &= label & "|" & type & "|" & htmlName & "|" & _
                       size & "|" & maxLen & "|" & dbName & "|" & vbCrLf
        Me.clearPasswordTab()
        Return True
    End Function

    Private Sub clearPasswordTab()
        Me.TextBoxPassLabel.Text = ""
        Me.TextBoxPassName.Text = ""
        Me.TextBoxPassDBName.Text = ""
    End Sub

    Private Function insertHidden() As Boolean
        Dim label As String = Trim(Me.TextBoxHiddenLabel.Text)
        Dim type As String = "Hidden"
        Dim htmlName As String = Trim(Me.TextBoxHiddenName.Text)
        Dim size As String = Trim(Me.TextBoxHiddenSize.Text)
        Dim maxLen As String = Trim(Me.TextBoxHiddenMaxLength.Text)
        Dim dbName As String = Trim(Me.TextBoxHiddenDBName.Text)

        If validateHidden(label, htmlName, size, maxLen, dbName) = False Then Return False

        Me.htmlNames.Add(htmlName)
        Me.dbNames.Add(dbName)

        Me.metadata &= label & "|" & type & "|" & htmlName & "|" & _
                       size & "|" & maxLen & "|" & dbName & "|" & vbCrLf
        Me.clearHiddenTab()
        Return True
    End Function

    Private Sub clearHiddenTab()
        Me.TextBoxHiddenLabel.Text = ""
        Me.TextBoxHiddenName.Text = ""
        Me.TextBoxHiddenDBName.Text = ""
    End Sub


    Private Function insertSelect() As Boolean
        Dim label As String = Trim(Me.TextBoxSelLabel.Text)
        Dim type As String = "Select"
        Dim htmlName As String = Trim(Me.TextBoxSelName.Text)
        Dim width As String = Trim(Me.TextBoxSelWidth.Text)
        Dim size As String = Trim(Me.TextBoxSelSize.Text)
        Dim dbName As String = Trim(Me.TextBoxSelDBName.Text)
        Dim options As String = Trim(Me.TextBoxSelOptions.Text)

        If validateSelect(label, htmlName, width, size, dbName, options) = False Then Return False

        Me.htmlNames.Add(htmlName)
        Me.dbNames.Add(dbName)

        Do While options.EndsWith("|")
            options = MyUtil.removeEndString(options, "|")
        Loop
        Dim optionCount As Integer = Me.getOptionCount(options)

        Me.metadata &= label & "|" & type & "|" & htmlName & "|" & width & "|" & _
                       size & "|" & optionCount & "|" & dbName & "|" & options & "|" & vbCrLf
        Me.clearSelectTab()
        Return True
    End Function

    Private Sub clearSelectTab()
        Me.TextBoxSelLabel.Text = ""
        Me.TextBoxSelName.Text = ""
        Me.TextBoxSelDBName.Text = ""
        Me.TextBoxSelOptions.Text = ""
    End Sub

    Private Function getOptionCount(ByVal options As String)
        options = MyUtil.chop(options)
        'If options.EndsWith("|") Then options = MyUtil.removeEndString(options, "|")
        Dim optionArray() As String = Split(options, "|")
        Return optionArray.Length
    End Function



    Private Function insertTextArea() As Boolean
        Dim label As String = Trim(Me.TextBoxTALabel.Text)
        Dim type As String = "TextArea"
        Dim htmlName As String = Trim(Me.TextBoxTAName.Text)
        Dim rows As String = Trim(Me.TextBoxTARows.Text)
        Dim cols As String = Trim(Me.TextBoxTACols.Text)
        Dim Wrap As String
        If Me.CheckBoxTAWrap.Checked Then
            Wrap = "Wrap"
        Else
            Wrap = "NOWrap"
        End If
        Dim dbName As String = Trim(Me.TextBoxTADBName.Text)

        If validateTextArea(label, htmlName, rows, cols, dbName) = False Then Return False

        Me.htmlNames.Add(htmlName)
        Me.dbNames.Add(dbName)

        Me.metadata &= label & "|" & type & "|" & htmlName & "|" & rows & "|" & _
               cols & "|" & Wrap & "|" & dbName & "|" & vbCrLf
        Me.clearTextAreaTab()
        Return True
    End Function

    Private Sub clearTextAreaTab()
        Me.TextBoxTALabel.Text = ""
        Me.TextBoxTAName.Text = ""
        Me.TextBoxTADBName.Text = ""
    End Sub


    Private Function insertRadio() As Boolean
        Dim label As String = Trim(Me.TextBoxRadioLabel.Text)
        Dim type As String = "Radio"
        Dim htmlName As String = Trim(Me.TextBoxRadioName.Text)
        Dim dbName As String = Trim(Me.TextBoxRadioDBName.Text)
        Dim options As String = Trim(Me.TextBoxRadioOptions.Text)

        If validateRadio(label, htmlName, dbName, options) = False Then Return False

        Me.htmlNames.Add(htmlName)
        Me.dbNames.Add(dbName)

        Do While options.EndsWith("|")
            options = MyUtil.removeEndString(options, "|")
        Loop
        Dim optionCount As Integer = Me.getOptionCount(options)
        Me.metadata &= label & "|" & type & "|" & htmlName & "|" & _
                       optionCount & "|" & dbName & "|" & options & "|" & vbCrLf
        Me.clearRadioTab()
        Return True
    End Function

    Private Sub clearRadioTab()
        Me.TextBoxRadioLabel.Text = ""
        Me.TextBoxRadioName.Text = ""
        Me.TextBoxRadioDBName.Text = ""
        Me.TextBoxRadioOptions.Text = ""
    End Sub


    Private Function insertCheckbox() As Boolean
        Dim label As String = Trim(Me.TextBoxCBLabel.Text)
        Dim type As String = "Checkbox"
        Dim htmlName As String = Trim(Me.TextBoxCBName.Text)
        Dim dbName As String = Trim(Me.TextBoxCBDBName.Text)
        Dim text As String = Trim(Me.TextBoxCBText.Text)

        If validateCheckbox(label, htmlName, dbName, text) = False Then Return False

        Me.htmlNames.Add(htmlName)
        Me.dbNames.Add(dbName)

        Me.metadata &= label & "|" & type & "|" & htmlName & "|" & _
                       "Y" & "|" & dbName & "|" & text & "|" & vbCrLf
        Me.clearCheckboxTab()
        Return True
    End Function

    Private Sub clearCheckboxTab()
        Me.TextBoxCBLabel.Text = ""
        Me.TextBoxCBName.Text = ""
        Me.TextBoxCBDBName.Text = ""
        Me.TextBoxCBText.Text = ""
    End Sub


    Private Function insertComment() As Boolean
        Dim label As String = Trim(Me.TextBoxComment.Text)
        Dim type As String = "Comment"

        If validateComment(label) = False Then Return False

        Me.metadata &= label & "|" & type & "|" & vbCrLf

        Me.clearCommentTab()
        Return True
    End Function

    Private Sub clearCommentTab()
        Me.TextBoxComment.Text = ""
    End Sub

    Private Function insertHR() As Boolean
        Me.metadata &= "HR|HR|" & vbCrLf
        Me.clearHRTab()
        Return True
    End Function

    Private Sub clearHRTab()
    End Sub

    Private Function insertBlankLine() As Boolean
        Dim lineNum As String = Me.TextBoxLineNumber.Text
        Dim type As String = "BlankLine"

        If validateBlankLine(lineNum) = False Then Return False

        Me.metadata &= lineNum & "|" & type & "|" & vbCrLf

        Me.clearBlankLineTab()
        Return True
    End Function

    Private Sub clearBlankLineTab()
        Me.TextBoxLineNumber.Text = ""
    End Sub


    Private Function insertNewTable() As Boolean
        Dim firstColWidth As String = Me.TextBoxNewTableCol1Width.Text
        Dim type As String = "NewTable"

        If validateNewTable(firstColWidth) = False Then Return False

        Me.metadata &= firstColWidth & "|" & type & "|" & vbCrLf

        Me.clearNewTableTab()
        Return True
    End Function

    Private Sub clearNewTableTab()
        Me.TextBoxNewTableCol1Width.Text = ""
    End Sub

    Friend Function previewMetadata()
        Me.RichTextBox1.Text = Me.metadata
        If Me.RadioButtonLayoutCustom.Checked = True Then
            clsMD.init(Me.metadata)
            Me.htmlPreviewCode = clsMD.previewEditTable_Custom()
        Else
            clsMD.init(Me.metadata)
            tblAttr.border = Me.TextBoxTableBorder.Text
            tblAttr.width = Me.TextBoxTableWidth.Text
            tblAttr.bgColor = Me.TextBoxTableBgColor.Text
            tblAttr.firstColWidth = Me.TextBoxTableCol1Width.Text
            tblAttr.cellPadding = Me.TextBoxTablePadding.Text
            tblAttr.cellSpacing = Me.TextBoxTableSpacing.Text
            Me.htmlPreviewCode = clsMD.previewEditTable(tblAttr)
        End If
        If IOManager.folderExists(Me.previewFolderPath) Then
            IOManager.SaveTextToFile(Me.htmlPreviewCode, Me.previewFilePath)
            Me.AxWebBrowser1.Navigate(Me.previewFilePath)
        End If

        Dim aray As String() = Split(Me.metadata, Chr(10))
        Me.ListBoxEdit.Items.Clear()
        Dim i As Integer
        Dim maxI As Integer = UBound(aray) - 1
        For i = 0 To maxI
            aray(i) = MyUtil.chop(aray(i))
            If aray(i) <> "" Then Me.ListBoxEdit.Items.Add(aray(i))
        Next
    End Function


    ' Edit metadata record functions

    Private Sub btnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Dim selectedIndex As Integer = Me.lbIndex
        If selectedIndex <= 0 Then Exit Sub

        Me.ListBoxEdit.Items.Insert(selectedIndex - 1, Me.ListBoxEdit.Items.Item(selectedIndex))
        Me.ListBoxEdit.Items.RemoveAt(selectedIndex + 1)

        Me.lbIndex = selectedIndex - 1
        Me.ListBoxEdit.SelectedIndex = Me.lbIndex
        Me.btnUpdate.Enabled = True
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Dim selectedIndex As Integer = Me.lbIndex
        If selectedIndex >= Me.ListBoxEdit.Items.Count - 1 Then Exit Sub

        Me.ListBoxEdit.Items.Insert(selectedIndex, Me.ListBoxEdit.Items.Item(selectedIndex + 1))
        Me.ListBoxEdit.Items.RemoveAt(selectedIndex + 2)

        Me.lbIndex = selectedIndex + 1
        Me.ListBoxEdit.SelectedIndex = Me.lbIndex
        Me.btnUpdate.Enabled = True
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.ListBoxEdit.SelectedItem Is Nothing Then Exit Sub
        Dim selectedIndex As Integer = Me.ListBoxEdit.SelectedIndex

        Me.ListBoxEdit.Items.RemoveAt(selectedIndex)
        If selectedIndex = Me.ListBoxEdit.Items.Count Then selectedIndex -= 1
        Me.ListBoxEdit.SelectedIndex = selectedIndex

        If Me.ListBoxEdit.Items.Count = 0 Then
            Me.btnUp.Enabled = False
            Me.btnDown.Enabled = False
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.lbIndex = -1
        End If
        Me.btnUpdate.Enabled = True
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Me.updateField()
    End Sub

    Private Sub updateField()
        If Me.ListBoxEdit.SelectedItem Is Nothing Then Exit Sub
        Dim selectedIndex As Integer = Me.ListBoxEdit.SelectedIndex

        If Me.TabControlField.SelectedTab Is Me.TabPageText Then
            If Me.updateText() = False Then Exit Sub
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPagePassword Then
            If Me.updatePassword() = False Then Exit Sub
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageHidden Then
            If Me.updateHidden() = False Then Exit Sub
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageSelect Then
            If Me.updateSelect() = False Then Exit Sub
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageTextarea Then
            If Me.updateTextArea() = False Then Exit Sub
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageRadio Then
            If Me.updateRadio() = False Then Exit Sub
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageCheckbox Then
            If Me.updateCheckbox() = False Then Exit Sub
        ElseIf Me.TabControlField.SelectedTab Is Me.TabPageComment Then
            If Me.updateComment() = False Then Exit Sub
        ElseIf Me.RadioButtonLayoutDefault.Checked = True Then
            If Me.TabControlField.SelectedTab Is Me.TabPageHR Then
                If Me.updateHR() = False Then Exit Sub
            ElseIf Me.TabControlField.SelectedTab Is Me.TabPageNewTable Then
                If Me.udpateNewTable() = False Then Exit Sub
            ElseIf Me.TabControlField.SelectedTab Is Me.TabPageBlankLine Then
                If Me.updateBlankLine() = False Then Exit Sub
            End If
        End If
        'Me.btnUpdate.Enabled = False
        Me.saveEditResultToMetadata()
        Me.previewMetadata()
        Me.ListBoxEdit.SelectedIndex = selectedIndex
    End Sub

    Private Function updateText() As Boolean
        Dim label As String = Trim(Me.TextBoxTextLabel.Text)
        Dim type As String = "Text"
        Dim htmlName As String = Trim(Me.TextBoxTextName.Text)
        Dim size As String = Trim(Me.TextBoxTextSize.Text)
        Dim maxLen As String = Trim(Me.TextBoxTextMaxLength.Text)
        Dim dbName As String = Trim(Me.TextBoxTextDBName.Text)

        If Me.validateText(label, htmlName, size, maxLen, dbName, False) = False Then Return False
        Dim newValue As String = label & "|" & type & "|" & htmlName & "|" & _
                       size & "|" & maxLen & "|" & dbName & "|" & vbCrLf

        Me.ListBoxEdit.Items.Item(lbIndex) = newValue
        'MsgBox("selected index = " & lbIndex & ", new value is: " & vbCrLf & newValue & vbCrLf & "item: " & vbCrLf & Me.ListBoxEdit.Items.Item(lbIndex))
        Return True
    End Function


    Private Function updatePassword() As Boolean
        Dim label As String = Trim(Me.TextBoxPassLabel.Text)
        Dim type As String = "Password"
        Dim htmlName As String = Trim(Me.TextBoxPassName.Text)
        Dim size As String = Trim(Me.TextBoxPassSize.Text)
        Dim maxLen As String = Trim(Me.TextBoxPassMaxLength.Text)
        Dim dbName As String = Trim(Me.TextBoxPassDBName.Text)

        If Me.validateText(label, htmlName, size, maxLen, dbName, False) = False Then Return False
        Dim newValue As String = label & "|" & type & "|" & htmlName & "|" & _
                       size & "|" & maxLen & "|" & dbName & "|" & vbCrLf

        Me.ListBoxEdit.Items.Item(lbIndex) = newValue
        'MsgBox("selected index = " & lbIndex & ", new value is: " & vbCrLf & newValue & vbCrLf & "item: " & vbCrLf & Me.ListBoxEdit.Items.Item(lbIndex))
        Return True
    End Function


    Private Function updateHidden() As Boolean
        Dim label As String = Trim(Me.TextBoxHiddenLabel.Text)
        Dim type As String = "Hidden"
        Dim htmlName As String = Trim(Me.TextBoxHiddenName.Text)
        Dim size As String = Trim(Me.TextBoxHiddenSize.Text)
        Dim maxLen As String = Trim(Me.TextBoxHiddenMaxLength.Text)
        Dim dbName As String = Trim(Me.TextBoxHiddenDBName.Text)

        If Me.validateText(label, htmlName, size, maxLen, dbName, False) = False Then Return False
        Me.ListBoxEdit.Items.Item(lbIndex) = label & "|" & type & "|" & htmlName & "|" & _
                       size & "|" & maxLen & "|" & dbName & "|" & vbCrLf

        Return True
    End Function


    Private Function updateSelect() As Boolean
        Dim label As String = Trim(Me.TextBoxSelLabel.Text)
        Dim type As String = "Select"
        Dim htmlName As String = Trim(Me.TextBoxSelName.Text)
        Dim width As String = Trim(Me.TextBoxSelWidth.Text)
        Dim size As String = Trim(Me.TextBoxSelSize.Text)
        Dim dbName As String = Trim(Me.TextBoxSelDBName.Text)
        Dim options As String = Trim(Me.TextBoxSelOptions.Text)
        Do While options.EndsWith("|")
            options = MyUtil.removeEndString(options, "|")
        Loop

        If validateSelect(label, htmlName, width, size, dbName, options, False) = False Then Return False

        Dim optionCount As Integer = Me.getOptionCount(options)

        Me.ListBoxEdit.Items.Item(lbIndex) = label & "|" & type & "|" & htmlName & "|" & width & "|" & _
                       size & "|" & optionCount & "|" & dbName & "|" & options & "|" & vbCrLf
        Return True
    End Function


    Private Function updateTextArea() As Boolean
        Dim label As String = Trim(Me.TextBoxTALabel.Text)
        Dim type As String = "Textarea"
        Dim htmlName As String = Trim(Me.TextBoxTAName.Text)
        Dim rows As String = Trim(Me.TextBoxTARows.Text)
        Dim cols As String = Trim(Me.TextBoxTACols.Text)
        Dim Wrap As String
        If Me.CheckBoxTAWrap.Checked Then
            Wrap = "Wrap"
        Else
            Wrap = "NoWrap"
        End If
        Dim dbName As String = Trim(Me.TextBoxTADBName.Text)

        If validateTextArea(label, htmlName, rows, cols, dbName, False) = False Then Return False

        Me.ListBoxEdit.Items.Item(lbIndex) = label & "|" & type & "|" & htmlName & "|" & rows & "|" & _
               cols & "|" & Wrap & "|" & dbName & "|" & vbCrLf

        Return True
    End Function


    Private Function updateRadio() As Boolean
        Dim label As String = Trim(Me.TextBoxRadioLabel.Text)
        Dim type As String = "Radio"
        Dim htmlName As String = Trim(Me.TextBoxRadioName.Text)
        Dim dbName As String = Trim(Me.TextBoxRadioDBName.Text)
        Dim options As String = Trim(Me.TextBoxRadioOptions.Text)

        Do While options.EndsWith("|")
            options = MyUtil.removeEndString(options, "|")
        Loop

        If validateRadio(label, htmlName, dbName, options, False) = False Then Return False

        Dim optionCount As Integer = Me.getOptionCount(options)
        Me.ListBoxEdit.Items.Item(lbIndex) = label & "|" & type & "|" & htmlName & "|" & _
                       optionCount & "|" & dbName & "|" & options & "|" & vbCrLf
        Return True
    End Function


    Private Function updateCheckbox() As Boolean
        Dim label As String = Trim(Me.TextBoxCBLabel.Text)
        Dim type As String = "Checkbox"
        Dim htmlName As String = Trim(Me.TextBoxCBName.Text)
        Dim dbName As String = Trim(Me.TextBoxCBDBName.Text)
        Dim text As String = Trim(Me.TextBoxCBText.Text)

        If validateCheckbox(label, htmlName, dbName, text, False) = False Then Return False

        Me.ListBoxEdit.Items.Item(lbIndex) = label & "|" & type & "|" & htmlName & "|" & _
                       "Y" & "|" & dbName & "|" & text & "|" & vbCrLf
        Return True
    End Function


    Private Function updateComment() As Boolean
        Dim label As String = Trim(Me.TextBoxComment.Text)
        Dim type As String = "Comment"

        label = Replace(label, Chr(10), "\r")
        label = Replace(label, Chr(13), "\n")

        If validateComment(label) = False Then Return False

        Me.ListBoxEdit.Items.Item(lbIndex) = label & "|" & type & "|" & vbCrLf
        Return True
    End Function


    Private Function updateHR() As Boolean
        'Me.ListBoxEdit.Items.Item(lbIndex) = "HR|HR|" & vbCrLf
        Return True
    End Function

    Private Function udpateNewTable() As Boolean
        Dim firstColWidth As String = Me.TextBoxNewTableCol1Width.Text
        Dim type As String = "NewTable"

        If validateNewTable(firstColWidth) = False Then Return False

        Me.ListBoxEdit.Items.Item(lbIndex) = firstColWidth & "|" & type & "|" & vbCrLf
        Return True
    End Function

    Private Function updateBlankLine() As Boolean
        Dim lineNum As String = Me.TextBoxLineNumber.Text
        Dim type As String = "BlankLine"

        If validateBlankLine(lineNum) = False Then Return False

        Me.ListBoxEdit.Items.Item(lbIndex) = lineNum & "|" & type & "|" & vbCrLf
        Return True
    End Function


    Private Sub ListBoxEdit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBoxEdit.SelectedIndexChanged
        Me.lbIndex = Me.ListBoxEdit.SelectedIndex
        'MsgBox("selected index = " & Me.lbIndex)
        If Me.btnFinishEdit.Enabled = False Then Exit Sub

        Me.btnUp.Enabled = True
        Me.btnDown.Enabled = True
        Me.btnDelete.Enabled = True
        Me.btnUpdate.Enabled = True

        Me.populateSelectedRecord()
    End Sub


    Private Sub populateSelectedRecord()
        If Me.ListBoxEdit.SelectedItem Is Nothing Then Exit Sub

        Dim recordArray As String() = Split(Me.ListBoxEdit.SelectedItem, "|")
        'MsgBox(Me.ListBoxEdit.SelectedItem & ", " & vbCrLf & "array count = " & recordArray.Length)
        Dim recordType = LCase(recordArray(1))

        If recordType = "text" Then
            Me.TabControlField.SelectedTab = Me.TabPageText
            Me.TextBoxTextLabel.Text = recordArray(0)
            Me.TextBoxTextName.Text = recordArray(2)
            Me.TextBoxTextSize.Text = recordArray(3)
            Me.TextBoxTextMaxLength.Text = recordArray(4)
            Me.TextBoxTextDBName.Text = recordArray(5)

        ElseIf recordType = "password" Then
            Me.TabControlField.SelectedTab = Me.TabPagePassword
            Me.TextBoxPassLabel.Text = recordArray(0)
            Me.TextBoxPassName.Text = recordArray(2)
            Me.TextBoxPassSize.Text = recordArray(3)
            Me.TextBoxPassMaxLength.Text = recordArray(4)
            Me.TextBoxPassDBName.Text = recordArray(5)

        ElseIf recordType = "hidden" Then
            Me.TabControlField.SelectedTab = Me.TabPageHidden
            Me.TextBoxHiddenLabel.Text = recordArray(0)
            Me.TextBoxHiddenName.Text = recordArray(2)
            Me.TextBoxHiddenSize.Text = recordArray(3)
            Me.TextBoxHiddenMaxLength.Text = recordArray(4)
            Me.TextBoxHiddenDBName.Text = recordArray(5)

        ElseIf recordType = "select" Then
            Me.TabControlField.SelectedTab = Me.TabPageSelect
            Me.TextBoxSelLabel.Text = recordArray(0)
            Me.TextBoxSelName.Text = recordArray(2)
            Me.TextBoxSelWidth.Text = recordArray(3)
            Me.TextBoxSelSize.Text = recordArray(4)
            Me.TextBoxSelDBName.Text = recordArray(6)
            Dim i As Integer
            Dim maxI As Integer = recordArray.Length - 1
            Me.TextBoxSelOptions.Text = ""
            For i = 7 To maxI - 1
                Me.TextBoxSelOptions.Text &= recordArray(i) & "|"
            Next
            If MyUtil.chop(recordArray(i)) <> "" Then Me.TextBoxSelOptions.Text &= recordArray(i) & "|"

        ElseIf recordType = "textarea" Then
            Me.TabControlField.SelectedTab = Me.TabPageTextarea
            Me.TextBoxTALabel.Text = recordArray(0)
            Me.TextBoxTAName.Text = recordArray(2)
            Me.TextBoxTARows.Text = recordArray(3)
            Me.TextBoxTACols.Text = recordArray(4)
            If LCase(recordArray(5)) = "wrap" Then
                Me.CheckBoxTAWrap.Checked = True
            Else
                Me.CheckBoxTAWrap.Checked = False
            End If
            Me.TextBoxTADBName.Text = recordArray(6)

        ElseIf recordType = "radio" Then
            Me.TabControlField.SelectedTab = Me.TabPageRadio
            Me.TextBoxRadioLabel.Text = recordArray(0)
            Me.TextBoxRadioName.Text = recordArray(2)
            Me.TextBoxRadioDBName.Text = recordArray(4)
            Dim i As Integer
            Dim maxI As Integer = recordArray.Length - 1
            Me.TextBoxRadioOptions.Text = ""
            For i = 5 To maxI - 1
                Me.TextBoxRadioOptions.Text &= recordArray(i) & "|"
            Next
            If MyUtil.chop(recordArray(i)) <> "" Then Me.TextBoxRadioOptions.Text &= recordArray(i) & "|"

        ElseIf recordType = "checkbox" Then
            Me.TabControlField.SelectedTab = Me.TabPageCheckbox
            Me.TextBoxCBLabel.Text = recordArray(0)
            Me.TextBoxCBName.Text = recordArray(2)
            Me.TextBoxCBDBName.Text = recordArray(4)
            Me.TextBoxCBText.Text = recordArray(5)

        ElseIf recordType = "comment" Then
            Me.TabControlField.SelectedTab = Me.TabPageComment
            Me.TextBoxComment.Text = recordArray(0)

        ElseIf recordType = "hr" Then
            Me.TabControlField.SelectedTab = Me.TabPageHR

        ElseIf recordType = "newtable" Then
            Me.TabControlField.SelectedTab = Me.TabPageNewTable
            Me.TextBoxNewTableCol1Width.Text = recordArray(0)

        ElseIf recordType = "blankline" Then
            Me.TabControlField.SelectedTab = Me.TabPageBlankLine
            Me.TextBoxLineNumber.Text = recordArray(0)
        End If
    End Sub

    Private Sub btnSaveFreeEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveFreeEdit.Click
        Me.metadata = Me.RichTextBox1.Text
        Me.btnSaveFreeEdit.Enabled = False
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.btnSaveFreeEdit.Enabled = True
    End Sub

    Private Sub RadioButtonLayoutDefault_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonLayoutDefault.CheckedChanged
        If Me.RadioButtonLayoutDefault.Checked = True Then
            Me.GroupBoxTableAttr.Enabled = True
            Me.TabPageHR.Enabled = True
            Me.TabPageNewTable.Enabled = True
            Me.TabPageBlankLine.Enabled = True
        Else
            Me.GroupBoxTableAttr.Enabled = False
            Me.TabPageHR.Enabled = False
            Me.TabPageNewTable.Enabled = False
            Me.TabPageBlankLine.Enabled = False
        End If
        Me.previewMetadata()
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Me.btnFinish.Enabled = False
        Me.btnInsert.Enabled = False
        Me.btnEdit.Enabled = False

        Me.TabControlView.SelectedTab = Me.TabPageEdit
        Me.btnFinishEdit.Enabled = True
        'Me.btnUpdate.Enabled = True
        If Not Me.ListBoxEdit.SelectedItem Is Nothing Then
            Me.btnUp.Enabled = True
            Me.btnDown.Enabled = True
            Me.btnUpdate.Enabled = True
            Me.btnDelete.Enabled = True
            Me.populateSelectedRecord()
        End If
    End Sub

    Private Sub btnFinishEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinishEdit.Click
        ' update the latest field in case it was changed.
        Me.updateField()

        Me.clearAllFieldTabs()

        Me.btnFinish.Enabled = True
        Me.btnInsert.Enabled = True
        Me.btnEdit.Enabled = True

        Me.btnFinishEdit.Enabled = False
        Me.btnUp.Enabled = False
        Me.btnDown.Enabled = False
        Me.btnUpdate.Enabled = False
        Me.btnDelete.Enabled = False
    End Sub

    Private Sub saveEditResultToMetadata()
        Dim str As String = ""
        Dim i As Integer
        Dim maxI As Integer = Me.ListBoxEdit.Items.Count - 1
        For i = 0 To maxI
            str = str & Trim(Me.ListBoxEdit.Items.Item(i)) & Chr(10) & Chr(13)
        Next
        Me.metadata = str
    End Sub


    '''''''''''
    ' Validate functions
    '''''''''''

    Private Function htmlNameExists(ByVal htmlName As String) As Boolean
        If Not MyUtil.arrayListContainsString(Me.htmlNames, htmlName, False) = -1 Then
            MsgBox("The HTML Name '" & htmlName & "' is already taken")
            Return True
        End If
        Return False
    End Function

    Private Function dbNameExists(ByVal dbName As String) As Boolean
        If Not MyUtil.arrayListContainsString(Me.dbNames, dbName, False) = -1 Then
            MsgBox("The Database Name '" & dbName & "' is already taken")
            Return True
        End If
        Return False
    End Function


    Private Function validateText(ByVal label As String, _
    ByVal htmlName As String, ByVal size As String, _
    ByVal maxLen As String, ByVal dbName As String, _
    Optional ByVal insert As Boolean = True) As Boolean

        If label = "" Or htmlName = "" Or size = "" Or maxLen = "" Or dbName = "" Then
            MsgBox("All fields must be filled.")
            Return False
        End If
        Try
            If CType(size, Integer) > CType(maxLen, Integer) Then
                MsgBox("size cannot be larger than maxLength")
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

        If insert = True Then
            If htmlNameExists(htmlName) Then Return False
            If dbNameExists(dbName) Then Return False
        End If

        Return True
    End Function


    Private Function validatePassword(ByVal label As String, _
            ByVal htmlName As String, ByVal size As String, _
            ByVal maxLen As String, ByVal dbName As String, _
            Optional ByVal insert As Boolean = True) As Boolean

        If label = "" Or htmlName = "" Or size = "" Or maxLen = "" Or dbName = "" Then
            MsgBox("All fields must be filled.")
            Return False
        End If
        Try
            If CType(size, Integer) > CType(maxLen, Integer) Then
                MsgBox("size cannot be larger than maxLength")
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

        If insert = True Then
            If htmlNameExists(htmlName) Then Return False
            If dbNameExists(dbName) Then Return False
        End If

        Return True
    End Function


    Private Function validateHidden(ByVal label As String, _
        ByVal htmlName As String, ByVal size As String, _
        ByVal maxLen As String, ByVal dbName As String, _
        Optional ByVal insert As Boolean = True) As Boolean

        If label = "" Or htmlName = "" Or size = "" Or maxLen = "" Or dbName = "" Then
            MsgBox("All fields must be filled.")
            Return False
        End If
        Try
            If CType(size, Integer) > CType(maxLen, Integer) Then
                MsgBox("size cannot be larger than maxLength")
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

        If insert = True Then
            If htmlNameExists(htmlName) Then Return False
            If dbNameExists(dbName) Then Return False
        End If

        Return True
    End Function


    Private Function validateSelect(ByVal label As String, _
            ByVal htmlName As String, ByVal width As String, _
            ByVal size As String, ByVal dbName As String, _
            ByVal options As String, Optional ByVal insert As Boolean = True) As Boolean

        If label = "" Or htmlName = "" Or size = "" Or width = "" Or dbName = "" Or options = "" Then
            MsgBox("All fields must be filled.")
            Return False
        End If
        Try
            Dim _size As Integer = CType(size, Integer)
            Dim _width As Integer = CType(width, Integer)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

        If insert = True Then
            If htmlNameExists(htmlName) Then Return False
            If dbNameExists(dbName) Then Return False
        End If

        Return True
    End Function


    Private Function validateTextArea(ByVal label As String, _
        ByVal htmlName As String, ByVal rows As String, _
        ByVal cols As String, ByVal dbName As String, _
        Optional ByVal insert As Boolean = True) As Boolean

        If label = "" Or htmlName = "" Or rows = "" Or cols = "" Or dbName = "" Then
            MsgBox("All fields must be filled.")
            Return False
        End If
        Try
            Dim _rows As Integer = CType(rows, Integer)
            Dim _cols As Integer = CType(cols, Integer)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

        If insert = True Then
            If htmlNameExists(htmlName) Then Return False
            If dbNameExists(dbName) Then Return False
        End If

        Return True
    End Function


    Private Function validateRadio(ByVal label As String, _
            ByVal htmlName As String, ByVal dbName As String, _
            ByVal options As String, Optional ByVal insert As Boolean = True) As Boolean

        If label = "" Or htmlName = "" Or dbName = "" Or options = "" Then
            MsgBox("All fields must be filled.")
            Return False
        End If

        If insert = True Then
            If htmlNameExists(htmlName) Then Return False
            If dbNameExists(dbName) Then Return False
        End If

        Return True
    End Function


    Private Function validateCheckbox(ByVal label As String, _
         ByVal htmlName As String, ByVal dbName As String, _
         ByVal text As String, Optional ByVal insert As Boolean = True) As Boolean

        If label = "" Or htmlName = "" Or dbName = "" Or text = "" Then
            MsgBox("All fields must be filled.")
            Return False
        End If

        If insert = True Then
            If htmlNameExists(htmlName) Then Return False
            If dbNameExists(dbName) Then Return False
        End If

        Return True
    End Function


    Private Function validateComment(ByVal label As String) As Boolean
        If label = "" Then
            MsgBox("All fields must be filled.")
            Return False
        End If

        Return True
    End Function


    Private Function validateHR() As Boolean
        Return True
    End Function


    Private Function validateBlankLine(ByVal lineNumber As String) As Boolean
        If lineNumber = "" Then
            MsgBox("Line Number must be filled.")
            Return False
        End If
        Try
            Dim lineNum As Integer = CType(lineNumber, Integer)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

        Return True
    End Function


    Private Function validateNewTable(ByVal firstColWidth As String) As Boolean
        If firstColWidth.EndsWith("%") Then
            MyUtil.removeEndString(firstColWidth, "%")
        End If
        Try
            Dim w As Integer
            If firstColWidth <> "" Then w = CType(firstColWidth, Integer)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

        Return True
    End Function


    ' FUnctions to clear relevant fields after a record is inserted.
    Private Sub clearAllFieldTabs()
        Me.clearBlankLineTab()
        Me.clearCheckboxTab()
        Me.clearCommentTab()
        Me.clearHiddenTab()
        Me.clearHRTab()
        Me.clearNewTableTab()
        Me.clearPasswordTab()
        Me.clearRadioTab()
        Me.clearSelectTab()
        Me.clearTextAreaTab()
        Me.clearTextTab()
    End Sub


    Private Sub btnFinish_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinish.Click
        saveMetadata()
        Me.Close()
    End Sub

    Private Function saveMetadata()
        'MsgBox("set metadata: " & metadata)
        ownerForm.setMetadata(Me.metadata, Me.tblAttr)
    End Function


    '''''''''''''''''''''''
    Private Sub handleTablePropertyChange(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Me.previewMetadata()
        End If
    End Sub

    Private Sub TextBoxTableWidth_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxTableWidth.KeyPress
        Me.handleTablePropertyChange(e)
    End Sub

    Private Sub TextBoxTableBorder_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxTableBorder.KeyPress
        Me.handleTablePropertyChange(e)
    End Sub

    Private Sub TextBoxTableBgColor_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxTableBgColor.KeyPress
        Me.handleTablePropertyChange(e)
    End Sub

    Private Sub TextBoxTablePadding_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxTablePadding.KeyPress
        Me.handleTablePropertyChange(e)
    End Sub

    Private Sub TextBoxTableSpacing_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxTableSpacing.KeyPress
        Me.handleTablePropertyChange(e)
    End Sub

    Private Sub TextBoxTableCol1Width_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxTableCol1Width.KeyPress
        Me.handleTablePropertyChange(e)
    End Sub

    Private Sub ButtonTblAttr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonTblAttr.Click
        Dim frmTableAttr As New FormTblAttr()
        frmTableAttr.init(Me, Me.TextBoxTableWidth)
        frmTableAttr.ShowDialog(Me)
    End Sub
End Class
