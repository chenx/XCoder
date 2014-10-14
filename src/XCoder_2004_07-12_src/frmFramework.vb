Public Class frmFramework
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
    Friend WithEvents Panel1GenFW As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtFWName As System.Windows.Forms.TextBox
    Friend WithEvents btnGenRWNext1 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtODBC As System.Windows.Forms.TextBox
    Friend WithEvents txtODBCLogin As System.Windows.Forms.TextBox
    Friend WithEvents txtODBCPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtGenFWOutputPath As System.Windows.Forms.TextBox
    Friend WithEvents Panel3GenFW As System.Windows.Forms.Panel
    Friend WithEvents lb3GenFW As System.Windows.Forms.Label
    Friend WithEvents btn3GenFWBack As System.Windows.Forms.Button
    Friend WithEvents Panel2GenFW As System.Windows.Forms.Panel
    Friend WithEvents lb2GenFW As System.Windows.Forms.Label
    Friend WithEvents btnGenFWNext2 As System.Windows.Forms.Button
    Friend WithEvents btnGenFWBack2 As System.Windows.Forms.Button
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents lblNote As System.Windows.Forms.Label
    Friend WithEvents TreeViewFileList As System.Windows.Forms.TreeView
    Friend WithEvents AxWebBrowser1 As AxSHDocVw.AxWebBrowser
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFramework))
        Me.Panel1GenFW = New System.Windows.Forms.Panel()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtFWName = New System.Windows.Forms.TextBox()
        Me.btnGenRWNext1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtODBC = New System.Windows.Forms.TextBox()
        Me.txtODBCLogin = New System.Windows.Forms.TextBox()
        Me.txtODBCPassword = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtGenFWOutputPath = New System.Windows.Forms.TextBox()
        Me.Panel3GenFW = New System.Windows.Forms.Panel()
        Me.lb3GenFW = New System.Windows.Forms.Label()
        Me.btn3GenFWBack = New System.Windows.Forms.Button()
        Me.Panel2GenFW = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.AxWebBrowser1 = New AxSHDocVw.AxWebBrowser()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.TreeViewFileList = New System.Windows.Forms.TreeView()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu()
        Me.btnGenFWNext2 = New System.Windows.Forms.Button()
        Me.btnGenFWBack2 = New System.Windows.Forms.Button()
        Me.lblNote = New System.Windows.Forms.Label()
        Me.lb2GenFW = New System.Windows.Forms.Label()
        Me.Panel1GenFW.SuspendLayout()
        Me.Panel3GenFW.SuspendLayout()
        Me.Panel2GenFW.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1GenFW
        '
        Me.Panel1GenFW.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label5, Me.txtFWName, Me.btnGenRWNext1, Me.Label2, Me.Label1, Me.txtODBC, Me.txtODBCLogin, Me.txtODBCPassword, Me.Label3, Me.Label4, Me.txtGenFWOutputPath})
        Me.Panel1GenFW.Location = New System.Drawing.Point(152, 24)
        Me.Panel1GenFW.Name = "Panel1GenFW"
        Me.Panel1GenFW.Size = New System.Drawing.Size(688, 552)
        Me.Panel1GenFW.TabIndex = 10
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(40, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(160, 24)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Framework Name:"
        '
        'txtFWName
        '
        Me.txtFWName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFWName.Location = New System.Drawing.Point(232, 88)
        Me.txtFWName.Name = "txtFWName"
        Me.txtFWName.Size = New System.Drawing.Size(160, 26)
        Me.txtFWName.TabIndex = 2
        Me.txtFWName.Text = "Tests"
        '
        'btnGenRWNext1
        '
        Me.btnGenRWNext1.Enabled = False
        Me.btnGenRWNext1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGenRWNext1.Location = New System.Drawing.Point(504, 392)
        Me.btnGenRWNext1.Name = "btnGenRWNext1"
        Me.btnGenRWNext1.Size = New System.Drawing.Size(112, 32)
        Me.btnGenRWNext1.TabIndex = 6
        Me.btnGenRWNext1.Text = "Next"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(80, 208)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(144, 24)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "ODBC Password:"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(80, 144)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 24)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "ODBC DSN:"
        '
        'txtODBC
        '
        Me.txtODBC.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtODBC.Location = New System.Drawing.Point(232, 144)
        Me.txtODBC.Name = "txtODBC"
        Me.txtODBC.Size = New System.Drawing.Size(160, 26)
        Me.txtODBC.TabIndex = 3
        Me.txtODBC.Text = "a"
        '
        'txtODBCLogin
        '
        Me.txtODBCLogin.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtODBCLogin.Location = New System.Drawing.Point(232, 176)
        Me.txtODBCLogin.Name = "txtODBCLogin"
        Me.txtODBCLogin.Size = New System.Drawing.Size(160, 26)
        Me.txtODBCLogin.TabIndex = 4
        Me.txtODBCLogin.Text = "a"
        '
        'txtODBCPassword
        '
        Me.txtODBCPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtODBCPassword.Location = New System.Drawing.Point(232, 208)
        Me.txtODBCPassword.Name = "txtODBCPassword"
        Me.txtODBCPassword.Size = New System.Drawing.Size(160, 26)
        Me.txtODBCPassword.TabIndex = 5
        Me.txtODBCPassword.Text = "a"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(80, 176)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 24)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "ODBC Login:"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 32)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 32)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Output path:"
        '
        'txtGenFWOutputPath
        '
        Me.txtGenFWOutputPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGenFWOutputPath.Location = New System.Drawing.Point(136, 32)
        Me.txtGenFWOutputPath.Name = "txtGenFWOutputPath"
        Me.txtGenFWOutputPath.ReadOnly = True
        Me.txtGenFWOutputPath.Size = New System.Drawing.Size(536, 26)
        Me.txtGenFWOutputPath.TabIndex = 1
        Me.txtGenFWOutputPath.Text = ""
        '
        'Panel3GenFW
        '
        Me.Panel3GenFW.Controls.AddRange(New System.Windows.Forms.Control() {Me.lb3GenFW, Me.btn3GenFWBack})
        Me.Panel3GenFW.Location = New System.Drawing.Point(152, 24)
        Me.Panel3GenFW.Name = "Panel3GenFW"
        Me.Panel3GenFW.Size = New System.Drawing.Size(688, 552)
        Me.Panel3GenFW.TabIndex = 12
        Me.Panel3GenFW.Visible = False
        '
        'lb3GenFW
        '
        Me.lb3GenFW.Location = New System.Drawing.Point(32, 32)
        Me.lb3GenFW.Name = "lb3GenFW"
        Me.lb3GenFW.Size = New System.Drawing.Size(616, 352)
        Me.lb3GenFW.TabIndex = 3
        '
        'btn3GenFWBack
        '
        Me.btn3GenFWBack.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn3GenFWBack.Location = New System.Drawing.Point(384, 480)
        Me.btn3GenFWBack.Name = "btn3GenFWBack"
        Me.btn3GenFWBack.Size = New System.Drawing.Size(248, 32)
        Me.btn3GenFWBack.TabIndex = 2
        Me.btn3GenFWBack.Text = "Back To First Page"
        '
        'Panel2GenFW
        '
        Me.Panel2GenFW.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel3, Me.Panel2, Me.TabControl1, Me.TreeViewFileList, Me.Panel1})
        Me.Panel2GenFW.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2GenFW.Name = "Panel2GenFW"
        Me.Panel2GenFW.Size = New System.Drawing.Size(880, 696)
        Me.Panel2GenFW.TabIndex = 11
        Me.Panel2GenFW.Visible = False
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Panel3.Location = New System.Drawing.Point(16, 8)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(168, 104)
        Me.Panel3.TabIndex = 51
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel2.Location = New System.Drawing.Point(16, 120)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(168, 552)
        Me.Panel2.TabIndex = 50
        '
        'TabControl1
        '
        Me.TabControl1.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPage1, Me.TabPage2})
        Me.TabControl1.Location = New System.Drawing.Point(200, 8)
        Me.TabControl1.Multiline = True
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.ShowToolTips = True
        Me.TabControl1.Size = New System.Drawing.Size(336, 416)
        Me.TabControl1.TabIndex = 48
        '
        'TabPage1
        '
        Me.TabPage1.Controls.AddRange(New System.Windows.Forms.Control() {Me.AxWebBrowser1})
        Me.TabPage1.Location = New System.Drawing.Point(4, 4)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(328, 387)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Quick View"
        '
        'AxWebBrowser1
        '
        Me.AxWebBrowser1.ContainingControl = Me
        Me.AxWebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AxWebBrowser1.Enabled = True
        Me.AxWebBrowser1.OcxState = CType(resources.GetObject("AxWebBrowser1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxWebBrowser1.Size = New System.Drawing.Size(328, 387)
        Me.AxWebBrowser1.TabIndex = 47
        '
        'TabPage2
        '
        Me.TabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.RichTextBox1})
        Me.TabPage2.Location = New System.Drawing.Point(4, 4)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(328, 387)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Source"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(328, 387)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.WordWrap = False
        '
        'TreeViewFileList
        '
        Me.TreeViewFileList.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeViewFileList.FullRowSelect = True
        Me.TreeViewFileList.ImageList = Me.ImageList1
        Me.TreeViewFileList.LabelEdit = True
        Me.TreeViewFileList.Location = New System.Drawing.Point(544, 8)
        Me.TreeViewFileList.Name = "TreeViewFileList"
        Me.TreeViewFileList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TreeViewFileList.Size = New System.Drawing.Size(328, 416)
        Me.TreeViewFileList.TabIndex = 45
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.ContextMenu = Me.ContextMenu1
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnGenFWNext2, Me.btnGenFWBack2, Me.lblNote, Me.lb2GenFW})
        Me.Panel1.ImeMode = System.Windows.Forms.ImeMode.On
        Me.Panel1.Location = New System.Drawing.Point(200, 440)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(672, 232)
        Me.Panel1.TabIndex = 49
        Me.Panel1.Tag = "a"
        '
        'btnGenFWNext2
        '
        Me.btnGenFWNext2.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.btnGenFWNext2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGenFWNext2.Location = New System.Drawing.Point(544, 184)
        Me.btnGenFWNext2.Name = "btnGenFWNext2"
        Me.btnGenFWNext2.Size = New System.Drawing.Size(112, 32)
        Me.btnGenFWNext2.TabIndex = 2
        Me.btnGenFWNext2.Text = "Next"
        '
        'btnGenFWBack2
        '
        Me.btnGenFWBack2.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.btnGenFWBack2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGenFWBack2.Location = New System.Drawing.Point(416, 184)
        Me.btnGenFWBack2.Name = "btnGenFWBack2"
        Me.btnGenFWBack2.Size = New System.Drawing.Size(112, 32)
        Me.btnGenFWBack2.TabIndex = 1
        Me.btnGenFWBack2.Text = "Back"
        '
        'lblNote
        '
        Me.lblNote.Location = New System.Drawing.Point(152, 96)
        Me.lblNote.Name = "lblNote"
        Me.lblNote.Size = New System.Drawing.Size(472, 64)
        Me.lblNote.TabIndex = 46
        '
        'lb2GenFW
        '
        Me.lb2GenFW.Location = New System.Drawing.Point(152, 32)
        Me.lb2GenFW.Name = "lb2GenFW"
        Me.lb2GenFW.Size = New System.Drawing.Size(480, 48)
        Me.lb2GenFW.TabIndex = 3
        '
        'frmFramework
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(880, 696)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel2GenFW, Me.Panel1GenFW, Me.Panel3GenFW})
        Me.Name = "frmFramework"
        Me.Text = "frmFramework"
        Me.TopMost = True
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1GenFW.ResumeLayout(False)
        Me.Panel3GenFW.ResumeLayout(False)
        Me.Panel2GenFW.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private defaultOutputPath As String
    Private curOutputPath As String

    ' File and folder names
    Private indexFile As String = "index.asp"
    Private cssFile As String = "xc.asp"
    Private adminFolder As String = "admin"
    Private adminHomeFile As String = "index.asp"
    Private adminApproveFile As String = "approveAdmin.asp"
    Private aspBinFolder As String = "asp-bin"
    Private aspBinAdovbsFile As String = "adovbs.inc"
    Private aspBinDBConfigFile As String = "db_config.asp"
    Private aspBinAspFuncLibFile As String = "aspFuncLib.asp"
    Private loginFolder As String = "login"
    Private loginLoginFile As String = "login.asp"
    Private loginLogoutFile As String = "logout.asp"
    Private loginAdminSecFile As String = "admin_Security.asp"
    Private loginMemSecFile As String = "mem_security.asp"
    Private memberFolder As String = "member"
    Private memberHomeFile As String = "index.asp"
    Private rootIncFolder As String = "rootInclude"
    Private rootIncHeaderFile As String = "header.asp"
    Private rootIncFooterFile As String = "footer.asp"
    Private imagesFolder As String = "images"

    ' Treeview nodes
    Private indexFileNode As New TreeNode("index.asp", 0, 0)
    Private cssFileNode As New TreeNode("xc.asp", 0, 0)
    Private adminFolderNode As New TreeNode("admin", 1, 1)
    Private adminHomeFileNode As New TreeNode("index.asp", 0, 0)
    Private adminApproveFileNode As New TreeNode("approveAdmin.asp", 0, 0)
    Private aspBinFolderNode As New TreeNode("asp-bin", 1, 1)
    Private aspBinAdovbsFileNode As New TreeNode("adovbs.inc", 0, 0)
    Private aspBinDBConfigFileNode As New TreeNode("db_config.asp", 0, 0)
    Private aspBinAspFuncLibFileNode As New TreeNode("aspFuncLib.asp", 0, 0)
    Private loginFolderNode As New TreeNode("login", 1, 1)
    Private loginLoginFileNode As New TreeNode("login.asp", 0, 0)
    Private loginLogoutFileNode As New TreeNode("logout.asp", 0, 0)
    Private loginAdminSecFileNode As New TreeNode("admin_Security.asp", 0, 0)
    Private loginMemSecFileNode As New TreeNode("mem_security.asp", 0, 0)
    Private memberFolderNode As New TreeNode("member", 1, 1)
    Private memberHomeFileNode As New TreeNode("index.asp", 0, 0)
    Private rootIncFolderNode As New TreeNode("rootInclude", 1, 1)
    Private rootIncHeaderFileNode As New TreeNode("header.asp", 0, 0)
    Private rootIncFooterFileNode As New TreeNode("footer.asp", 0, 0)
    Private imagesFolderNode As New TreeNode("images", 1, 1)

    Private Sub setCurOutputPath(ByVal relativePath As String)
        Me.curOutputPath = defaultOutputPath & relativePath
    End Sub


    Private Sub frmFramework_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Panel1GenFW.Visible = True
        Me.defaultOutputPath = SetupManager.OUTPUT_PATH
        Me.setCurOutputPath("")
        Me.txtGenFWOutputPath.Text = Me.curOutputPath
        Me.txtFWName.Select()
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Generate Framework Functions

    Private Sub txtODBCPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtODBCPassword.TextChanged
        tryEnableBtn1GenFW()
    End Sub

    Private Sub txtODBC_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtODBC.TextChanged
        tryEnableBtn1GenFW()
    End Sub

    Private Sub txtODBCLogin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtODBCLogin.TextChanged
        tryEnableBtn1GenFW()
    End Sub

    Private Sub txtFWName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFWName.TextChanged
        tryEnableBtn1GenFW()
        Me.setCurOutputPath(Me.txtFWName.Text)
        Me.txtGenFWOutputPath.Text = Me.curOutputPath
    End Sub

    Private Sub tryEnableBtn1GenFW()
        If Me.txtODBC.Text <> "" And Me.txtODBCLogin.Text <> "" And _
           Me.txtODBCPassword.Text <> "" And Me.txtFWName.Text <> "" Then
            Me.btnGenRWNext1.Enabled = True
        Else
            Me.btnGenRWNext1.Enabled = False
        End If
    End Sub




    Private Sub btnGenRWNext1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenRWNext1.Click
        If IOManager.folderExists(Me.curOutputPath) Then
            If MessageBox.Show("This output folder already exists, do you want to overwrite?", "XinCoder - Create Web Site Framework", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                ' do not overwrite
                Me.txtFWName.Select()
                Me.txtFWName.SelectAll()
                Exit Sub
            End If
        End If

        ''''''''''''''''''''''''''''''''''''''
        Me.constructTreeViewFileList(Me.TreeViewFileList)
        ''''''''''''''''''''''''''''''''''''''

        Me.hideGenFWPanels()
        Me.Panel2GenFW.Visible = True
        Me.lb2GenFW.Text = _
            "Now the web site framwork will be generated. " & vbCrLf & _
            "A summary of files to be generated is shown below."

    End Sub

    Private Sub btnGenFWBack2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenFWBack2.Click
        Me.hideGenFWPanels()
        Me.Panel1GenFW.Visible = True
    End Sub

    Private Sub btnGenFWNext2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenFWNext2.Click
        Call generateFramework()
        Me.hideGenFWPanels()
        Me.Panel3GenFW.Visible = True
        Me.lb3GenFW.Text = _
            "Congradulations! You have successfully created the web site frame work at " & _
            vbCrLf & Me.curOutputPath & vbCrLf & _
            "Now you can upload it to the server, and go back to the main wizard page."
    End Sub


    ''''
    ' Generate Web Site Framework
    '
    Private Sub generateFramework()
        If Me.curOutputPath.EndsWith("\") = False Then
            Me.curOutputPath &= "\"
        End If
        IOManager.CreateFolder(Me.curOutputPath & "admin\")
        IOManager.SaveTextToFile("", Me.curOutputPath & "admin\" & "index.asp")
        IOManager.SaveTextToFile("", Me.curOutputPath & "admin\" & "approveAdmin.asp")

        IOManager.CreateFolder(Me.curOutputPath & "asp-bin\")
        IOManager.SaveTextToFile("", Me.curOutputPath & "asp-bin\" & "adovbs.inc")
        IOManager.SaveTextToFile("", Me.curOutputPath & "asp-bin\" & "aspFuncLib.asp")
        IOManager.SaveTextToFile("", Me.curOutputPath & "asp-bin\" & "database_config.asp")

        IOManager.CreateFolder(Me.curOutputPath & "rootInclude\")
        IOManager.SaveTextToFile("", Me.curOutputPath & "rootInclude\" & "header.asp")
        IOManager.SaveTextToFile("", Me.curOutputPath & "rootInclude\" & "footer.asp")

        IOManager.CreateFolder(Me.curOutputPath & "member\")
        IOManager.SaveTextToFile("", Me.curOutputPath & "member\" & "index.asp")

        IOManager.CreateFolder(Me.curOutputPath & "login\")
        IOManager.SaveTextToFile("", Me.curOutputPath & "login\" & "login.asp")
        IOManager.SaveTextToFile("", Me.curOutputPath & "login\" & "logoff.asp")
        IOManager.SaveTextToFile("", Me.curOutputPath & "login\" & "admin_security.asp")
        IOManager.SaveTextToFile("", Me.curOutputPath & "login\" & "mem_security.asp")

        IOManager.CreateFolder(Me.curOutputPath & "images\")

        IOManager.SaveTextToFile("", Me.curOutputPath & "xc.css")
        IOManager.SaveTextToFile(getFWIndexPage(), Me.curOutputPath & "index.asp")
    End Sub

    Private Function getFWIndexPage() As String
        Dim str As String
        str = "Hello!"
        Return str
    End Function

    Private Sub btn3GenFWBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn3GenFWBack.Click
        Me.hideGenFWPanels()
        Me.Panel1GenFW.Visible = True
    End Sub

    Private Sub hideGenFWPanels()
        Me.Panel1GenFW.Visible = False
        Me.Panel2GenFW.Visible = False
        Me.Panel3GenFW.Visible = False
    End Sub


    ' Do not allow the user to edit the root node.
    Private Sub trvOrg_BeforeLabelEdit(ByVal sender As Object, _
        ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) _
        Handles TreeViewFileList.BeforeLabelEdit
        e.CancelEdit = (e.Node.Text = Me.curOutputPath)
    End Sub


    Private Sub initializeTreeViewNodes()
        indexFileNode = New TreeNode("index.asp", 0, 0)
        cssFileNode = New TreeNode("xc.css", 0, 0)
        adminFolderNode = New TreeNode("admin", 1, 1)
        adminHomeFileNode = New TreeNode("index.asp", 0, 0)
        adminApproveFileNode = New TreeNode("approveAdmin.asp", 0, 0)
        aspBinFolderNode = New TreeNode("asp-bin", 1, 1)
        aspBinAdovbsFileNode = New TreeNode("adovbs.inc", 0, 0)
        aspBinDBConfigFileNode = New TreeNode("db_config.asp", 0, 0)
        aspBinAspFuncLibFileNode = New TreeNode("aspFuncLib.asp", 0, 0)
        loginFolderNode = New TreeNode("login", 1, 1)
        loginLoginFileNode = New TreeNode("login.asp", 0, 0)
        loginLogoutFileNode = New TreeNode("logout.asp", 0, 0)
        loginAdminSecFileNode = New TreeNode("admin_Security.asp", 0, 0)
        loginMemSecFileNode = New TreeNode("mem_security.asp", 0, 0)
        memberFolderNode = New TreeNode("member", 1, 1)
        memberHomeFileNode = New TreeNode("index.asp", 0, 0)
        rootIncFolderNode = New TreeNode("rootInclude", 1, 1)
        rootIncHeaderFileNode = New TreeNode("header.asp", 0, 0)
        rootIncFooterFileNode = New TreeNode("footer.asp", 0, 0)
        imagesFolderNode = New TreeNode("images", 1, 1)
    End Sub

    Friend Sub constructTreeViewFileList(ByRef TreeViewFileList As TreeView)
        Try
            TreeViewFileList.Nodes.Clear()
            initializeTreeViewNodes()

            Me.rootIncFolderNode.Nodes.Add(Me.rootIncHeaderFileNode)
            Me.rootIncFolderNode.Nodes.Add(Me.rootIncFooterFileNode)

            Me.adminFolderNode.Nodes.Add(Me.adminHomeFileNode)
            Me.adminFolderNode.Nodes.Add(Me.adminApproveFileNode)

            Me.aspBinFolderNode.Nodes.Add(Me.aspBinAdovbsFileNode)
            Me.aspBinFolderNode.Nodes.Add(Me.aspBinDBConfigFileNode)
            Me.aspBinFolderNode.Nodes.Add(Me.aspBinAspFuncLibFileNode)

            Me.loginFolderNode.Nodes.Add(Me.loginLoginFileNode)
            Me.loginFolderNode.Nodes.Add(Me.loginLogoutFileNode)
            Me.loginFolderNode.Nodes.Add(Me.loginAdminSecFileNode)
            Me.loginFolderNode.Nodes.Add(Me.loginMemSecFileNode)

            Me.memberFolderNode.Nodes.Add(Me.memberHomeFileNode)

            Dim rootNode As New TreeNode(Me.curOutputPath, 1, 1)

            rootNode.Nodes.Add(Me.indexFileNode)
            rootNode.Nodes.Add(Me.cssFileNode)
            rootNode.Nodes.Add(Me.adminFolderNode)
            rootNode.Nodes.Add(Me.aspBinFolderNode)
            rootNode.Nodes.Add(Me.loginFolderNode)
            rootNode.Nodes.Add(Me.memberFolderNode)
            rootNode.Nodes.Add(Me.rootIncFolderNode)
            rootNode.Nodes.Add(Me.imagesFolderNode)

            rootNode.ExpandAll()

            TreeViewFileList.Nodes.Add(rootNode)
        Catch ex As Exception
            MsgBox("Cannot create initial node:" & ex.ToString)
            End
        End Try

        showAsHTML()
    End Sub

    Private Sub showAsHTML()
        Dim zero As Integer = 0
        Dim oZero As Object = zero
        Dim emptyString As String = ""
        Dim oEmptyString As Object = emptyString
        'AxWebBrowser1.Navigate("a.html", oZero, oEmptyString, _
        '                       oEmptyString, oEmptyString)
        Dim file As String = SetupManager.ROOT_PATH & "bin/a.html"
        AxWebBrowser1.Navigate(file)
        Me.RichTextBox1.Text = IOManager.GetFileContents(file)
        'AxWebBrowser1.Text = "a.html"
    End Sub

    Private Sub assignFileNames()
        Me.indexFile = Me.indexFileNode.Text
        Me.cssFile = Me.cssFileNode.Text

        Me.adminFolder = Me.adminFolderNode.Text
        Me.adminHomeFile = Me.adminHomeFileNode.Text
        Me.adminApproveFile = Me.adminApproveFileNode.Text

        Me.aspBinFolder = Me.aspBinFolderNode.Text
        Me.aspBinAdovbsFile = Me.aspBinAdovbsFileNode.Text
        Me.aspBinDBConfigFile = Me.aspBinDBConfigFileNode.Text
        Me.aspBinAspFuncLibFile = Me.aspBinAspFuncLibFileNode.Text

        Me.loginFolder = Me.loginFolderNode.Text
        Me.loginLoginFile = Me.loginLoginFileNode.Text
        Me.loginLogoutFile = Me.loginLogoutFileNode.Text
        Me.loginAdminSecFile = Me.loginAdminSecFileNode.Text
        Me.loginMemSecFile = Me.loginMemSecFileNode.Text

        Me.memberFolder = Me.memberFolderNode.Text
        Me.memberHomeFile = Me.memberHomeFileNode.Text

        Me.rootIncFolder = Me.rootIncFolderNode.Text
        Me.rootIncHeaderFile = Me.rootIncHeaderFileNode.Text
        Me.rootIncFooterFile = Me.rootIncFooterFileNode.Text

        Me.imagesFolder = Me.imagesFolderNode.Text
    End Sub


    Private Sub TreeViewFileList_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeViewFileList.AfterSelect
        Me.lblNote.Text = TreeViewFileList.SelectedNode.Text & vbCrLf & vbCrLf _
            & "Explanations."
    End Sub


    Private Sub TreeViewFileList_AfterLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TreeViewFileList.AfterLabelEdit
        If e.Label = "" Then
            e.CancelEdit() = True
        End If
    End Sub

End Class
