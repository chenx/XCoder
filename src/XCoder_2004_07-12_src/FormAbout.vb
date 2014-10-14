Public Class FormAbout
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
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnSysInfo As System.Windows.Forms.Button
    Friend WithEvents PictureBoxAbout As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblSoftware As System.Windows.Forms.Label
    Friend WithEvents lblVer As System.Windows.Forms.Label
    Friend WithEvents lblCopyright As System.Windows.Forms.Label
    Friend WithEvents lblAuthor As System.Windows.Forms.Label
    Friend WithEvents lblLicensee As System.Windows.Forms.Label
    Friend WithEvents lblWarning As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormAbout))
        Me.btnOk = New System.Windows.Forms.Button()
        Me.PictureBoxAbout = New System.Windows.Forms.PictureBox()
        Me.lblSoftware = New System.Windows.Forms.Label()
        Me.lblVer = New System.Windows.Forms.Label()
        Me.lblCopyright = New System.Windows.Forms.Label()
        Me.btnSysInfo = New System.Windows.Forms.Button()
        Me.lblAuthor = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblLicensee = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblWarning = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(272, 336)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(112, 32)
        Me.btnOk.TabIndex = 0
        Me.btnOk.Text = "Ok"
        '
        'PictureBoxAbout
        '
        Me.PictureBoxAbout.Image = CType(resources.GetObject("PictureBoxAbout.Image"), System.Drawing.Bitmap)
        Me.PictureBoxAbout.Location = New System.Drawing.Point(32, 32)
        Me.PictureBoxAbout.Name = "PictureBoxAbout"
        Me.PictureBoxAbout.Size = New System.Drawing.Size(56, 296)
        Me.PictureBoxAbout.TabIndex = 2
        Me.PictureBoxAbout.TabStop = False
        '
        'lblSoftware
        '
        Me.lblSoftware.Font = New System.Drawing.Font("Algerian", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSoftware.Location = New System.Drawing.Point(120, 32)
        Me.lblSoftware.Name = "lblSoftware"
        Me.lblSoftware.Size = New System.Drawing.Size(184, 40)
        Me.lblSoftware.TabIndex = 3
        Me.lblSoftware.Text = "XinCoder"
        '
        'lblVer
        '
        Me.lblVer.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.lblVer.Location = New System.Drawing.Point(320, 34)
        Me.lblVer.Name = "lblVer"
        Me.lblVer.Size = New System.Drawing.Size(136, 32)
        Me.lblVer.TabIndex = 4
        Me.lblVer.Text = "Version 1.0"
        '
        'lblCopyright
        '
        Me.lblCopyright.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCopyright.Location = New System.Drawing.Point(120, 88)
        Me.lblCopyright.Name = "lblCopyright"
        Me.lblCopyright.Size = New System.Drawing.Size(296, 16)
        Me.lblCopyright.TabIndex = 5
        Me.lblCopyright.Text = "Copyright @ 2004. All rights reserved"
        '
        'btnSysInfo
        '
        Me.btnSysInfo.Location = New System.Drawing.Point(392, 336)
        Me.btnSysInfo.Name = "btnSysInfo"
        Me.btnSysInfo.Size = New System.Drawing.Size(112, 32)
        Me.btnSysInfo.TabIndex = 1
        Me.btnSysInfo.Text = "System Info ..."
        '
        'lblAuthor
        '
        Me.lblAuthor.Location = New System.Drawing.Point(120, 112)
        Me.lblAuthor.Name = "lblAuthor"
        Me.lblAuthor.Size = New System.Drawing.Size(192, 16)
        Me.lblAuthor.TabIndex = 8
        Me.lblAuthor.Text = "Author: Xin Chen"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblLicensee})
        Me.GroupBox1.Location = New System.Drawing.Point(120, 160)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(384, 48)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        '
        'lblLicensee
        '
        Me.lblLicensee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblLicensee.Location = New System.Drawing.Point(3, 18)
        Me.lblLicensee.Name = "lblLicensee"
        Me.lblLicensee.Size = New System.Drawing.Size(378, 27)
        Me.lblLicensee.TabIndex = 0
        Me.lblLicensee.Text = "Thomas Xin Chen"
        Me.lblLicensee.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(120, 136)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(280, 23)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "This software is licensed to:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblWarning
        '
        Me.lblWarning.Location = New System.Drawing.Point(120, 224)
        Me.lblWarning.Name = "lblWarning"
        Me.lblWarning.Size = New System.Drawing.Size(384, 104)
        Me.lblWarning.TabIndex = 12
        Me.lblWarning.Text = "Warning: This computer program is protected by copyright law and international tr" & _
        "eaties. Unauthorized reproduction or distribution of this program, or any portio" & _
        "n of it, may result in severe civil and criminal penalties, and will be prosecut" & _
        "ed to the maximum extent possible under law."
        '
        'FormAbout
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(544, 392)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblWarning, Me.Label6, Me.GroupBox1, Me.lblAuthor, Me.btnSysInfo, Me.lblCopyright, Me.lblVer, Me.lblSoftware, Me.PictureBoxAbout, Me.btnOk})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormAbout"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "About"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myFrmSysInfo As FormSysInfo

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Me.Hide()
    End Sub

    Private Sub btnSysInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSysInfo.Click
        If Me.myFrmSysInfo Is Nothing Then
            Me.myFrmSysInfo = New FormSysInfo()
        End If
        Me.myFrmSysInfo.Show()
    End Sub

    Private Sub FormAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        setLicensee()
    End Sub

    Private Sub setLicensee()
        Me.lblLicensee.Text = "Thomas Xin Chen"
    End Sub
End Class
