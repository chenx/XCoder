Public Class FormSysInfo
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
    Friend WithEvents tbSysInfo As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.tbSysInfo = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(184, 248)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(112, 32)
        Me.btnOk.TabIndex = 1
        Me.btnOk.Text = "OK"
        '
        'tbSysInfo
        '
        Me.tbSysInfo.Location = New System.Drawing.Point(24, 24)
        Me.tbSysInfo.Multiline = True
        Me.tbSysInfo.Name = "tbSysInfo"
        Me.tbSysInfo.ReadOnly = True
        Me.tbSysInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.tbSysInfo.Size = New System.Drawing.Size(416, 208)
        Me.tbSysInfo.TabIndex = 2
        Me.tbSysInfo.Text = ""
        '
        'FormSysInfo
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(462, 300)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.tbSysInfo, Me.btnOk})
        Me.Name = "FormSysInfo"
        Me.Text = "System Information"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmSysInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim objWMI As New clsWMI()
        Dim sysInfoStr As String = ""
        With objWMI
            sysInfoStr = sysInfoStr & ("Computer Name = " & .ComputerName) & vbCrLf
            sysInfoStr = sysInfoStr & ("Computer Manufacturer = " & .Manufacturer) & vbCrLf
            sysInfoStr = sysInfoStr & ("Computer Model = " & .Model) & vbCrLf
            sysInfoStr = sysInfoStr & ("OS Name = " & .OsName) & vbCrLf
            sysInfoStr = sysInfoStr & ("OS Version = " & .OSVersion) & vbCrLf
            sysInfoStr = sysInfoStr & ("System Type = " & .SystemType) & vbCrLf
            sysInfoStr = sysInfoStr & ("Total Physical Memory = " & .TotalPhysicalMemory) & vbCrLf
            sysInfoStr = sysInfoStr & ("Windows Directory = " & .WindowsDirectory)
        End With
        Me.tbSysInfo.Text = sysInfoStr
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Me.Hide()
    End Sub
End Class
