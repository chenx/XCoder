''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This file is not used by the current project.
' After some work it can show almost the same content as an old window OpenFolderDialog.
' And it can dynamically add new folder to the FileSystem. But it lacks the context menu
' functions. 
' THe one provided by FolderNameEditor class is not flexible in sense that you
' 1) cannot select starting selected folder, 2) control of UI
' For 1) it's equally difficult in my own one, for 2) it's kind of offset by the
' other features (e.g., context menu) of the FolderNameEditor.
' So I decide to use the FolderNameEditor one.
'
' Note: This file also makes use of clsGetSpecialFolder.vb to get the path of "NetHood".
' @ This file is NOT used in this application @
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System.IO

Public Class FormOpenFolderDialog
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
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents ImageListFolder As System.Windows.Forms.ImageList
    Friend WithEvents TreeViewFolder As System.Windows.Forms.TreeView
    Friend WithEvents LabelCurPathTitle As System.Windows.Forms.Label
    Friend WithEvents LabelCurPath As System.Windows.Forms.Label
    Friend WithEvents btnNewDir As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormOpenFolderDialog))
        Me.TreeViewFolder = New System.Windows.Forms.TreeView()
        Me.ImageListFolder = New System.Windows.Forms.ImageList(Me.components)
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.LabelCurPathTitle = New System.Windows.Forms.Label()
        Me.LabelCurPath = New System.Windows.Forms.Label()
        Me.btnNewDir = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TreeViewFolder
        '
        Me.TreeViewFolder.ImageList = Me.ImageListFolder
        Me.TreeViewFolder.LabelEdit = True
        Me.TreeViewFolder.Location = New System.Drawing.Point(24, 96)
        Me.TreeViewFolder.Name = "TreeViewFolder"
        Me.TreeViewFolder.Size = New System.Drawing.Size(416, 264)
        Me.TreeViewFolder.TabIndex = 0
        '
        'ImageListFolder
        '
        Me.ImageListFolder.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageListFolder.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageListFolder.ImageStream = CType(resources.GetObject("ImageListFolder.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListFolder.TransparentColor = System.Drawing.Color.Transparent
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(248, 376)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(96, 32)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(352, 376)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(96, 32)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        '
        'LabelCurPathTitle
        '
        Me.LabelCurPathTitle.Location = New System.Drawing.Point(24, 16)
        Me.LabelCurPathTitle.Name = "LabelCurPathTitle"
        Me.LabelCurPathTitle.Size = New System.Drawing.Size(100, 16)
        Me.LabelCurPathTitle.TabIndex = 3
        Me.LabelCurPathTitle.Text = "Current Path:"
        Me.LabelCurPathTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelCurPath
        '
        Me.LabelCurPath.Location = New System.Drawing.Point(24, 40)
        Me.LabelCurPath.Name = "LabelCurPath"
        Me.LabelCurPath.Size = New System.Drawing.Size(424, 48)
        Me.LabelCurPath.TabIndex = 4
        Me.LabelCurPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnNewDir
        '
        Me.btnNewDir.Location = New System.Drawing.Point(24, 376)
        Me.btnNewDir.Name = "btnNewDir"
        Me.btnNewDir.Size = New System.Drawing.Size(104, 32)
        Me.btnNewDir.TabIndex = 5
        Me.btnNewDir.Text = "Make New Dir"
        '
        'FormOpenFolderDialog
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(466, 426)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnNewDir, Me.LabelCurPath, Me.LabelCurPathTitle, Me.btnCancel, Me.btnOK, Me.TreeViewFolder})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormOpenFolderDialog"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Browse For Folder"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Const drvUnknown As String = "Unknown"
    Const drvNotRootDirectory As String = "Not Root Directory"
    Const drvRemovableDisk As String = "Removable Disk" ' A Drive ?
    Const drvLocalDisk As String = "Local Disk"
    Const drvNetworkDrive As String = "Network Drive"
    Const drvCompactDisc As String = "Compact Disc"
    Const drvRamDisk As String = "RAM Disk"

    Const iconClosedFolder As Integer = 0
    Const iconOpenedFolder As Integer = 1
    Const iconLocalDisk As Integer = 2
    Const iconCompactDisk As Integer = 3
    Const iconNetDisk As Integer = 4
    Const iconMyComputer As Integer = 5
    Const iconMyNetworkPlaces As Integer = 6
    Const iconDesktop As Integer = 7
    Const iconMyDocuments As Integer = 8

    Const strDeskTop As String = "Desktop"
    Const strMyDocuments As String = "My Documents"
    Const strMyComputer As String = "My Computer"
    Const strMyNetworkPlaces As String = "My Network Places"

    Dim deskTopTreeNode As TreeNode
    Dim myDocumentsTreeNode As TreeNode
    Dim myComputerTreeNode As TreeNode
    Dim myNetworkPlacesTreeNode As TreeNode

    ' Special folders
    Dim DESKTOP_PATH As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
    Dim MY_DOCUMENTS_PATH As String = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
    Dim MY_NETWORK_PALCES_PATH As String = clsGetSpecialFolder.GetSpecialFolder(clsGetSpecialFolder.CSIDL.NETHOOD)

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Hide()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Hide()
    End Sub

    Private Sub FormOpenFolderDialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initTreeViewFolder()
    End Sub

    Private Sub initTreeViewFolder()
        Me.TreeViewFolder.Nodes.Clear()
        Me.TreeViewFolder.Sorted = False
        Me.adddeskTopTreeNode()
        Me.addmyDocumentsTreeNode()
        Me.addMyComputerNode()
        Me.addmyNetworkPlacesNode()
        Me.addDeskTopFolders()
        Me.deskTopTreeNode.Expand()
    End Sub

    Private Sub adddeskTopTreeNode()
        Me.deskTopTreeNode = New TreeNode(Me.strDeskTop)
        Me.deskTopTreeNode.ImageIndex = Me.iconDesktop
        Me.deskTopTreeNode.SelectedImageIndex = Me.iconDesktop
        Me.TreeViewFolder.Nodes.Add(Me.deskTopTreeNode)
    End Sub

    Private Sub addmyDocumentsTreeNode()
        Me.myDocumentsTreeNode = New TreeNode(Me.strMyDocuments)
        Me.myDocumentsTreeNode.ImageIndex = Me.iconMyDocuments
        Me.myDocumentsTreeNode.SelectedImageIndex = Me.iconMyDocuments
        Me.myDocumentsTreeNode.Nodes.Add("")
        Me.deskTopTreeNode.Nodes.Add(Me.myDocumentsTreeNode)
    End Sub

    Private Sub addMyComputerNode()
        myComputerTreeNode = New TreeNode(Me.strMyComputer)
        myComputerTreeNode.ImageIndex = Me.iconMyComputer
        myComputerTreeNode.SelectedImageIndex = Me.iconMyComputer
        Me.deskTopTreeNode.Nodes.Add(myComputerTreeNode)
        ShowDrives()
    End Sub

    Private Sub addmyNetworkPlacesNode()
        Me.myNetworkPlacesTreeNode = New TreeNode(Me.strMyNetworkPlaces)
        Me.myNetworkPlacesTreeNode.ImageIndex = Me.iconMyNetworkPlaces
        Me.myNetworkPlacesTreeNode.SelectedImageIndex = Me.iconMyNetworkPlaces
        Me.myNetworkPlacesTreeNode.Nodes.Add("")
        Me.deskTopTreeNode.Nodes.Add(Me.myNetworkPlacesTreeNode)
    End Sub

    Private Sub addDeskTopFolders()
        Me.EnumerateChildren(Me.deskTopTreeNode)
    End Sub

    Private Sub ShowDrives()
        Dim d() As String
        d = System.IO.Directory.GetLogicalDrives()

        Dim en As System.Collections.IEnumerator
        en = d.GetEnumerator
        While en.MoveNext
            getDriveFolderList(removeEndSlash(CStr(en.Current)))
        End While

        ' Get all logical hard drives    
        'Dim drives As String()
        'Dim counter As Int32
        'drives = Environment.GetLogicalDrives()
        'Console.WriteLine("======================")
        'Console.WriteLine("Available drives:")
        'For counter = 0 To drives.Length - 1
        'Console.WriteLine(drives(counter).ToString() & " : " & getDriveType(removeEndSlash(drives(counter).ToString())))
        'Next
        'Environment.SpecialFolder.DesktopDirectory()
        'Environment.personal.
    End Sub


    Private Function getDriveType(ByVal drive As String)
        Dim str As String
        Try
            'Dim sms As New System.Management.ManagementObject("Win32_logicaldisk='" & drive & "'")
            Dim sms As New System.Management.ManagementObject("Win32_logicaldisk='" & drive & "'")
            sms.Get()
            Select Case Convert.ToInt32(sms.Properties("DriveType").Value)
                Case 0 : str = Me.drvUnknown
                Case 1 : str = Me.drvNotRootDirectory
                Case 2 : str = Me.drvRemovableDisk
                Case 3 : str = Me.drvLocalDisk
                Case 4 : str = Me.drvNetworkDrive
                Case 5 : str = Me.drvCompactDisc
                Case 6 : str = Me.drvRamDisk
            End Select
        Catch ex As Exception
            MsgBox("getDriveType error: " & ex.ToString)
        End Try
        getDriveType = str
    End Function


    Private Function removeEndSlash(ByVal str As String)
        If str Is Nothing Then Exit Function
        If str.EndsWith("\") Or str.EndsWith("/") Then
            str = StrReverse(StrReverse(str).Substring(1))
        End If
        removeEndSlash = str
    End Function

    Private Sub getDriveFolderList(ByVal drive As String)
        Dim oNode As New System.Windows.Forms.TreeNode()
        Try
            Dim driveType As String = Me.getDriveType(drive)
            If driveType = Me.drvCompactDisc Then
                oNode.ImageIndex = Me.iconCompactDisk     ' Compact Disc 
                oNode.SelectedImageIndex = Me.iconCompactDisk
            ElseIf driveType = Me.drvNetworkDrive Then
                oNode.ImageIndex = Me.iconNetDisk
                oNode.SelectedImageIndex = Me.iconNetDisk
            Else
                oNode.ImageIndex = Me.iconLocalDisk
                oNode.SelectedImageIndex = Me.iconLocalDisk
            End If
            oNode.Text = drive
            Me.myComputerTreeNode.Nodes.Add(oNode)
            oNode.Nodes.Add("")
        Catch ex As Exception
            MsgBox("Cannot create initial node:" & ex.ToString)
            End
        End Try

    End Sub

    Private Sub TreeViewFolder_BeforeExpand(ByVal sender As Object, ByVal e As _
        System.Windows.Forms.TreeViewCancelEventArgs) Handles _
        TreeViewFolder.BeforeExpand

        'If e.Node.ImageIndex = 2 Then Exit Sub

        Try
            If e.Node.GetNodeCount(False) = 1 And e.Node.Nodes(0).Text = "" Then
                e.Node.Nodes(0).Remove()
                EnumerateChildren(e.Node)
            End If
        Catch ex As Exception
            MsgBox("Unable to expand " & e.Node.FullPath & ":" & ex.ToString)
        End Try

        If e.Node.GetNodeCount(False) > 0 Then
            If e.Node Is Me.deskTopTreeNode Then Exit Sub
            If e.Node Is Me.myDocumentsTreeNode Then Exit Sub
            If e.Node Is Me.myComputerTreeNode Then Exit Sub
            If e.Node Is Me.myNetworkPlacesTreeNode Then Exit Sub
            If InStr(e.Node.Text, ":") Then Exit Sub

            e.Node.ImageIndex = Me.iconOpenedFolder
            e.Node.SelectedImageIndex = Me.iconOpenedFolder
        End If

    End Sub

    Private Sub EnumerateChildren(ByVal oParent As System.Windows.Forms.TreeNode)
        'MsgBox(oParent.FullPath)
        Dim path As String = getActualPath(oParent.FullPath)

        'Dim oFS As New DirectoryInfo(oParent.FullPath & "\")
        Dim oFS As New DirectoryInfo(path & "\")
        Dim oDir As DirectoryInfo
        'Dim oFile As FileInfo

        Try
            For Each oDir In oFS.GetDirectories()
                ' Don't show system or hidden folders - as option
                If Not (CDbl(oDir.Attributes() And FileAttributes.System) = FileAttributes.System) Then 'And _
                    '        CDbl(oDir.Attributes() And FileAttributes.Hidden) = FileAttributes.Hidden) Then
                    Dim oNode As New System.Windows.Forms.TreeNode()
                    oNode.Text = oDir.Name
                    oNode.ImageIndex = Me.iconClosedFolder
                    oNode.SelectedImageIndex = Me.iconClosedFolder
                    oParent.Nodes.Add(oNode)
                    If oDir.GetDirectories.GetLength(0) > 0 Then
                        oNode.Nodes.Add("")
                    End If
                End If
            Next
        Catch ex As Exception
            'MsgBox("Cannot list folders of " & oParent.FullPath) '& ":" & ex.ToString)
        End Try

        ' For files
        '        Try
        '        For Each oFile In oFS.GetFiles()
        '            Dim oNode As New System.Windows.Forms.TreeNode()
        '            oNode.Text = oFile.Name & " (" & oFile.Length & " bytes)"
        '            oNode.ImageIndex = 2
        '            oNode.SelectedImageIndex = 2
        '            oParent.Nodes.Add(oNode)
        '        Next
        '        Catch ex As Exception
        '            MsgBox("Cannot list files in " & oParent.FullPath & ":" & ex.ToString)
        '        End Try
    End Sub

    Private Sub TreeViewFolder_BeforeCollapse(ByVal sender As Object, ByVal e As _
        System.Windows.Forms.TreeViewCancelEventArgs) Handles _
        TreeViewFolder.BeforeCollapse

        If e.Node Is Me.deskTopTreeNode Then Exit Sub
        If e.Node Is Me.myDocumentsTreeNode Then Exit Sub
        If e.Node Is Me.myComputerTreeNode Then Exit Sub
        If e.Node Is Me.myNetworkPlacesTreeNode Then Exit Sub
        If InStr(e.Node.Text, ":") Then Exit Sub

        e.Node.ImageIndex = Me.iconClosedFolder
        e.Node.SelectedImageIndex = Me.iconClosedFolder

    End Sub


    Private Sub TreeViewFolder_AfterSelect(ByVal sender As System.Object, _
            ByVal e As System.Windows.Forms.TreeViewEventArgs) _
            Handles TreeViewFolder.AfterSelect
        Me.TreeViewFolder.SelectedNode.Expand()

        Dim curPath As String = getActualPath(Me.TreeViewFolder.SelectedNode.FullPath)

        Me.LabelCurPath.Text = curPath

        If curPath = "" Then
            Me.btnOK.Enabled = False
        Else
            Me.btnOK.Enabled = True
        End If
    End Sub

    Private Function getActualPath(ByVal curPath As String)

        If curPath = Me.strDeskTop Then
            curPath = Me.DESKTOP_PATH
        Else
            curPath = curPath.Substring(Len(Me.strDeskTop) + 1)

            If curPath.StartsWith(strMyComputer) Then
                If curPath = strMyComputer Then
                    curPath = ""
                Else    ' Remove "My Computer\"
                    curPath = curPath.Substring(Len(Me.strMyComputer) + 1)
                End If
            ElseIf curPath.StartsWith(Me.strMyDocuments) Then
                If curPath = strMyDocuments Then
                    curPath = Me.MY_DOCUMENTS_PATH
                Else
                    curPath = Me.MY_DOCUMENTS_PATH & "\" & curPath.Substring(Len(Me.strMyDocuments) + 1)
                End If
            ElseIf curPath.StartsWith(Me.strMyNetworkPlaces) Then
                If curPath = strMyNetworkPlaces Then
                    curPath = Me.MY_NETWORK_PALCES_PATH
                Else
                    curPath = Me.MY_NETWORK_PALCES_PATH & "\" & curPath.Substring(Len(Me.strMyNetworkPlaces) + 1)
                End If
            Else ' Desktop folders
                curPath = Me.DESKTOP_PATH & "\" & curPath
            End If
        End If

        Return curPath
    End Function


    Private Sub btnNewDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewDir.Click
        Dim curDir As String = Me.LabelCurPath.Text
        Dim curTreeNode As TreeNode = Me.TreeViewFolder.SelectedNode

        If curDir <> getActualPath(Me.TreeViewFolder.SelectedNode.FullPath) Then
            MsgBox("error - btnNewDir_Click")
            Exit Sub
        End If

        If curDir <> "" Then
            Dim newFolderName As String
            newFolderName = Trim(InputBox("Name", "New Folder"))

            If newFolderName = "" Then
                MsgBox("Name is empty, this folder will not be created")
            Else
                Dim newDir As String = curDir & "\" & newFolderName
                If IOManager.folderExists(newDir) Then
                    MsgBox("This folder already exists, please use another name")
                Else
                    IOManager.CreateFolder(newDir)
                    Me.TreeViewFolder.SelectedNode.Nodes.Add(New TreeNode(newFolderName))
                    Me.TreeViewFolder.SelectedNode.TreeView.Sorted = True
                    Me.TreeViewFolder.SelectedNode.TreeView.Refresh()
                End If
            End If

        End If
    End Sub
End Class
