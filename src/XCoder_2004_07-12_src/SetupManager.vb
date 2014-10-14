Public Class SetupManager

    Friend Shared SoftwareName As String = "XinCoder"
    Friend Shared version As String = SoftwareName & " 1.0"

    ' Path information
    Friend Shared ROOT_PATH As String = AppPath(True)
    Friend Shared OUTPUT_FOLDER As String = "xcOutput\"
    Friend Shared OUTPUT_PATH As String = ROOT_PATH & OUTPUT_FOLDER
    Friend Shared SOUND_PATH As String = ROOT_PATH & "Sound\"
    Friend Shared IMAGE_PATH As String = ROOT_PATH & "Image\"
    Friend Shared BACKUP_PATH As String = ROOT_PATH & "Bck\"
    Friend Shared DOC_PATH As String = ROOT_PATH & "doc\"
    Friend Shared CONFIG_FILE As String = "config.txt"
    Friend Shared CONFIG_FILE_PATH As String = ROOT_PATH & CONFIG_FILE
    Friend Shared LOG_FILE As String = "log.txt"
    Friend Shared HELP_FILE As String = DOC_PATH & "Help.chm"

    Friend Shared PROJ_FILE_EXTENSION As String = ".xcproj"
    Friend Shared MODULE_FILE_EXTENSION As String = ".xc"

    Friend Shared RECENT_PROJECTS_FILE_PATH As String = ROOT_PATH & "rpf.txt"
    Friend Shared MaxRecentProjectsCount As Integer = 4

    Friend Shared USE_SOUND As Boolean = True

    Friend Shared LIB_PATH As String = ROOT_PATH & "lib\"
    Friend Shared WEBSITE_LIB_PATH As String = LIB_PATH & "website\"
    Friend Shared IMS_LIB_PATH As String = LIB_PATH & "ims\"

    ' Get application directory
    Public Shared Function AppPath(Optional ByVal WithSlash As Boolean = False) As String
        'MsgBox("Application.StartupPath = " & Application.StartupPath)
        'MsgBox("Environment.CurrentDirectory = " & Environment.CurrentDirectory)
        'MsgBox("IO.Path.GetDirectoryName(Application.ExecutablePath) = " & IO.Path.GetDirectoryName(Application.ExecutablePath))
        'MsgBox("Application.ExecutablePath = " & Application.ExecutablePath())
        Dim _startupPath As String = Application.StartupPath
        _startupPath = getUpperDir(_startupPath)

        Dim Result As String = _startupPath

        If WithSlash Then
            If Result.EndsWith("\") = False Then
                Result &= "\"
            End If
        Else
            If Result.EndsWith("\") Then
                Result = Strings.Left(Result, Len(Result) - 1)
            End If
        End If

        Return Result
    End Function


    ' Given a file or folder path, get the parent folder path
    Private Shared Function getUpperDir(ByVal dir As String) As String
        dir = StrReverse(dir)
        If dir.StartsWith("\") Then
            dir = Mid(dir, 2)
        End If
        dir = Mid(dir, InStr(dir, "\"))
        dir = StrReverse(dir)
        Return dir
    End Function


    ' Read configuation from file config.txt
    Public Shared Function readConfig() As Boolean
        Dim errorInfo As String

        Dim configs() As String
        IOManager.getFileAsLines(CONFIG_FILE_PATH, configs, errorInfo)

        If errorInfo <> "" Then
            MsgBox(SoftwareName & " cannot start: " & vbCrLf & errorInfo, MsgBoxStyle.Critical)
            Return False
        End If

        Dim i As Integer
        Dim maxI As Integer = configs.Length - 1

        For i = 0 To maxI
            If LCase(configs(i)).StartsWith("recent_projects_count") Then
                'recent_projects_count = CType(MyUtil.getStrAfterDelimiter(configs(i), "="), Integer)
            End If
        Next

        Return True
    End Function

End Class
