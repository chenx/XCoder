' This works for VB Only !!!

Public Class clsSpecialFolder
    Public Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" _
   (ByVal hWnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean

    Private Const CSF_DESKTOP = &H0
    Private Const CSF_INTERNET = &H1
    Private Const CSF_PROGRAMS = &H2
    Private Const CSF_CONTROLS = &H3
    Private Const CSF_PRINTERS = &H4
    Private Const CSF_PERSONAL = &H5
    Private Const CSF_FAVORITES = &H6
    Private Const CSF_STARTUP = &H7
    Private Const CSF_RECENT = &H8
    Private Const CSF_SENDTO = &H9
    Private Const CSF_BITBUCKET = &HA
    Private Const CSF_STARTMENU = &HB
    Private Const CSF_DESKTOPDIRECTORY = &H10
    Private Const CSF_DRIVES = &H11
    Private Const CSF_NETWORK = &H12
    Private Const CSF_NETHOOD = &H13
    Private Const CSF_FONTS = &H14
    Private Const CSF_TEMPLATES = &H15
    Private Const CSF_COMMON_STARTMENU = &H16
    Private Const CSF_COMMON_PROGRAMS = &H17
    Private Const CSF_COMMON_STARTUP = &H18
    Private Const CSF_COMMON_DESKTOPDIRECTORY = &H19
    Private Const CSF_APPDATA = &H1A
    Private Const CSF_PRINTHOOD = &H1B
    Private Const CSF_ALTSTARTUP = &H1D
    Private Const CSF_COMMON_ALTSTARTUP = &H1E
    Private Const CSF_COMMON_FAVORITES = &H1F
    Private Const CSF_INTERNET_CACHE = &H20
    Private Const CSF_COOKIES = &H21
    Private Const CSF_HISTORY = &H22

    Public Enum SpecialFolder
        Desktop
        Internet
        MyDocuments
        Programs
        Cookies
        '  Add whichever ones you want to commonly reference here
    End Enum

    Private Const iMaxPathLength As Integer = 512

    Public Function Get_Path(ByVal WhichPath As SpecialFolder) As String
        Dim lReturn As Long
        Dim sBuf As String
        Dim lMode As Long

        sBuf = Space$(iMaxPathLength)  '-- Need space filled buffer for API calls

        Select Case WhichPath
            Case SpecialFolder.Desktop
                lMode = CSF_DESKTOP
            Case SpecialFolder.Internet
                lMode = CSF_INTERNET
            Case SpecialFolder.MyDocuments
                lMode = CSF_PERSONAL
            Case SpecialFolder.Programs
                lMode = CSF_PROGRAMS
            Case SpecialFolder.Cookies
                lMode = CSF_COOKIES
            Case Else
                lMode = CSF_DESKTOPDIRECTORY
        End Select

        lReturn = SHGetSpecialFolderPath(0, sBuf, lMode, False)

        ' The string that the API call returns is terminated by a Chr(0)
        ' Trim everything from the (0) on

        lReturn = InStr(sBuf, Chr$(0))
        If lReturn > 1 Then
            Get_Path = Left$(sBuf, lReturn - 1)
        Else
            Get_Path = ""
        End If

    End Function
End Class
