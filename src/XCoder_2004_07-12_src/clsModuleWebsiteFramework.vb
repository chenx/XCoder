Imports XinCoder.xcUtil

Public Class clsModuleWebsiteFramework
    ' This class contains functions to generate website framework files.

    Private indexFileContent As String
    Private cssFileContent As String
    Private rootIncHeaderFileContent As String
    Private rootIncFooterFileContent As String
    'Private xcLoginForm As String
    'Private xcLogoutForm As String

    Private ODBCName As String
    Private ODBCLogin As String
    Private ODBCPassword As String

    Private indexFile As String = "index.asp"
    Private cssFile As String = "xc.css"
    Private loginFolder As String = "login"
    Private rootIncFolder As String = "rootInclude"
    Private rootIncHeaderFile As String = "header.asp"
    Private rootIncFooterFile As String = "footer.asp"
    Private loginAdmSecFile As String = "admin_security.asp"
    Private loginMemSecFile As String = "mem_security.asp"
    Private loginLoginFile As String = "login.asp"
    Private loginLogoutFile As String = "logout.asp"
    Private memFolder As String = "member"
    Private memHomeFile As String = "index.asp"
    Private aspBinFolder As String = "asp-bin"
    Private aspBinAspFuncLibFile As String = "aspFuncLib.asp"
    Private aspBinDBConfigFile As String = "database_config.asp"
    Private aspBinAdovbsFile As String = "adovbs.inc"
    Private admFolder As String = "admin"
    Private admHomeFile As String = "index.asp"
    Private admApproveFile As String
    Private imageFolder As String = "images"

    Private regFilePath As String
    Private memRegConfirmFilePath As String
    Private admRegListFilePath As String

    Public Function init(ByVal indexFileName As String, ByVal cssFileName As String, _
        ByVal loginFolderName As String, ByVal rootIncFolderName As String, _
        ByVal rootIncHeaderFileName As String, ByVal rootIncFooterFileName As String, _
        ByVal loginAdmSecFileName As String, ByVal loginMemSecFileName As String, _
        ByVal loginLoginFileName As String, ByVal loginLogoutFileName As String, _
        ByVal memFolderName As String, ByVal memHomeFileName As String, _
        ByVal aspBinFolderName As String, ByVal aspBinAspFuncLibFileName As String, _
        ByVal aspBinDBConfigFileName As String, ByVal aspBinAdovbsFileName As String, _
        ByVal admFolderName As String, ByVal admHomeFileName As String, _
        ByVal admApproveFileName As String, ByVal imageFolderName As String, _
        ByVal odbcName As String, ByVal odbcLogin As String, ByVal odbcPassword As String, _
        ByVal indexFileContent As String, ByVal cssFileContent As String, _
        ByVal rootIncHeaderFileContent As String, ByVal rootIncFooterFileContent As String, _
        ByVal registerFilePath As String, ByVal memRegisterConfirmFilePath As String, _
        ByVal admRegisterListFilePath As String)

        Me.indexFileContent = indexFileContent
        Me.cssFileContent = cssFileContent
        Me.rootIncHeaderFileContent = rootIncHeaderFileContent
        Me.rootIncFooterFileContent = rootIncFooterFileContent

        Me.ODBCName = odbcName
        Me.ODBCLogin = odbcLogin
        Me.ODBCPassword = odbcPassword

        Me.indexFile = indexFileName
        Me.cssFile = cssFileName
        Me.loginFolder = loginFolderName
        Me.loginLoginFile = loginLoginFileName
        Me.loginLogoutFile = loginLogoutFileName
        Me.loginAdmSecFile = loginAdmSecFileName
        Me.loginMemSecFile = loginMemSecFileName
        Me.rootIncFolder = rootIncFolderName
        Me.rootIncHeaderFile = rootIncHeaderFileName
        Me.rootIncFooterFile = rootIncFooterFileName
        Me.memFolder = memFolderName
        Me.memHomeFile = memHomeFileName
        Me.aspBinFolder = aspBinFolderName
        Me.aspBinAspFuncLibFile = aspBinAspFuncLibFileName
        Me.aspBinDBConfigFile = aspBinDBConfigFileName
        Me.aspBinAdovbsFile = aspBinAdovbsFile
        Me.admFolder = admFolderName
        Me.admHomeFile = admHomeFileName
        Me.admApproveFile = admApproveFileName
        Me.imageFolder = imageFolderName

        Me.regFilePath = registerFilePath
        Me.memRegConfirmFilePath = memRegisterConfirmFilePath
        Me.admRegListFilePath = admRegisterListFilePath
    End Function


    'Private adminHomeFileNode As New TreeNode("index.asp", 0, 0)
    Public Function getAdminHomeFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & "../" & loginFolder & "/" & loginAdmSecFile & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & "../" & rootIncFolder & "/" & rootIncHeaderFile & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "	' used by admin pages to decide this is admin page." & vbCrLf
        str = str & "	' Only session(" & Chr(34) & "_isAdmin" & Chr(34) & ") is not always enough," & vbCrLf
        str = str & "	' e.g., a page is used by both member and admin," & vbCrLf
        str = str & "	' then one can't tell by session(" & Chr(34) & "_isAdmin" & Chr(34) & ") whether the user is in" & vbCrLf
        str = str & "	' admin mode or member mode." & vbCrLf
        str = str & "	session(" & Chr(34) & "_adminMode" & Chr(34) & ") = true" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<h2>Admin Home Page</h2>" & vbCrLf
        str = str & "<table width=" & Chr(34) & "90%" & Chr(34) & " border=0><tr><td>" & vbCrLf
        str = str & vbCrLf
        str = str & "<p><A href=" & Chr(34) & "../" & memFolder & "/" & memHomeFile & Chr(34) & ">Back to Member Home</a> </p>" & vbCrLf
        str = str & vbCrLf
        str = str & "      <UL>" & vbCrLf
        str = str & "      " & xcUtil.admTaskListStart & vbCrLf
        str = str & "      <LI><A HREF=" & Chr(34) & Me.admApproveFile & Chr(34) & ">Approve/Disapprove Administrators</A>" & vbCrLf
        str = str & "      <LI><A HREF=" & Chr(34) & Me.admRegListFilePath & Chr(34) & ">Member registration management</A>" & vbCrLf
        str = str & "      " & xcUtil.admTaskListEnd & vbCrLf
        str = str & "      </UL>" & vbCrLf
        str = str & vbCrLf
        str = str & "</td></tr></table>" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & "../" & rootIncFolder & "/" & rootIncFooterFile & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        Return str
    End Function

    'Private adminApproveFileNode As New TreeNode("approveAdmin.asp", 0, 0)
    Public Function getAdminApproveFile() As String
        Dim str As String = ""
        Return str
    End Function

    'Private aspBinAdovbsFileNode As New TreeNode("adovbs.inc", 0, 0)
    Public Function getAspBinAdovbsFile() As String
        Dim str As String = ""
        str = IOManager.GetFileContents(SetupManager.WEBSITE_LIB_PATH & "files/asp-bin/adovbs.inc")
        Return str
    End Function

    'Private aspBinDBConfigFileNode As New TreeNode("db_config.asp", 0, 0)
    Public Function getAspBinDBConfigFile() As String
        Dim str As String = ""
        str += "<%" & vbCrLf
        str += "''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str += "' SET LOGIN INFORMATION " & vbCrLf
        str += "''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str += "Dim db_db, db_login, db_password" & vbCrLf
        str += vbCrLf
        str += "db_db = " & Chr(34) & Me.ODBCName & Chr(34) & vbCrLf
        str += "db_login = " & Chr(34) & Me.ODBCLogin & Chr(34) & vbCrLf
        str += "db_password = " & Chr(34) & Me.ODBCPassword & Chr(34) & vbCrLf
        str += "%>"
        Return str
    End Function

    'Private aspBinAspFuncLibFileNode As New TreeNode("aspFuncLib.asp", 0, 0)
    Public Function getAspBinAspFuncLibFile() As String
        Dim str As String = ""
        str += "<!--#include file=" & Chr(34) & Me.aspBinAdovbsFile & Chr(34) & "-->" & vbCrLf
        str += "<!--#include file=" & Chr(34) & Me.aspBinDBConfigFile & Chr(34) & "-->" & vbCrLf
        str += vbCrLf
        str += IOManager.GetFileContents(SetupManager.WEBSITE_LIB_PATH & "files/asp-bin/xcDatabase.asp")
        Return str
    End Function

    'Private loginLoginFileNode As New TreeNode("login.asp", 0, 0)
    ' Use default login - not using database, but hardcoded login account
    Public Function getLoginLoginFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & "../" & Me.aspBinFolder & "/" & Me.aspBinAspFuncLibFile & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "Dim login, password, validateResult" & vbCrLf
        str = str & vbCrLf
        str = str & "login = trim(request(" & Chr(34) & "usrLoginname" & Chr(34) & "))" & vbCrLf
        str = str & "password = trim(request(" & Chr(34) & "usrPassword" & Chr(34) & "))" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & "' Default admin account. To be removed after formal admin account is setup" & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & vbCrLf
        str = str & "if isDefaultAdmin(login, password) then" & vbCrLf
        str = str & "    session(" & Chr(34) & "_isMem" & Chr(34) & ") = " & Chr(34) & "true" & Chr(34) & vbCrLf
        str = str & "    session(" & Chr(34) & "_isAdmin" & Chr(34) & ") = " & Chr(34) & "true" & Chr(34) & vbCrLf
        str = str & "    Response.redirect " & Chr(34) & "../" & Me.memFolder & "/" & Me.memHomeFile & "?isDefaultAdmin=true" & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & vbCrLf
        str = str & "function isDefaultAdmin(login, password)" & vbCrLf
        str = str & "    if login = " & Chr(34) & "admin" & Chr(34) & " and password = " & Chr(34) & "pass" & Chr(34) & " then" & vbCrLf
        str = str & "        isDefaultAdmin = true" & vbCrLf
        str = str & "    else" & vbCrLf
        str = str & "        isDefaultAdmin = false" & vbCrLf
        str = str & "    end if" & vbCrLf
        str = str & "end function" & vbCrLf
        str = str & vbCrLf
        str = str & "Response.Redirect " & Chr(34) & "../index.asp?login='" & Chr(34) & " & login & " & Chr(34) & "'&error=Invalid login information." & Chr(34) & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        Return str
    End Function

    'Private loginLogoutFileNode As New TreeNode("logout.asp", 0, 0)
    Public Function getLoginLogoutFile() As String
        Dim str As String = ""
        str = str & "<%" & vbCrLf
        str = str & "session.Abandon ()" & vbCrLf
        str = str & "Response.Redirect " & Chr(34) & "../" & Me.indexFile & Chr(34) & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        Return str
    End Function

    'Private loginAdminSecFileNode As New TreeNode("admin_Security.asp", 0, 0)
    Public Function getLoginAdminSecFile() As String
        Dim str As String = ""
        str = str & "<%" & vbCrLf
        str = str & "' Server side code to prevent browser cache page content." & vbCrLf
        str = str & "Response.Buffer = True" & vbCrLf
        str = str & "Response.ExpiresAbsolute = Now() - 1" & vbCrLf
        str = str & "Response.Expires = 0" & vbCrLf
        str = str & "Response.CacheControl = " & Chr(34) & "no-cache" & Chr(34) & vbCrLf
        str = str & vbCrLf
        str = str & "if session(" & Chr(34) & "_isAdmin" & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	Response.Redirect " & Chr(34) & "../" & Me.indexFile & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        Return str
    End Function

    'Private loginMemSecFileNode As New TreeNode("mem_security.asp", 0, 0)
    Public Function getLoginMemSecFile() As String
        Dim str As String = ""
        str = str & "<%" & vbCrLf
        str = str & "' Server side code to prevent browser cache page content." & vbCrLf
        str = str & "Response.Buffer = True" & vbCrLf
        str = str & "Response.ExpiresAbsolute = Now() - 1" & vbCrLf
        str = str & "Response.Expires = 0" & vbCrLf
        str = str & "Response.CacheControl = " & Chr(34) & "no-cache" & Chr(34) & vbCrLf
        str = str & vbCrLf
        str = str & "if session(" & Chr(34) & "_isMem" & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	Response.Redirect " & Chr(34) & "../" & Me.indexFile & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function

    'Private memberHomeFileNode As New TreeNode("index.asp", 0, 0)
    Public Function getMemHomeFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & "../" & Me.loginFolder & "/" & Me.loginMemSecFile & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & "../" & Me.rootIncFolder & "/" & Me.rootIncHeaderFile & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "' This finishes an admin working session if any" & vbCrLf
        str = str & "session(" & Chr(34) & "_adminTasks" & Chr(34) & ") = false" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<h2>Member Home Page</h2>" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<table border=0 width=90%><tr><td>" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<%if session(" & Chr(34) & "_isAdmin" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
        str = str & "<h2>Site Administration</h2>" & vbCrLf
        str = str & "<ul>" & vbCrLf
        str = str & "<li><a href=" & Chr(34) & "../" & Me.admFolder & "/" & Me.admHomeFile & Chr(34) & ">Administrator's home</a>" & vbCrLf
        str = str & "</ul>" & vbCrLf
        str = str & "<%end if%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<h2>Member functions</h2>" & vbCrLf
        str = str & "<ul>" & vbCrLf
        str = str & xcUtil.memTaskListStart & vbCrLf
        str = str & "<li><a href=" & Chr(34) & Me.memRegConfirmFilePath & Chr(34) & ">My Profile</a>" & vbCrLf
        str = str & xcUtil.memTaskListEnd & vbCrLf
        str = str & "</ul>" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "</td></tr></table>" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & "../" & Me.rootIncFolder & "/" & Me.rootIncFooterFile & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        Return str
    End Function


    Public Function getCssFile() As String
        Return Me.cssFileContent
    End Function


    Public Function getIndexFile() As String
        Return Me.indexFileContent
    End Function


    Public Function getRootIncHeaderFile() As String
        Return Me.rootIncHeaderFileContent
    End Function


    Public Function getRootIncFooterFile() As String
        Return Me.rootIncFooterFileContent
    End Function

End Class
