Imports XinCoder.xcUtil

Public Class clsModuleLogin
    ' This class creates information needed for a Member Register module

    Private clsMD As New clsMetadata()
    Private websiteHomeFilePath As String
    Private websiteFuncFilePath As String = "../asp-bin/aspFuncLib.asp"
    Private moduleFuncFilePath As String
    Private moduleEditTableFilePath As String
    Private moduleConfirmTableFilePath As String
    Private loginAdmSecFilePath As String = "../login/admin_security.asp"
    Private loginMemSecFilePath As String = "../login/mem_security.asp"
    Private rootIncHeaderFilePath As String = "../rootInclude/header.asp"
    Private rootIncFooterFilePath As String = "../rootInclude/footer.asp"
    Private admHomeFilePath As String
    Private memHomeFilePath As String

    Private regFile As String
    Private regConfirmFile As String
    Private memRegEditFile As String
    Private memRegConfirmFile As String
    Private memRegChgPasswordFile As String
    Private admRegListFile As String
    Private admRegEditFile As String
    Private admRegConfirmFile As String
    Private admRegChgPasswordFile As String
    'Private DBTableFileName As String
    Private DBTableName As String

    Private loginFld As String
    Private passwordFld As String
    Private passwordRptFld As String
    Private passwordRptFldHTMLName As String
    Private loginIsCaseSensitive As Boolean
    Private passwordIsCaseSensitive As Boolean
    Private passwordMinLength As Integer
    Private loginNeedsAdminApproval As Boolean


    Public Sub init(ByVal metadata As String, _
        ByVal globalHomeFilePath As String, ByVal globalFuncFilePath As String, _
        ByVal loginAdmSecurityFilePath As String, ByVal loginMemSecurityFilePath As String, _
        ByVal headerFilePath As String, ByVal footerFilePath As String, _
        ByVal adminHomeFilePath As String, ByVal memberHomeFilePath As String, _
        ByVal moduleFuncFilePath As String, ByVal moduleEditTableFilePath As String, _
        ByVal moduleConfirmTableFilePath As String, _
        ByVal regEditFile As String, ByVal regConfirmFile As String, _
        ByVal memRegEditFile As String, ByVal memRegConfirmFile As String, _
        ByVal memRegChangePasswordFile As String, _
        ByVal admRegChangePasswordFile As String, ByVal admRegListFile As String, _
        ByVal admRegEditFile As String, ByVal admRegConfirmFile As String, _
        ByVal DBTableName As String, ByVal loginDBField As String, _
        ByVal passwordDBField As String, ByVal passwordRptDBField As String, _
        ByVal passwordMinLength As Integer, _
        ByVal loginIsCaseSensitive As Boolean, ByVal passwdIsCaseSensitive As Boolean, _
        ByVal loginNeedsApproval As Boolean)

        Me.clsMD.init(metadata, loginDBField, passwordDBField, passwordRptDBField)
        Me.websiteHomeFilePath = globalHomeFilePath
        Me.websiteFuncFilePath = globalFuncFilePath
        Me.loginAdmSecFilePath = loginAdmSecurityFilePath
        Me.loginMemSecFilePath = loginMemSecurityFilePath
        Me.rootIncHeaderFilePath = headerFilePath
        Me.rootIncFooterFilePath = footerFilePath
        Me.moduleFuncFilePath = moduleFuncFilePath
        Me.moduleEditTableFilePath = moduleEditTableFilePath
        Me.moduleConfirmTableFilePath = moduleConfirmTableFilePath
        Me.admHomeFilePath = adminHomeFilePath
        Me.memHomeFilePath = memberHomeFilePath

        Me.regFile = regEditFile
        Me.regConfirmFile = regConfirmFile
        Me.memRegEditFile = memRegEditFile
        Me.memRegConfirmFile = memRegConfirmFile
        Me.memRegChgPasswordFile = memRegChangePasswordFile
        Me.admRegListFile = admRegListFile
        Me.admRegEditFile = admRegEditFile
        Me.admRegConfirmFile = admRegConfirmFile
        Me.admRegChgPasswordFile = admRegChangePasswordFile
        Me.DBTableName = DBTableName

        Me.loginFld = loginDBField
        Me.passwordFld = passwordDBField
        Me.passwordRptFld = passwordRptDBField
        Me.passwordRptFldHTMLName = clsMD.passwdRptHTMLField
        Me.loginIsCaseSensitive = loginIsCaseSensitive
        Me.passwordIsCaseSensitive = passwdIsCaseSensitive
        Me.passwordMinLength = passwordMinLength
        Me.loginNeedsAdminApproval = loginNeedsApproval
    End Sub


    ' nodeRegFolder.Text = "memRegFolder"
    ' nodeReg.Text = "register.asp"
    Public Function getRegFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & Me.moduleFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "Dim err" & vbCrLf
        str = str & "err = " & Chr(34) & "" & Chr(34) & vbCrLf
        str = str & vbCrLf
        str = str & "if request(" & Chr(34) & "btnSubmit" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call saveSession()" & vbCrLf
        str = str & "    err = checkRegError()" & vbCrLf
        str = str & "    if err = " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "        Response.Redirect " & Chr(34) & regConfirmFile & Chr(34) & vbCrLf
        str = str & "    end if" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnCancel" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call clearSession()" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & Me.websiteHomeFilePath & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncHeaderFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--enter the specific registration information here-->" & vbCrLf
        str = str & "<h3>Registration Information</h3>" & vbCrLf
        str = str & "<!--end the specific registration information-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "if err <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    Response.Write " & Chr(34) & "<font color=red>Registration Error: <br>" & Chr(34) & " & err & " & Chr(34) & "</font>" & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<form  method=post name=RegisterForm id=RegisterForm>" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleEditTableFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<br>" & vbCrLf
        str = str & "<table><tr><td align=left>" & vbCrLf
        str = str & "<input type=submit name=btnSubmit value=submit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnCancel value=cancel style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "</td></tr></table>" & vbCrLf
        str = str & "</form>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncFooterFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function


    ' nodeRegConfirm.Text = "registerConfirm.asp"
    Public Function getRegConfirmFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & Me.moduleFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "if request(" & Chr(34) & "btnEdit" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & regFile & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncHeaderFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<% if request(" & Chr(34) & "btnSubmit" & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
        str = str & "<%else%>" & vbCrLf
        str = str & "<!--enter the specific confirmation information here-->" & vbCrLf
        str = str & "<p><b><font color=red>Thank you for submitting the registration." & vbCrLf
        str = str & "The following is your registration information.<br> Please print a copy for your record.</font></b></p>" & vbCrLf
        str = str & "<p><a href='../'>Back to Homepage</a></p>" & vbCrLf
        str = str & "<!--end the specific confirmation information-->" & vbCrLf
        str = str & "<%end if%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--enter the specific registration information here-->" & vbCrLf
        str = str & "<h3>Registration Information</h3>" & vbCrLf
        str = str & "<!--end the specific registration information-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<form  method=post name=RegisterConfirmForm id=RegisterConfirmForm>" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleConfirmTableFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<br>" & vbCrLf
        str = str & vbCrLf
        str = str & "<% if request(" & Chr(34) & "btnSubmit" & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
        str = str & "<table><tr><td align=left>" & vbCrLf
        str = str & "<input type=submit name=btnEdit value=Edit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnSubmit value=submit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "</td></tr></table>" & vbCrLf
        str = str & "<%else" & vbCrLf
        str = str & "    insertData()" & vbCrLf
        str = str & "    clearSession()" & vbCrLf
        str = str & "end if%>" & vbCrLf
        str = str & vbCrLf
        str = str & "</form>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncFooterFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function

    ' nodeMemRegEdit.Text = "memRegEdit.asp"
    Public Function getMemRegEditFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & Me.loginMemSecFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "Dim ID, SRC, err" & vbCrLf
        str = str & "err = " & Chr(34) & "" & Chr(34) & vbCrLf
        str = str & vbCrLf
        str = str & "ID = cint(cstr(session(" & Chr(34) & "_isMEM" & Chr(34) & ")))" & vbCrLf
        str = str & "SRC = Request(" & Chr(34) & "src" & Chr(34) & ")" & vbCrLf
        str = str & vbCrLf
        str = str & "If request(" & Chr(34) & "btnSubmit" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call saveSession()" & vbCrLf
        str = str & "    err = checkError()" & vbCrLf
        str = str & "    if err = " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "        call updateData(ID)" & vbCrLf
        str = str & "        Response.Redirect " & Chr(34) & "memRegConfirm.asp?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID" & vbCrLf
        str = str & "    end if" & vbCrLf
        str = str & "ElseIf request(" & Chr(34) & "btnCancel" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & "memRegConfirm.asp?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID" & vbCrLf
        str = str & "End If" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncHeaderFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--enter the specific registration information here-->" & vbCrLf
        str = str & "<h3>Registration Information</h3>" & vbCrLf
        str = str & "<!--end the specific registration information-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "if err <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    Response.Write " & Chr(34) & "<font color=red>Registration Error: <br>" & Chr(34) & " & err & " & Chr(34) & "</font>" & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<form  method=post name=RegisterForm id=RegisterForm>" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleEditTableFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<br>" & vbCrLf
        str = str & "<table><tr><td align=left>" & vbCrLf
        str = str & "<input type=submit name=btnSubmit value=Submit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=reset name=btnReset value=Reset style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnCancel value=Cancel style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "</td></tr></table>" & vbCrLf
        str = str & "</form>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncFooterFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function

    ' nodeMemRegConfirm.Text = "memRegConfirm.asp"
    Public Function getMemRegConfirmFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & Me.loginMemSecFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "Dim ID, SRC" & vbCrLf
        str = str & vbCrLf
        str = str & "ID = cint(cstr(session(" & Chr(34) & "_isMem" & Chr(34) & ")))" & vbCrLf
        str = str & "SRC = Request(" & Chr(34) & "src" & Chr(34) & ")" & vbCrLf
        str = str & vbCrLf
        str = str & "call retrieveData(ID)" & vbCrLf
        str = str & vbCrLf
        str = str & "if request(" & Chr(34) & "btnEdit" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	Response.Redirect " & Chr(34) & Me.memRegEditFile & Chr(34) & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnBack" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call clearSession()" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & Me.memHomeFilePath & Chr(34) & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnChgPassword" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	response.redirect " & Chr(34) & "memChgPass.asp" & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncHeaderFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--enter the specific registration information here-->" & vbCrLf
        str = str & "<h3>Registration Information</h3>" & vbCrLf
        str = str & "<!--end the specific registration information-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<form  method=post name=RegisterConfirmForm id=RegisterConfirmForm>" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleConfirmTableFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<br>" & vbCrLf
        str = str & "<br>" & vbCrLf
        str = str & "<table><tr><td align=left>" & vbCrLf
        str = str & "<input type=submit name=btnEdit value=Edit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnBack value=Back style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "</td></tr></table>" & vbCrLf
        str = str & "</form>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncFooterFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function

    ' nodeMemChgPassword.Text = "memChgPass.asp"
    Public Function getMemChgPasswordFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & Me.loginMemSecFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncHeaderFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "Dim msg, oldPassword, newPassword, newPassword2" & vbCrLf
        str = str & "oldPassword = trim(request(" & Chr(34) & "oldPassword" & Chr(34) & "))" & vbCrLf
        str = str & "newPassword = trim(request(" & Chr(34) & "newPassword" & Chr(34) & "))" & vbCrLf
        str = str & "newPassword2 = trim(request(" & Chr(34) & "newPassword2" & Chr(34) & "))" & vbCrLf
        str = str & vbCrLf
        str = str & "if request(" & Chr(34) & "btnCancel" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	response.redirect " & Chr(34) & "memRegConfirm.asp" & Chr(34) & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnSubmit" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	if oldPassword <> session(" & Chr(34) & Me.passwordFld & Chr(34) & ") then" & vbCrLf
        str = str & "		msg = " & Chr(34) & "Please enter correct old password" & Chr(34) & vbCrLf
        str = str & "	elseif checkPasswordError(newPassword, newPassword2) <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "		msg = checkPasswordError(newPassword, newPassword2)" & vbCrLf
        str = str & "	else" & vbCrLf
        str = str & "		call changePassword(session(" & Chr(34) & "_isMem" & Chr(34) & "), newPassword)" & vbCrLf
        str = str & "		'msg = " & Chr(34) & "Your password is changed." & Chr(34) & vbCrLf
        str = str & "		response.redirect " & Chr(34) & "memRegConfirm.asp" & Chr(34) & vbCrLf
        str = str & "	end if" & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & vbCrLf
        str = str & "Function changePassword(memID, newPassword)" & vbCrLf
        str = str & "	Dim sql, db" & vbCrLf
        str = str & "	sql = " & Chr(34) & "Update members SET password = " & Chr(34) & " & sqlEncode(newPassword) & " & Chr(34) & " WHERE ID = " & Chr(34) & " & memID" & vbCrLf
        str = str & "	set db = dbConnect()" & vbCrLf
        str = str & "	call dbExecute(db, sql)" & vbCrLf
        str = str & "	call dbClose(db)" & vbCrLf
        str = str & "	session(" & Chr(34) & Me.passwordFld & Chr(34) & ") = newPassword" & vbCrLf
        str = str & "End Function" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<form method=post>" & vbCrLf
        str = str & vbCrLf
        str = str & "<h3>Change Password</h3>" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "if msg <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	call writeError (msg & " & Chr(34) & "<br>" & Chr(34) & ")" & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<table cellspacing=0 cellpadding=0>" & vbCrLf
        str = str & "<tr><td>Old Password:</td><td><input type=password name=oldPassword value=" & Chr(34) & "" & Chr(34) & "></td></tr>" & vbCrLf
        str = str & "<tr><td>New Password:</td><td><input type=password name=newPassword value=" & Chr(34) & "" & Chr(34) & "></td></tr>" & vbCrLf
        str = str & "<tr><td>Repeat New Password:</td><td><input type=password name=newPassword2 value=" & Chr(34) & "" & Chr(34) & "></td></tr>" & vbCrLf
        str = str & "</table>" & vbCrLf
        str = str & vbCrLf
        str = str & "<p>" & vbCrLf
        str = str & "<input type=submit name=btnSubmit value=" & Chr(34) & "Change" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnCancel value=" & Chr(34) & "Cancel" & Chr(34) & ">" & vbCrLf
        str = str & "</p>" & vbCrLf
        str = str & vbCrLf
        str = str & "</form>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncFooterFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function

    ' nodeAdmRegList.Text = "adminRegList.asp"
    Public Function getAdmRegListFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & Me.loginAdmSecFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncHeaderFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<!--enter the specific registration information here-->" & vbCrLf
        str = str & "<h2>Registration Information for members </h2>" & vbCrLf
        str = str & "<a href='../admin/index.asp'>Back to Admin Home</a>" & vbCrLf
        str = str & "<!--end the specific registration information-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<% call getPendingList() %>" & vbCrLf
        str = str & "<% call getApprovedList() %>" & vbCrLf
        str = str & vbCrLf
        str = str & "<P align=center>&nbsp;</P>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncFooterFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function

    ' nodeAdmRegEdit.Text = "adminRegEdit.asp"
    Public Function getAdmRegEditFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & Me.loginAdmSecFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "Dim ID, memID, SRC, err" & vbCrLf
        str = str & "err = " & Chr(34) & "" & Chr(34) & vbCrLf
        str = str & vbCrLf
        str = str & "ID = Request(" & Chr(34) & "id" & Chr(34) & ")" & vbCrLf
        str = str & "memID = Request(" & Chr(34) & "memID" & Chr(34) & ")" & vbCrLf
        str = str & "SRC = Request(" & Chr(34) & "src" & Chr(34) & ")" & vbCrLf
        str = str & vbCrLf
        str = str & "If request(" & Chr(34) & "btnSubmit" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call saveSession()" & vbCrLf
        str = str & "    err = checkError()" & vbCrLf
        str = str & "    if err = " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "        call updateData(ID)" & vbCrLf
        str = str & "        Response.Redirect " & Chr(34) & Me.admRegConfirmFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID" & vbCrLf
        str = str & "    end if" & vbCrLf
        str = str & "ElseIf request(" & Chr(34) & "btnCancel" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & Me.admRegConfirmFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID" & vbCrLf
        str = str & "End If" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncHeaderFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--enter the specific registration information here-->" & vbCrLf
        str = str & "<h3>Registration Information</h3>" & vbCrLf
        str = str & "<!--end the specific registration information-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "if err <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    Response.Write " & Chr(34) & "<font color=red>Registration Error: <br>" & Chr(34) & " & err & " & Chr(34) & "</font>" & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<form  method=post name=RegisterForm id=RegisterForm>" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleEditTableFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<br>" & vbCrLf
        str = str & "<table><tr><td align=left>" & vbCrLf
        str = str & "<input type=submit name=btnSubmit value=Submit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=reset name=btnReset value=Reset style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnCancel value=Cancel style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "</td></tr></table>" & vbCrLf
        str = str & "</form>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncFooterFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function

    ' nodeAdmRegConfirm.Text = "adminRegConfirm.asp"
    Public Function getAdmRegConfirmFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & Me.loginAdmSecFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "Dim ID, memID, SRC" & vbCrLf
        str = str & vbCrLf
        str = str & "ID = Request(" & Chr(34) & "id" & Chr(34) & ")" & vbCrLf
        str = str & "memID = Request(" & Chr(34) & "memID" & Chr(34) & ")" & vbCrLf
        str = str & "SRC = Request(" & Chr(34) & "src" & Chr(34) & ")" & vbCrLf
        str = str & vbCrLf
        str = str & "call retrieveData(ID)" & vbCrLf
        str = str & vbCrLf
        str = str & "if request(" & Chr(34) & "btnEdit" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "'To edit" & vbCrLf
        str = str & "Response.Redirect " & Chr(34) & Me.admRegEditFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID" & vbCrLf
        str = str & "'To go back" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnBack" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call clearSession()" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & Me.admRegListFile & Chr(34) & vbCrLf
        str = str & "'To approve" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnApprove" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnYesApprove" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call approveData(ID)" & vbCrLf
        str = str & "    call clearSession()" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & Me.admRegListFile & Chr(34) & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnNoApprove" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "'To disapprove" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnDisApprove" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnYesDisApprove" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call disapproveData(ID)" & vbCrLf
        str = str & "    call clearSession()" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & Me.admRegListFile & Chr(34) & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnNoDisApprove" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "'To delete" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnDelete" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnYesDelete" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call deleteData(ID)" & vbCrLf
        str = str & "    call clearSession()" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & "adminRegList.asp" & Chr(34) & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnNoDelete" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnChgPassword" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	response.redirect " & Chr(34) & Me.admRegChgPasswordFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&ID=" & Chr(34) & " & ID" & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncHeaderFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--enter the specific registration information here-->" & vbCrLf
        str = str & "<h3>Registration Information</h3>" & vbCrLf
        str = str & "<!--end the specific registration information-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<form  method=post name=RegisterConfirmForm id=RegisterConfirmForm>" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleConfirmTableFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<br>" & vbCrLf
        str = str & vbCrLf
        str = str & "<br>" & vbCrLf
        str = str & "<table><tr><td align=left>" & vbCrLf
        str = str & "<%if request(" & Chr(34) & "btnDelete" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
        str = str & "<font color=red><b>Are you sure to delete?</b></font>" & vbCrLf
        str = str & "<input type=submit name=btnYesDelete value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnNoReset value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<%elseif request(" & Chr(34) & "btnApprove" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
        str = str & "<font color=red><b>Are you sure to appove?</b></font>" & vbCrLf
        str = str & "<input type=submit name=btnYesApprove value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnNoApprove value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<%elseif request(" & Chr(34) & "btnDisapprove" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
        str = str & "<font color=red><b>Are you sure to disappove?</b></font>" & vbCrLf
        str = str & "<input type=submit name=btnYesDisapprove value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnNoDisapprove value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<%else%>" & vbCrLf
        str = str & "<input type=submit name=btnDelete value=Delete style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<%if SRC=" & Chr(34) & "pending" & Chr(34) & "then%>" & vbCrLf
        str = str & "<input type=submit name=btnApprove value=Approve style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<%elseif SRC=" & Chr(34) & "approved" & Chr(34) & "then%>" & vbCrLf
        str = str & "<input type=submit name=btnDisapprove value=Disapprove style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<%end if%>" & vbCrLf
        str = str & "<input type=submit name=btnEdit value=Edit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnBack value=Back style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">" & vbCrLf
        str = str & "<%end if%>" & vbCrLf
        str = str & "</td></tr></table>" & vbCrLf
        str = str & "</form>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncFooterFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function


    ' nodeAdmChgPassword.Text = "adminChgPass.asp"
    Public Function getAdmChgPasswordFile() As String
        'Dim admRegConfirmFile As String = Me.nodeAdmRegConfirm.Text
        'Dim loginTableName As String = "members"
        'Dim passwordFld As String = Me.TextBoxPassword.Text

        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & loginAdmSecFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.moduleFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & rootIncHeaderFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "Dim msg, oldPassword, newPassword, newPassword2" & vbCrLf
        str = str & vbCrLf
        str = str & "newPassword = trim(request(" & Chr(34) & "newPassword" & Chr(34) & "))" & vbCrLf
        str = str & "newPassword2 = trim(request(" & Chr(34) & "newPassword2" & Chr(34) & "))" & vbCrLf
        str = str & vbCrLf
        str = str & "if request(" & Chr(34) & "btnCancel" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	response.redirect " & Chr(34) & admRegConfirmFile & "?src=" & Chr(34) & " & request(" & Chr(34) & "src" & Chr(34) & ") & " & Chr(34) & "&id=" & Chr(34) & " & request(" & Chr(34) & "id" & Chr(34) & ")" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnSubmit" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	If checkPasswordError(newPassword, newPassword2) <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "		msg = checkPasswordError(newPassword, newPassword2)" & vbCrLf
        str = str & "	else" & vbCrLf
        str = str & "		call changePassword(request(" & Chr(34) & "id" & Chr(34) & "), newPassword)" & vbCrLf
        str = str & "		'msg = " & Chr(34) & "Your password is changed." & Chr(34) & vbCrLf
        str = str & "		response.redirect " & Chr(34) & admRegConfirmFile & "?src=" & Chr(34) & " & request(" & Chr(34) & "src" & Chr(34) & ") & " & Chr(34) & "&id=" & Chr(34) & " & request(" & Chr(34) & "id" & Chr(34) & ")" & vbCrLf
        str = str & "	end if" & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & vbCrLf
        str = str & "Function changePassword(memID, newPassword)" & vbCrLf
        str = str & "	Dim sql, db" & vbCrLf
        str = str & "	sql = " & Chr(34) & "Update " & DBTableName & " SET " & passwordFld & " = " & Chr(34) & " & sqlEncode(newPassword) & " & Chr(34) & " WHERE ID = " & Chr(34) & " & memID" & vbCrLf
        str = str & "	set db = dbConnect()" & vbCrLf
        str = str & "	call dbExecute(db, sql)" & vbCrLf
        str = str & "	call dbClose(db)" & vbCrLf
        str = str & "	session(" & Chr(34) & Me.passwordFld & Chr(34) & ") = newPassword" & vbCrLf
        str = str & "End Function" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<form method=post>" & vbCrLf
        str = str & vbCrLf
        str = str & "<h3>Change Member Password</h3>" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "if msg <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "	call writeError (msg & " & Chr(34) & "<br>" & Chr(34) & ")" & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<table cellspacing=0 cellpadding=0>" & vbCrLf
        str = str & "<tr><td>New Password:</td><td><input type=password name=newPassword value=" & Chr(34) & "" & Chr(34) & "></td></tr>" & vbCrLf
        str = str & "<tr><td>Repeat New Password:</td><td><input type=password name=newPassword2 value=" & Chr(34) & "" & Chr(34) & "></td></tr>" & vbCrLf
        str = str & "</table>" & vbCrLf
        str = str & vbCrLf
        str = str & "<p>" & vbCrLf
        str = str & "<input type=submit name=btnSubmit value=" & Chr(34) & "Change" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=submit name=btnCancel value=" & Chr(34) & "Cancel" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=hidden name=ID value=" & Chr(34) & "<%=request(" & Chr(34) & "id" & Chr(34) & ")%>" & Chr(34) & ">" & vbCrLf
        str = str & "<input type=hidden name=SRC value=" & Chr(34) & "<%=request(" & Chr(34) & "src" & Chr(34) & ")%>" & Chr(34) & ">" & vbCrLf
        str = str & "</p>" & vbCrLf
        str = str & vbCrLf
        str = str & "</form>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & rootIncFooterFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function


    ' nodeIncEditTable.Text = "editTable.asp"
    Public Function getIncEditTableFile(Optional ByVal tblAttr _
        As clsHTMLTableAttributes = Nothing) As String
        Dim i, j, count As Integer
        Dim metadataArray() As String = Me.clsMD.metadataArray

        Dim tableStart, firstColWidth, secondColWidth As String
        xcUtil.getTableAttributes(tblAttr, tableStart, firstColWidth, secondColWidth)

        Dim str As String = tableStart & vbCrLf & vbCrLf

        'str = str & "<%If session(" & Chr(34) & "regDate" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " Then" & vbCrLf
        'str = str & "	Response.Write " & Chr(34) & "<tr><td" & firstColWidth & ">Register Date </td><td" & secondColWidth & ">" & Chr(34) & " & session(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td></tr>" & Chr(34) & vbCrLf
        'str = str & "End if%>" & vbCrLf & vbCrLf

        count = UBound(metadataArray) - 1
        For i = 0 To count
            If LCase(metadataArray(i + 1)) = "text" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf ' label
                If LCase(metadataArray(i + 5)) = LCase(Me.loginFld) Then str &= notEmptyMark & vbCrLf
                str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf

                If LCase(metadataArray(i + 5)) = LCase(Me.loginFld) Then
                    str = str & "<%If session(" & Chr(34) & "_isMem" & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
                    str = str & vbIndent & ("<input type=" & metadataArray(i + 1) & " name=" & metadataArray(i + 2) & " size=" & metadataArray(i + 3) & " maxlength=" & metadataArray(i + 4) & " value=" & Chr(34) & "<%=TextHTMLEncode( session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ") )%>" & Chr(34) & ">") & vbCrLf
                    str = str & "<%Else%>" & vbCrLf
                    str = str & vbIndent & ("<%=TextHTMLEncode( session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ") )%>") & vbCrLf
                    str = str & vbIndent & ("<input type=hidden name=" & metadataArray(i + 2) & " size=" & metadataArray(i + 3) & " maxlength=" & metadataArray(i + 4) & " value=" & Chr(34) & "<%=TextHTMLEncode( session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ") )%>" & Chr(34) & ">") & vbCrLf
                    str = str & "<%End If%>" & vbCrLf
                Else
                    str = str & ("<input type=" & metadataArray(i + 1) & " name=" & metadataArray(i + 2) & " size=" & metadataArray(i + 3) & " maxlength=" & metadataArray(i + 4) & " value=" & Chr(34) & "<%=TextHTMLEncode( session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ") )%>" & Chr(34) & ">") & vbCrLf
                End If

                str = str & ("</td></tr>") & vbCrLf
                i = i + 5
            ElseIf LCase(metadataArray(i + 1)) = "password" Then
                If LCase(metadataArray(i + 5)) = LCase(Me.passwordFld) Then
                    str = str & vbIndent & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                    str = str & vbIndent & (metadataArray(i)) & vbCrLf  ' label
                    str = str & notEmptyMark & vbCrLf
                    str = str & vbIndent & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf

                    str = str & "<%If session(" & Chr(34) & "_isMem" & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
                    str = str & vbIndent & ("<input type=password name=" & metadataArray(i + 2) & " size=" & metadataArray(i + 3) & " maxlength=" & metadataArray(i + 4) & " value=" & Chr(34) & "<%=TextHTMLEncode( session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ") )%>" & Chr(34) & ">") & vbCrLf
                    str = str & "<%Else%>" & vbCrLf
                    str = str & vbIndent & "********" & vbCrLf
                    str = str & vbIndent & ("<input type=hidden name=" & metadataArray(i + 2) & " size=" & metadataArray(i + 3) & " maxlength=" & metadataArray(i + 4) & " value=" & Chr(34) & "<%=TextHTMLEncode( session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ") )%>" & Chr(34) & ">") & vbCrLf
                    str = str & "<%End If%>" & vbCrLf

                    str = str & vbIndent & ("</td></tr>") & vbCrLf
                ElseIf LCase(metadataArray(i + 5)) = LCase(Me.passwordRptFld) Then
                    ' the "Repeat password" field, so not shown if not New registration.
                    str = str & "<%If session(" & Chr(34) & "_isMem" & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
                    str = str & vbIndent & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                    str = str & vbIndent & (metadataArray(i)) & vbCrLf  ' label
                    str = str & notEmptyMark & vbCrLf
                    str = str & vbIndent & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf
                    str = str & vbIndent & ("<input type=password name=" & metadataArray(i + 2) & " size=" & metadataArray(i + 3) & " maxlength=" & metadataArray(i + 4) & " value=" & Chr(34) & "<%=TextHTMLEncode( session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ") )%>" & Chr(34) & ">") & vbCrLf
                    str = str & vbIndent & ("</td></tr>") & vbCrLf
                    str = str & "<%End If%>" & vbCrLf
                Else ' a password field, but neither login password nor passwordRpt field
                    str = str & vbIndent & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                    str = str & vbIndent & (metadataArray(i)) & vbCrLf  ' label
                    str = str & vbIndent & ("</td><td valign=top>") & vbCrLf
                    str = str & vbIndent & ("<input type=password name=" & metadataArray(i + 2) & " size=" & metadataArray(i + 3) & " maxlength=" & metadataArray(i + 4) & " value=" & Chr(34) & "<%=TextHTMLEncode( session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ") )%>" & Chr(34) & ">") & vbCrLf
                    str = str & vbIndent & ("</td></tr>") & vbCrLf
                End If
                i = i + 5
            ElseIf LCase(metadataArray(i + 1)) = "hidden" Then
                'str = str & ("<tr><td valign=top>") & vbCrLf
                'str = str & (metadataArray(i)) & vbCrLf ' label
                'str = str & ("</td><td valign=top>") & vbCrLf
                str = str & ("<input type=hidden name=" & metadataArray(i + 2) & " size=" & metadataArray(i + 3) & " maxlength=" & metadataArray(i + 4) & " value=" & Chr(34) & "<%=TextHTMLEncode( session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ") )%>" & Chr(34) & ">") & vbCrLf
                'str = str & ("</td></tr>") & vbCrLf
                i = i + 5
            ElseIf LCase(metadataArray(i + 1)) = "select" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf
                str = str & ("<select name=" & metadataArray(i + 2) & " style=" & Chr(34) & "WIDTH: " & metadataArray(i + 3) & Chr(34) & " size=" & metadataArray(i + 4) & ">") & vbCrLf
                count = CInt(metadataArray(i + 5))
                For j = 1 To count
                    str = str & ("<option value=" & j & " <" & "%if session(" & Chr(34) & metadataArray(i + 6) & Chr(34) & ") = " & Chr(34) & j & Chr(34) & " then%" & "> selected <" & "%end if%" & ">" & ">" & metadataArray(i + 6 + j)) & vbCrLf
                Next
                str = str & ("</select>") & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 5 + count
            ElseIf LCase(metadataArray(i + 1)) = "textarea" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf
                str = str & ("<TEXTAREA NAME=" & metadataArray(i + 2) & " ROWS=" & CInt(metadataArray(i + 3)) & " COLS=" & CInt(metadataArray(i + 4)) & " " & metadataArray(i + 5) & ">" & "<%=TextAreaEditHTMLEncode( session(" & Chr(34) & metadataArray(i + 6) & Chr(34) & ") )%>" & "</TEXTAREA>") & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 6
            ElseIf LCase(metadataArray(i + 1)) = "radio" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf
                count = CInt(metadataArray(i + 3))
                For j = 1 To count
                    str = str & (metadataArray(i + 4 + j) & "<input type=radio name=" & metadataArray(i + 2) & " value=" & j & " <" & "%if session(" & Chr(34) & metadataArray(i + 4) & Chr(34) & ")=" & Chr(34) & j & Chr(34) & " then%" & "> checked <" & "%end if%" & ">>") & vbCrLf
                Next
                str = str & ("</td></tr>") & vbCrLf
                i = i + 3 + count
            ElseIf LCase(metadataArray(i + 1)) = "checkbox" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf
                'str = str &  "<input type=checkbox name=" & chr(34) & metadataArray(i+2) & chr(34) & " value=" & chr(34) & metadataArray(i+3) & chr(34) & " <" & "%if session(" & chr(34) & metadataArray(i+4) & chr(34) & ") <> " & chr(34) & chr(34) & " then%" & "> checked <" & "%end if%" & ">>" & metadataArray(i+5)
                ' Note: Now the value is hardcoded to "Y".
                str = str & ("<input type=checkbox name=" & Chr(34) & metadataArray(i + 2) & Chr(34) & " value=" & Chr(34) & "Y" & Chr(34) & " <" & "%if session(" & Chr(34) & metadataArray(i + 4) & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then%" & "> checked <" & "%end if%" & ">>" & metadataArray(i + 5)) & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 4
            ElseIf LCase(metadataArray(i + 1)) = "comment" Then
                str = str & ("<tr><td valign=top colspan=2>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 1
            ElseIf LCase(metadataArray(i + 1)) = "blankline" Then
                str = str & ("<tr><td valign=top colspan=2>") & vbCrLf
                For j = 1 To CInt(metadataArray(i))
                    str = str & ("<br>") & vbCrLf
                Next
                str = str & ("</td></tr>") & vbCrLf
                i = i + 1
            ElseIf LCase(metadataArray(i + 1)) = "hr" Then
                str = str & ("<tr><td valign=top colspan=2>") & vbCrLf
                str = str & ("<hr>") & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 1
            ElseIf LCase(metadataArray(i + 1)) = "newtable" Then
                str = str & ("</table>") & vbCrLf
                str = str & tableStart & vbCrLf
                If metadataArray(i) <> "" Then  ' First col width of the new table
                    str = str & "<tr><td width=" & metadataArray(i) & "><br></td>" & "<td><br></td></tr>" & vbCrLf
                Else
                    str = str & "<tr><td" & firstColWidth & "><br></td>" & "<td" & secondColWidth & "><br></td></tr>" & vbCrLf
                End If
                i = i + 1
            End If
            str = str & vbCrLf
        Next

        str = str & ("</table>") & vbCrLf

        Return str
    End Function


    ' nodeIncConfirmTable.Text = "confirTable.asp"
    Public Function getIncConfirmTableFile(Optional ByVal tblAttr _
        As clsHTMLTableAttributes = Nothing) As String
        Dim metadataArray() As String = Me.clsMD.metadataArray
        Dim i, j, count As Integer

        Dim tableStart, firstColWidth, secondColWidth As String
        xcUtil.getTableAttributes(tblAttr, tableStart, firstColWidth, secondColWidth)

        Dim str As String = tableStart & vbCrLf & vbCrLf

        'str = str & "<%If session(" & Chr(34) & "regDate" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " Then" & vbCrLf
        'str = str & "	Response.Write " & Chr(34) & "<tr><td" & firstColWidth & ">Register Date </td><td" & secondColWidth & ">" & Chr(34) & " & session(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td></tr>" & Chr(34) & vbCrLf
        'str = str & "End if%>" & vbCrLf & vbCrLf

        count = UBound(metadataArray) - 1
        For i = 0 To count
            If LCase(metadataArray(i + 1)) = "text" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf ' label
                If LCase(metadataArray(i + 5)) = LCase(Me.loginFld) Then str &= notEmptyMark & vbCrLf
                str = str & ("</td><td" & secondColWidth & " valign=top><b>") & vbCrLf
                str = str & ("<%=session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ")%>" & "&nbsp;") & vbCrLf
                str = str & ("</b></td></tr>") & vbCrLf
                i = i + 5
            ElseIf LCase(metadataArray(i + 1)) = "password" Then
                If LCase(metadataArray(i + 5)) = LCase(Me.passwordFld) Then
                    str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                    str = str & (metadataArray(i)) & vbCrLf ' label
                    str = str & notEmptyMark & vbCrLf
                    str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf
                    str = str & ("<b>********</b>") & vbCrLf
                    str = str & "<%if session(" & Chr(34) & "_isMem" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
                    str = str & vbIndent & "<input type=submit name=btnChgPassword value=" & Chr(34) & "Change Password" & Chr(34) & ">" & vbCrLf
                    str = str & "<%end if%>" & vbCrLf
                    str = str & ("</td></tr>") & vbCrLf
                ElseIf LCase(metadataArray(i + 5)) = LCase(Me.passwordRptFld) Then
                    str = str & "<%if session(" & Chr(34) & "_isMem" & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & " then%>" & vbCrLf
                    str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                    str = str & (metadataArray(i)) & vbCrLf ' label
                    str = str & notEmptyMark & vbCrLf
                    str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf
                    str = str & ("<b>********</b>") & vbCrLf
                    str = str & ("</td></tr>") & vbCrLf
                    str = str & "<%end if%>" & vbCrLf
                Else
                    str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                    str = str & (metadataArray(i)) & vbCrLf ' label
                    str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf
                    str = str & ("<b>********</b>") & vbCrLf
                    str = str & ("</td></tr>") & vbCrLf
                End If
                i = i + 5
            ElseIf LCase(metadataArray(i + 1)) = "hidden" Then
                i = i + 5
            ElseIf LCase(metadataArray(i + 1)) = "select" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf
                count = CInt(metadataArray(i + 5))
                For j = 1 To count
                    str = str & ("<" & "%if session(" & Chr(34) & metadataArray(i + 6) & Chr(34) & ") = " & Chr(34) & j & Chr(34) & " then%" & "><b>" & metadataArray(i + 6 + j) & "</b><" & "%end if%" & ">") & vbCrLf
                Next
                str = str & ("&nbsp;</td></tr>") & vbCrLf
                i = i + 5 + count
            ElseIf LCase(metadataArray(i + 1)) = "textarea" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf

                str = str & ("<" & "%if session(" & Chr(34) & metadataArray(i + 6) & Chr(34) & ") <> " & _
                Chr(34) & Chr(34) & " then%" & "><b>" & "<%=TextAreaViewHTMLEncode( session(" & Chr(34) & metadataArray(i + 6) & Chr(34) & ") )%>" & "</b><" & _
                "%else%" & "> <br> <" & "%end if%" & ">") & vbCrLf

                str = str & ("</td></tr>") & vbCrLf
                i = i + 6
            ElseIf LCase(metadataArray(i + 1)) = "radio" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td" & secondColWidth & " valign=top>") & vbCrLf
                count = CInt(metadataArray(i + 3))
                For j = 1 To count
                    str = str & "<" & "%if session(" & Chr(34) & metadataArray(i + 4) & _
                      Chr(34) & ") = " & Chr(34) & j & Chr(34) & " then%" & "><b>" & _
                      metadataArray(i + 4 + j) & "</b><" & "%end if%" & ">" & vbCrLf
                Next
                str = str & ("&nbsp;</td></tr>") & vbCrLf
                i = i + 3 + count
            ElseIf LCase(metadataArray(i + 1)) = "checkbox" Then
                str = str & "<" & "%if session(" & Chr(34) & metadataArray(i + 4) & _
                  Chr(34) & ") <> " & Chr(34) & Chr(34) & " then%" & ">" & _
                  "<tr><td" & firstColWidth & " valign=top>" & metadataArray(i) & "</td><td" & secondColWidth & " valign=top><b>" & _
                  metadataArray(i + 5) & "</b></td></tr>" & "<" & "%end if%" & ">" & vbCrLf
                i = i + 4
            ElseIf LCase(metadataArray(i + 1)) = "comment" Then
                str = str & ("<tr><td valign=top colspan=2>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 1
            ElseIf LCase(metadataArray(i + 1)) = "blankline" Then
                str = str & ("<tr><td valign=top colspan=2>") & vbCrLf
                For j = 1 To CInt(metadataArray(i))
                    str = str & ("<br>") & vbCrLf
                Next
                str = str & ("</td></tr>") & vbCrLf
                i = i + 1
            ElseIf LCase(metadataArray(i + 1)) = "hr" Then
                str = str & ("<tr><td valign=top colspan=2>") & vbCrLf
                str = str & ("<hr>") & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 1
            ElseIf LCase(metadataArray(i + 1)) = "newtable" Then
                str = str & ("</table>") & vbCrLf
                str = tableStart & vbCrLf
                If metadataArray(i) <> "" Then  ' First col width of the new table
                    str = str & "<tr><td width=" & metadataArray(i) & "><br></td>" & "<td><br></td></tr>" & vbCrLf
                Else
                    str = str & "<tr><td" & firstColWidth & "><br></td>" & "<td" & secondColWidth & "><br></td></tr>" & vbCrLf
                End If
                'str = str & ("<table cellpadding=0 cellspacing=0 border = 0>") & vbCrLf
                'tdWidth = CInt(metadataArray(i))
                'If tdWidth >= 0 And tdWidth <= 100 Then
                'str = str & ("<tr><td width=" & tdWidth & "%" & "><br></td>" & "<td width=" & (100 - tdWidth) & "%" & "><br></td></tr>")
                'End If
                i = i + 1
            End If
            str = str & vbCrLf
        Next

        str = str & ("</table>")
        Return str
    End Function


    ' nodeXMLFileName.Text = "memreg.xml"
    Public Function getXMLFile() As String
        Dim str As String = Me.clsMD.generateXML
        Return str
    End Function


    Private Function getLoginCaseSensitiveParams( _
        ByRef caseIndent As String, ByRef caseConditionIf As String)

        caseIndent = vbIndent

        If Me.loginIsCaseSensitive And Me.passwordIsCaseSensitive Then
            caseConditionIf = caseIndent & caseIndent & _
                "if caseSensitiveEqual(login, rs(" & Chr(34) & Me.loginFld & Chr(34) & _
                ")) and caseSensitiveEqual(password, rs(" & Chr(34) & Me.passwordFld & _
                Chr(34) & ")) then" & vbCrLf
        ElseIf Me.loginIsCaseSensitive And Not Me.passwordIsCaseSensitive Then
            caseConditionIf = caseIndent & caseIndent & _
                "if caseSensitiveEqual(login, rs(" & Chr(34) & Me.loginFld & Chr(34) & ")) then" & vbCrLf
        ElseIf Not Me.loginIsCaseSensitive And Me.passwordIsCaseSensitive Then
            caseConditionIf = caseIndent & caseIndent & _
                "if caseSensitiveEqual(password, rs(" & Chr(34) & Me.passwordFld & Chr(34) & ")) then" & vbCrLf
        Else
            caseConditionIf = ""
            caseIndent = ""
        End If

    End Function


    ' Create login file using database register table account/password fields
    Public Function getLoginLoginFile() As String
        Dim caseSensitive As Boolean = (Me.loginIsCaseSensitive Or Me.passwordIsCaseSensitive)

        Dim caseIndent As String '= vbIndent
        Dim caseConditionIf As String
        Me.getLoginCaseSensitiveParams(caseIndent, caseConditionIf)

        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & Me.websiteFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "Dim login, password, validateResult" & vbCrLf
        str = str & vbCrLf
        str = str & "login = trim(request(" & Chr(34) & "usrLoginname" & Chr(34) & "))" & vbCrLf
        str = str & "password = trim(request(" & Chr(34) & "usrPassword" & Chr(34) & "))" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "validateResult = loginValidate(login, password)" & vbCrLf
        str = str & vbCrLf
        str = str & "if validateResult = " & Chr(34) & "true" & Chr(34) & " then" & vbCrLf
        str = str & "    Response.redirect " & Chr(34) & Me.memHomeFilePath & Chr(34) & vbCrLf
        str = str & "elseif validateResult = " & Chr(34) & "notFound" & Chr(34) & " then" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & Me.websiteHomeFilePath & "?login='" & Chr(34) & " & login & " & Chr(34) & "'&error=Incorrect login information" & Chr(34) & vbCrLf
        str = str & "elseif validateResult = " & Chr(34) & "notApproved" & Chr(34) & " then" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & Me.websiteHomeFilePath & "?login='" & Chr(34) & " & login & " & Chr(34) & "'&error=Your registration is not approved yet. Please wait." & Chr(34) & vbCrLf
        str = str & "elseif validateResult = " & Chr(34) & "duplicate" & Chr(34) & " then" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & Me.websiteHomeFilePath & "?login='" & Chr(34) & " & login & " & Chr(34) & "'&error=You have a duplicate email with another user. Please contact your system administrator." & Chr(34) & vbCrLf
        str = str & "else" & vbCrLf
        str = str & "    Response.Redirect " & Chr(34) & Me.websiteHomeFilePath & "?login='" & Chr(34) & " & login & " & Chr(34) & "'&error=Unknown error. Please contact your system administrator." & Chr(34) & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & vbCrLf
        str = str & "function loginValidate (login, password)" & vbCrLf
        str = str & "    dim db, sql" & vbCrLf
        str = str & vbCrLf
        str = str & "    sql_head = " & Chr(34) & "SELECT * FROM " & Me.DBTableName & " where " & Me.loginFld & "='" & Chr(34) & " & login & " & Chr(34) & "' and " & Me.passwordFld & "='" & Chr(34) & " & password & " & Chr(34) & "'" & Chr(34) & vbCrLf
        str = str & "    sql = sql_head & " & Chr(34) & " AND approval = 1" & Chr(34) & vbCrLf
        str = str & vbCrLf
        str = str & "    set db = dbConnect()" & vbCrLf
        str = str & "    set rs = dbExecuteRS(db, sql)" & vbCrLf
        str = str & vbCrLf
        str = str & "    if rs.recordcount > 1 then" & vbCrLf
        str = str & "        loginValidate = " & Chr(34) & "duplicate" & Chr(34) & "	' this should not happen." & vbCrLf
        str = str & "    elseif rs.recordcount = 0 then" & vbCrLf
        str = str & "        sql = sql_head & " & Chr(34) & " AND approval = 0" & Chr(34) & vbCrLf
        str = str & "        set rs = dbExecuteRS(db, sql)" & vbCrLf
        str = str & "        if rs.recordcount = 0 then" & vbCrLf
        str = str & "            loginValidate = " & Chr(34) & "notFound" & Chr(34) & vbCrLf
        str = str & "        else" & vbCrLf

        If caseSensitive Then str = str & caseIndent & caseConditionIf
        str = str & caseIndent & "            loginValidate = " & Chr(34) & "notApproved" & Chr(34) & vbCrLf
        If caseSensitive Then
            str = str & caseIndent & "		else" & vbCrLf
            str = str & caseIndent & "			loginValidate = " & Chr(34) & "notFound" & Chr(34) & vbCrLf
            str = str & caseIndent & "		end if" & vbCrLf
        End If

        str = str & "        end if" & vbCrLf
        str = str & "    else" & vbCrLf

        If caseSensitive Then str = str & caseConditionIf
        str = str & caseIndent & "        loginValidate = " & Chr(34) & "true" & Chr(34) & vbCrLf
        str = str & caseIndent & "        session(" & Chr(34) & "_isMem" & Chr(34) & ") = rs(" & Chr(34) & "ID" & Chr(34) & ")" & vbCrLf
        str = str & caseIndent & "        if rs(" & Chr(34) & "isAdmin" & Chr(34) & ") = 1 then" & vbCrLf
        str = str & caseIndent & "            session(" & Chr(34) & "_isAdmin" & Chr(34) & ") = rs(" & Chr(34) & "ID" & Chr(34) & ")" & vbCrLf
        str = str & caseIndent & "        end if" & vbCrLf
        If caseSensitive Then
            str = str & "		else" & vbCrLf
            str = str & "			loginValidate = " & Chr(34) & "notFound" & Chr(34) & vbCrLf
            str = str & "		end if" & vbCrLf
        End If

        str = str & "    end if" & vbCrLf
        str = str & vbCrLf
        str = str & "    call dbRSClose(db, rs)" & vbCrLf
        str = str & vbCrLf
        str = str & "end function" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf

        Return str
    End Function



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SQL file
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Function getRegSQLFile()
        Dim sqlIndent As String = "    "
        ' Given dbFieldLenArray, dbArray
        Dim count As Integer
        Dim i As Integer
        Dim hasTextField As Boolean = False

        Dim str As String = ""

        str = str & "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & Me.DBTableName & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        str = str & "drop table [dbo].[" & Me.DBTableName & "]" & vbCrLf
        str = str & "GO" & vbCrLf

        str = str & vbCrLf & vbCrLf

        str = str & "CREATE TABLE [dbo].[" & Me.DBTableName & "] (" & vbCrLf
        str = str & sqlIndent & "[ID] [int] IDENTITY (1, 1) NOT NULL ," & vbCrLf
        str = str & sqlIndent & "[regDate] [datetime] NULL ," & vbCrLf
        str = str & sqlIndent & "[isAdmin] [varchar] (1) NULL ," & vbCrLf
        str = str & sqlIndent & "[approval] [varchar] (1) NULL ," & vbCrLf

        count = UBound(Me.clsMD.dbFieldLenArray) - 1
        For i = 0 To count
            If UCase(Me.clsMD.dbFieldLenArray(i)) = "TEXT" Then
                str = str & sqlIndent & "[" & Me.clsMD.dbArray(i) & "] [text] NULL ," & vbCrLf
                hasTextField = True
            Else
                str = str & sqlIndent & "[" & Me.clsMD.dbArray(i) & "] [varchar] (" & CInt(Me.clsMD.dbFieldLenArray(i)) & ") NULL ," & vbCrLf
            End If
        Next

        If UCase(Me.clsMD.dbFieldLenArray(i)) = "TEXT" Then
            str = str & sqlIndent & "[" & Me.clsMD.dbArray(i) & "] [text] NULL" & vbCrLf
            hasTextField = True
        Else
            str = str & sqlIndent & "[" & Me.clsMD.dbArray(i) & "] [varchar] (" & CInt(Me.clsMD.dbFieldLenArray(i)) & ") NULL" & vbCrLf
        End If

        If hasTextField = True Then
            str = str & ") ON [PRIMARY] TEXTIMAGE_ON  [PRIMARY]" & vbCrLf
        Else
            str = str & ") ON [PRIMARY]" & vbCrLf
        End If
        str = str & "GO" & vbCrLf

        str = str & vbCrLf
        str = str & "ALTER TABLE [dbo].[" & Me.DBTableName & "] WITH NOCHECK ADD " & vbCrLf
        str = str & sqlIndent & "CONSTRAINT [PK_" & Me.DBTableName & "] PRIMARY KEY  NONCLUSTERED " & vbCrLf
        str = str & sqlIndent & "(" & vbCrLf
        str = str & sqlIndent & sqlIndent & "[ID]" & vbCrLf
        str = str & sqlIndent & ") ON [PRIMARY]" & vbCrLf
        str = str & "GO" & vbCrLf

        str = str & vbCrLf
        str = str & "INSERT INTO " & Me.DBTableName & " (" & vbCrLf
        str = str & sqlIndent & "regDate, approval, isAdmin, " & Me.loginFld & ", " & Me.passwordFld & vbCrLf
        str = str & ") VALUES (" & vbCrLf
        str = str & sqlIndent & "'" & Now() & "', 1, 1, 'admin', 'password'" & vbCrLf
        str = str & ")" & vbCrLf

        Return str
    End Function


    Public Function getApproveAdminFile() As String
        Dim str As String = ""
        str = str & "<!--#include file=" & Chr(34) & Me.loginAdmSecFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.websiteFuncFilePath & Chr(34) & "-->" & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncHeaderFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & "if request(" & Chr(34) & "btnApprove" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call approveAdm(trim(request(" & Chr(34) & "pendingAdmin" & Chr(34) & ")))" & vbCrLf
        str = str & "elseif request(" & Chr(34) & "btnDisapprove" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "    call DisapproveAdm(trim(request(" & Chr(34) & "approvedAdmin" & Chr(34) & ")))" & vbCrLf
        str = str & "end if" & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf
        str = str & "<h2>Approve Administrator</h2>" & vbCrLf
        str = str & "<p><a href='" & Me.admHomeFilePath & "'>Back to Admin Home</a></p>" & vbCrLf
        str = str & "<p><br><p>" & vbCrLf
        str = str & vbCrLf
        str = str & "<form id=form1 name=form1>" & vbCrLf
        str = str & vbCrLf
        str = str & "<table border=0>" & vbCrLf
        str = str & "<tr><td>Pending administrators:</td><td> <%call getPendingAdmin()%></td><td><input type=submit name=btnApprove value=Approve></td></tr>" & vbCrLf
        str = str & "<tr><td>Approved administrators:</td><td> <%call getApprovedAdmin()%></td><td><input type=submit name=btnDisapprove value=Disapprove></td></tr>" & vbCrLf
        str = str & "</table>" & vbCrLf
        str = str & vbCrLf
        str = str & "</form>" & vbCrLf
        str = str & vbCrLf
        str = str & "<!--#include file=" & Chr(34) & Me.rootIncFooterFilePath & Chr(34) & "-->" & vbCrLf
        str = str & vbCrLf
        str = str & "<%" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & "' Assign administrator power to a member." & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & vbCrLf
        str = str & "sub approveAdm(ID)" & vbCrLf
        str = str & "    call updateTableByID(" & Chr(34) & Me.DBTableName & Chr(34) & ", " & Chr(34) & "isAdmin" & Chr(34) & ", " & Chr(34) & "1" & Chr(34) & ", ID)" & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & "' Revoke the administrator power assigned to a member." & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & vbCrLf
        str = str & "sub disapproveAdm(ID)" & vbCrLf
        str = str & "    call updateTableByID(" & Chr(34) & Me.DBTableName & Chr(34) & ", " & Chr(34) & "isAdmin" & Chr(34) & ", " & Chr(34) & "0" & Chr(34) & ", ID)" & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & "' Get a list of non-admin members." & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & vbCrLf
        str = str & "sub getPendingAdmin()" & vbCrLf
        str = str & "    dim db, rs, sql" & vbCrLf
        str = str & vbCrLf
        str = str & "    sql = " & Chr(34) & "SELECT * FROM " & Me.DBTableName & " where approval = 1 and isAdmin = 0" & Chr(34) & vbCrLf
        str = str & "    set db = dbConnect()" & vbCrLf
        str = str & "    set rs = dbExecuteRS(db, sql)" & vbCrLf
        str = str & vbCrLf
        str = str & "    if not rs.eof then" & vbCrLf
        str = str & "Response.Write " & Chr(34) & "<select name=pendingAdmin style=" & Chr(34) & " & chr(34) & " & Chr(34) & "WIDTH: 250" & Chr(34) & " & chr(34) & " & Chr(34) & " size=1>" & Chr(34) & vbCrLf
        str = str & vbCrLf
        str = str & "        do while not rs.EOF" & vbCrLf
        str = str & "            Response.Write " & Chr(34) & "<option value=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & ">" & Chr(34) & " & rs(" & Chr(34) & Me.loginFld & Chr(34) & ")" & vbCrLf
        str = str & "            rs.movenext()" & vbCrLf
        str = str & "        loop" & vbCrLf
        str = str & vbCrLf
        str = str & "        Response.Write " & Chr(34) & "</select>" & Chr(34) & vbCrLf
        str = str & "    else" & vbCrLf
        str = str & "        Response.Write " & Chr(34) & "There is no pending admin." & Chr(34) & vbCrLf
        str = str & "    end if" & vbCrLf
        str = str & vbCrLf
        str = str & "    call dbRSClose(db, rs)" & vbCrLf
        str = str & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & "' Get a list of admin members." & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & vbCrLf
        str = str & "sub getApprovedAdmin()" & vbCrLf
        str = str & "    dim db, rs, sql" & vbCrLf
        str = str & vbCrLf
        str = str & "    sql = " & Chr(34) & "SELECT * FROM " & Me.DBTableName & " where approval = 1 and isAdmin = 1" & Chr(34) & vbCrLf
        str = str & "    set db = dbConnect()" & vbCrLf
        str = str & "    set rs = dbExecuteRS(db, sql)" & vbCrLf
        str = str & vbCrLf
        str = str & "    if not rs.eof then" & vbCrLf
        str = str & "        Response.Write " & Chr(34) & "<select name=approvedAdmin style=" & Chr(34) & " & chr(34) & " & Chr(34) & "WIDTH: 250" & Chr(34) & " & chr(34) & " & Chr(34) & " size=1>" & Chr(34) & vbCrLf
        str = str & vbCrLf
        str = str & "        do while not rs.EOF" & vbCrLf
        str = str & "            Response.Write " & Chr(34) & "<option value=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & ">" & Chr(34) & " & rs(" & Chr(34) & Me.loginFld & Chr(34) & ")" & vbCrLf
        str = str & "            rs.movenext()" & vbCrLf
        str = str & "        loop" & vbCrLf
        str = str & vbCrLf
        str = str & "        Response.Write " & Chr(34) & "</select>" & Chr(34) & vbCrLf
        str = str & "    else" & vbCrLf
        str = str & "        Response.Write " & Chr(34) & "There is no approved admin." & Chr(34) & vbCrLf
        str = str & "    end if" & vbCrLf
        str = str & vbCrLf
        str = str & "    call dbRSClose(db, rs)" & vbCrLf
        str = str & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf
        str = str & "%>" & vbCrLf
        str = str & vbCrLf

        Return str

    End Function


    ''''''''''''''''''''''''''''''''''''''''''''
    ' Regitration include/regFuncs.asp
    ''''''''''''''''''''''''''''''''''''''''''''

    Public Function getIncFuncsFile() As String
        Dim sessionArray() As String = Me.clsMD.sessionArray
        Dim dbArray() As String = Me.clsMD.dbArray
        Dim formArray() As String = Me.clsMD.formArray
        Dim str As String = ""
        Dim vbIndent As String = "    "
        Dim i As Integer

        str = str & "<!--#include file=" & Chr(34) & "../" & Me.websiteFuncFilePath & Chr(34) & "-->" & vbCrLf

        str = str & vbCrLf
        str = str & "<%" & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function saveSession()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Saves session variables.")
        str = str & "sub saveSession()" & vbCrLf
        For i = 0 To UBound(sessionArray)
            str = str & vbIndent & "session(" & Chr(34) & sessionArray(i) & Chr(34) & ") = trim(request(" & Chr(34) & formArray(i) & Chr(34) & "))" & vbCrLf
        Next
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function clearSession()	
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Clears session variables.")
        str = str & "sub clearSession()" & vbCrLf
        For i = 0 To UBound(sessionArray)
            str = str & vbIndent & "session(" & Chr(34) & sessionArray(i) & Chr(34) & ") = " & Chr(34) & Chr(34) & vbCrLf
        Next
        'str = str & vbIndent & "session(" & Chr(34) & "regDate" & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function insertData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Insert data into database table " & Me.DBTableName & ".")
        str = str & "sub insertData()" & vbCrLf
        str = str & vbIndent & "dim db, sql, rs" & vbCrLf
        str = str & vbCrLf

        str = str & vbIndent & "sql = " & Chr(34) & "INSERT INTO " & Me.DBTableName & " (" & Chr(34) & vbCrLf
        str = str & vbIndent & "sql = sql & " & Chr(34) & "regDate, " & Chr(34) & vbCrLf
        str = str & vbIndent & "sql = sql & " & Chr(34) & "isAdmin, " & Chr(34) & vbCrLf
        str = str & vbIndent & "sql = sql & " & Chr(34) & "approval, " & Chr(34) & vbCrLf
        For i = 0 To UBound(dbArray) - 1
            str = str & vbIndent & "sql = sql & " & Chr(34) & dbArray(i) & ", " & Chr(34) & vbCrLf
        Next
        str = str & vbIndent & "sql = sql & " & Chr(34) & dbArray(i) & Chr(34) & vbCrLf

        str = str & vbCrLf
        str = str & vbIndent & "sql = sql & " & Chr(34) & ") VALUES (" & Chr(34) & vbCrLf
        str = str & vbCrLf

        str = str & vbIndent & "sql = sql & sqlEncode(date()) & " & Chr(34) & ", " & Chr(34) & vbCrLf
        ''''''''''''''''''''''''' insert isAdmin
        str = str & vbIndent & "sql = sql & " & Chr(34) & "0" & Chr(34) & " & " & Chr(34) & ", " & Chr(34) & vbCrLf
        ''''''''''''''''''''''''' insert approval
        If Me.loginNeedsAdminApproval Then
            str = str & vbIndent & "sql = sql & " & Chr(34) & "0" & Chr(34) & " & " & Chr(34) & ", " & Chr(34) & vbCrLf
        Else
            str = str & vbIndent & "sql = sql & " & Chr(34) & "1" & Chr(34) & " & " & Chr(34) & ", " & Chr(34) & vbCrLf
        End If
        'str = str &  vbIndent & "sql = sql & 0 & " & chr(34) & ", " & Chr(34) & vbCrLf
        For i = 0 To UBound(dbArray) - 1
            str = str & vbIndent & "sql = sql & sqlEncode(session(" & Chr(34) & sessionArray(i) & Chr(34) & "))" & " & " & Chr(34) & ", " & Chr(34) & vbCrLf
        Next
        str = str & vbIndent & "sql = sql & sqlEncode(session(" & Chr(34) & sessionArray(i) & Chr(34) & "))" & " & " & Chr(34) & ") " & Chr(34) & vbCrLf

        str = str & vbCrLf
        str = str & vbIndent & "set db = dbConnect()" & vbCrLf
        str = str & vbIndent & "call dbExecute(db, sql)" & vbCrLf
        str = str & vbCrLf
        str = str & vbIndent & "call db.close()" & vbCrLf
        str = str & vbIndent & "set db = nothing" & vbCrLf

        str = str & "end sub" & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function updateData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Update data in database table " & Me.DBTableName & ".")
        str = str & "sub updateData(id)" & vbCrLf
        str = str & vbIndent & "Dim db, sql" & vbCrLf
        str = str & vbCrLf

        str = str & vbIndent & "sql = " & Chr(34) & "UPDATE " & Me.DBTableName & " SET " & Chr(34) & vbCrLf
        For i = 0 To UBound(dbArray) - 1
            str = str & vbIndent & "sql = sql & " & Chr(34) & dbArray(i) & " = " & Chr(34) & " & sqlEncode(session(" & Chr(34) & sessionArray(i) & Chr(34) & ")) & " & Chr(34) & ", " & Chr(34) & vbCrLf
        Next
        str = str & vbIndent & "sql = sql & " & Chr(34) & dbArray(i) & " = " & Chr(34) & " & sqlEncode(session(" & Chr(34) & sessionArray(i) & Chr(34) & "))" & vbCrLf
        str = str & vbIndent & "sql = sql & " & Chr(34) & " WHERE ID = " & Chr(34) & " & id " & vbCrLf

        str = str & vbCrLf
        str = str & vbIndent & "set db = dbConnect()" & vbCrLf
        str = str & vbIndent & "call dbExecute(db, sql)" & vbCrLf
        str = str & vbCrLf
        str = str & vbIndent & "call db.close()" & vbCrLf
        str = str & vbIndent & "set db = nothing" & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function deleteData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Delete data in database table " & Me.DBTableName & ".")
        str = str & "sub deleteData(id)" & vbCrLf
        str = str & vbIndent & "Dim db, sql" & vbCrLf

        str = str & vbIndent & "sql = " & Chr(34) & "DELETE FROM " & Me.DBTableName & " WHERE ID = " & Chr(34) & " & id" & vbCrLf
        str = str & vbCrLf

        str = str & vbIndent & "set db = dbConnect()" & vbCrLf
        str = str & vbIndent & "call dbExecute(db, sql)" & vbCrLf
        str = str & vbCrLf
        str = str & vbIndent & "call db.close()" & vbCrLf
        str = str & vbIndent & "set db = nothing" & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf & vbCrLf


        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function retrieveData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Retrieve data from database table " & Me.DBTableName & ".")
        str = str & "sub retrieveData(id)" & vbCrLf
        str = str & vbIndent & "Dim db, sql, sr, count" & vbCrLf

        str = str & vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & Me.DBTableName & " WHERE ID = " & Chr(34) & " & id" & vbCrLf
        str = str & vbCrLf
        str = str & vbIndent & "set db = dbConnect()" & vbCrLf
        str = str & vbIndent & "set rs = dbExecuteRS(db, sql)" & vbCrLf
        str = str & vbCrLf
        str = str & vbIndent & "session(" & Chr(34) & "regDate" & Chr(34) & ") = rs(" & Chr(34) & "regDate" & Chr(34) & ")" & vbCrLf

        For i = 0 To UBound(sessionArray)
            str = str & vbIndent & "session(" & Chr(34) & sessionArray(i) & Chr(34) & ") = trim(rs(" & Chr(34) & dbArray(i) & Chr(34) & "))" & vbCrLf
        Next

        str = str & vbCrLf
        str = str & vbIndent & "rs.close()" & vbCrLf
        str = str & vbIndent & "set rs = nothing" & vbCrLf
        str = str & vbCrLf
        str = str & vbIndent & "call db.close()" & vbCrLf
        str = str & vbIndent & "set db = nothing" & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function approveData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Approve registration in database table " & Me.DBTableName & ".")
        str = str & "sub approveData(id)" & vbCrLf
        str = str & vbIndent & "Dim db, sql" & vbCrLf

        str = str & vbIndent & "sql = " & Chr(34) & "UPDATE " & Me.DBTableName & " SET approval = 1 WHERE ID = " & Chr(34) & " & id" & vbCrLf

        str = str & vbCrLf
        str = str & vbIndent & "set db = dbConnect()" & vbCrLf
        str = str & vbIndent & "call dbExecute(db, sql)" & vbCrLf
        str = str & vbCrLf
        str = str & vbIndent & "call db.close()" & vbCrLf
        str = str & vbIndent & "set db = nothing" & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function disapproveData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Disapprove registration in database table " & Me.DBTableName & ".")
        str = str & "sub disApproveData(id)" & vbCrLf
        str = str & vbIndent & "Dim db, sql" & vbCrLf

        str = str & vbIndent & "sql = " & Chr(34) & "UPDATE " & Me.DBTableName & " SET approval = 0 WHERE ID = " & Chr(34) & " & id" & vbCrLf

        str = str & vbCrLf
        str = str & vbIndent & "set db = dbConnect()" & vbCrLf
        str = str & vbIndent & "call dbExecute(db, sql)" & vbCrLf
        str = str & vbCrLf
        str = str & vbIndent & "call db.close()" & vbCrLf
        str = str & vbIndent & "set db = nothing" & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf & vbCrLf


        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function getPendingList()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Get pending registration in database table " & Me.DBTableName & ".")
        str = str & "sub getPendingList()" & vbCrLf
        str = str & vbIndent & "Dim db, sql, count" & vbCrLf

        str = str & vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & Me.DBTableName & " WHERE approval = 0" & Chr(34) & vbCrLf

        str = str & vbCrLf
        str = str & vbIndent & "set db = dbConnect()" & vbCrLf
        str = str & vbIndent & "set rs =  dbExecuteRS(db, sql)" & vbCrLf
        str = str & vbCrLf

        str = str & vbIndent & "if rs.recordcount > 0 then" & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<H3>List of Pending Registrants</H3>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<table width=600 border=1>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & "If rs.recordcount = 1 then" & vbCrLf
        str = str & vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td colspan=3>There is 1 pending registrant.</td></tr>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & "Else ' recordCount > 1" & vbCrLf
        str = str & vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td colspan=3>There are " & Chr(34) & " & rs.recordcount & " & Chr(34) & " pending registrants.</td></tr>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & "End If" & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td width=150><b>Register Date</b></td><td width=150><b>Name</b></td><td><b>Detail Information</b></td></tr>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & "do while not rs.eof" & vbCrLf
        str = str & vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>" & Chr(34) & " & rs(" & Chr(34) & Me.loginFld & Chr(34) & ") & " & Chr(34) & ", " & Chr(34) & " & rs(" & Chr(34) & Me.loginFld & Chr(34) & ") & " & Chr(34) & "</td><td><a href=" & Chr(39) & Me.admRegConfirmFile & "?src=pending&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & vbIndent & "rs.MoveNext()" & vbCrLf
        str = str & vbIndent & vbIndent & "loop" & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Write " & Chr(34) & "</table>" & Chr(34) & vbCrLf
        str = str & vbIndent & "else" & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<h3>Currently there is no pending registrant.</h3>" & Chr(34) & vbCrLf
        str = str & vbIndent & "end if" & vbCrLf

        str = str & vbCrLf
        str = str & vbIndent & "call dbRSclose(db, rs)" & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function getApprovedList()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Get approved registration in database table " & Me.DBTableName & ".")
        str = str & "sub getApprovedList()" & vbCrLf
        str = str & vbIndent & "Dim db, sql, count" & vbCrLf

        str = str & vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & Me.DBTableName & " WHERE approval = 1" & Chr(34) & vbCrLf

        str = str & vbCrLf
        str = str & vbIndent & "set db = dbConnect()" & vbCrLf
        str = str & vbIndent & "set rs =  dbExecuteRS(db, sql)" & vbCrLf
        str = str & vbCrLf

        str = str & vbIndent & "if rs.recordcount > 0 then" & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<H3>List of Approved Registrants</H3>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<table width=600 border=1>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & "If rs.recordcount = 1 then" & vbCrLf
        str = str & vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td colspan=3>There is 1 approved registrant.</td></tr>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & "Else ' recordCount > 1" & vbCrLf
        str = str & vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td colspan=3>There are " & Chr(34) & " & rs.recordcount & " & Chr(34) & " approved registrants.</td></tr>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & "End If" & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td width=150><b>Register Date</b></td><td width=150><b>Name</b></td><td><b>Detail Information</b></td></tr>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & "do while not rs.eof" & vbCrLf
        str = str & vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>" & Chr(34) & " & rs(" & Chr(34) & Me.loginFld & Chr(34) & ") & " & Chr(34) & ", " & Chr(34) & " & rs(" & Chr(34) & Me.loginFld & Chr(34) & ") & " & Chr(34) & "</td><td><a href=" & Chr(39) & Me.admRegConfirmFile & "?src=approved&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34) & vbCrLf
        str = str & vbIndent & vbIndent & vbIndent & "rs.MoveNext()" & vbCrLf
        str = str & vbIndent & vbIndent & "loop" & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Write " & Chr(34) & "</table>" & Chr(34) & vbCrLf
        str = str & vbIndent & "else" & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<h3>Currently there is no approved registrant.</h3>" & Chr(34) & vbCrLf
        str = str & vbIndent & "end if" & vbCrLf
        str = str & vbCrLf
        str = str & vbIndent & "call dbRSclose(db, rs)" & vbCrLf
        str = str & "end sub" & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function checkRegError() and CheckError()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Check error in form data.")
        str = str & vbCrLf

        str = str & "' Used for new member registration" & vbCrLf
        str = str & "function checkRegError()" & vbCrLf
        str = str & "    checkRegError = " & Chr(34) & "" & Chr(34) & vbCrLf
        str = str & "    if session(" & Chr(34) & Me.loginFld & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & "		checkRegError = checkRegError & " & Chr(34) & "<li>Account name cannot be empty." & Chr(34) & vbCrLf
        str = str & "    end if" & vbCrLf
        str = str & "    checkRegError = checkRegError & checkPasswordError(session(" & Chr(34) & Me.passwordFld & Chr(34) & "), Trim(request(" & Chr(34) & Me.passwordRptFldHTMLName & Chr(34) & ")))" & vbCrLf
        str = str & "    if fieldExist(" & Chr(34) & Me.DBTableName & Chr(34) & ", " & Chr(34) & Me.loginFld & Chr(34) & ", session(" & Chr(34) & Me.loginFld & Chr(34) & ")) then" & vbCrLf
        str = str & "	    checkRegError = checkRegError & " & Chr(34) & "<li>Account '" & Chr(34) & " & session(" & Chr(34) & Me.loginFld & Chr(34) & ") & " & Chr(34) & "' aready exists. Please choose another account name." & Chr(34) & vbCrLf
        str = str & "    end if" & vbCrLf
        str = str & "end function" & vbCrLf
        str = str & vbCrLf & vbCrLf


        str = str & "' Check error in password." & vbCrLf
        str = str & "Function checkPasswordError(password, passwordRpt)" & vbCrLf
        str = str & "	 Dim errMsg" & vbCrLf
        str = str & "    errMsg = " & Chr(34) & "" & Chr(34) & vbCrLf
        str = str & "    If password = " & Chr(34) & "" & Chr(34) & " Then" & vbCrLf
        str = str & "		errMsg = errMsg & " & Chr(34) & "<li>Password cannot be empty." & Chr(34) & vbCrLf
        str = str & "    Elseif len(password) < " & Me.passwordMinLength & " Then" & vbCrLf
        str = str & "		errMsg = errMsg & " & Chr(34) & "<li>Password should be at least " & Me.passwordMinLength & " characters in length." & Chr(34) & vbCrLf
        str = str & "    End if" & vbCrLf
        If Me.passwordIsCaseSensitive Then
            str = str & "    If NOT caseSensitiveEqual(password, passwordRpt) Then" & vbCrLf
        Else
            str = str & "    If password <> passwordRpt Then" & vbCrLf
        End If
        str = str & "		errMsg = errMsg & " & Chr(34) & "<li>The two passwords should match each other." & Chr(34) & vbCrLf
        str = str & "    End if" & vbCrLf
        str = str & "    checkPasswordError = errMsg" & vbCrLf
        str = str & "End Function" & vbCrLf
        str = str & vbCrLf & vbCrLf


        str = str & "' Check error in member or administrator edit." & vbCrLf
        str = str & "function checkError()" & vbCrLf
        str = str & vbIndent & "checkError = " & Chr(34) & Chr(34) & vbCrLf
        str = str & "end function" & vbCrLf
        str = str & vbCrLf
        str = str & vbCrLf

        str = str & vbCrLf & vbCrLf
        str = str & "%" & ">" & vbCrLf

        Return str
    End Function

End Class
