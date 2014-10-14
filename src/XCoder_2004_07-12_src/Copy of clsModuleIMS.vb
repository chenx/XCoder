Imports XinCoder.xcUtil

Public Class clsModuleIMS

    Private moduleName As String

    ' This class provides methods to process metadata.
    ' In short, from metadata we get:
    '   1) edit form and confirm form
    '   2) relevant functions in regFuncs
    '   3) sql file
    '   4) xml file
    Private clsMD As New clsMetadata()
    Private dbFieldLenArray() As String
    Private dbArray() As String
    Private sessionArray() As String
    Private formArray() As String


    Private websiteHomeFilePath As String
    Private websiteFuncFilePath As String = "../asp-bin/aspFuncLib.asp"
    Private loginAdmSecFilePath As String = "../login/admin_security.asp"
    Private loginMemSecFilePath As String = "../login/mem_security.asp"
    Private headerFilePath As String ' can be site header or module header
    Private footerFilePath As String ' can be site footer or module footer
    Private admHomeFilePath As String
    Private memHomeFilePath As String

    Private DataBaseTable1Name As String  ' Used by sql file.
    Private DataBaseTable2Name As String  ' Used by sql file.

    ' File names
    Private regEditFile As String
    Private regConfirmFile As String
    Private memListFile As String
    Private memEditFile As String
    Private memConfirmFile As String
    Private admListFile As String
    Private admEditFile As String
    Private admConfirmFile As String
    Private admFinalEditFile As String
    Private admFinalConfirmFile As String

    Private incFolder As String
    Private incEditTableFile As String
    Private incConfirmTableFile As String
    Private incFuncFile As String

    ' Options
    Private IMSType As Integer
    Private UseAccessorySearch As Boolean
    Private UseAccessoryCalendar As Boolean
    Private UseAccessoryAccess As Boolean
    Private UseAccessoryWord As Boolean

    Private UserRightNonMemView As Boolean
    Private UserRightMemView As Boolean
    Private UserRightMemEdit As Boolean
    Private UserRightMemApprove As Boolean
    Private UserRightMemDisapprove As Boolean
    Private UserRightMemDelete As Boolean
    Private UserRightAdmView As Boolean
    Private UserRightAdmEdit As Boolean
    Private UserRightAdmApprove As Boolean
    Private UserRightAdmDisapprove As Boolean
    Private UserRightAdmDelete As Boolean


    Public Sub init(ByRef metadata As String, ByVal moduleName As String, _
        ByVal DBTable1Name As String, ByVal DBTable2Name As String, _
        ByVal globalHomeFilePath As String, ByVal globalFuncFilePath As String, _
        ByVal loginAdmSecurityFilePath As String, ByVal loginMemSecurityFilePath As String, _
        ByVal headerFilePath As String, ByVal footerFilePath As String, _
        ByVal adminHomeFilePath As String, ByVal memberHomeFilePath As String, _
        ByVal regEditFileName As String, ByVal regConfirmFileName As String, _
        ByVal memListFileName As String, ByVal memEditFileName As String, ByVal memConfirmFileName As String, _
        ByVal admListFileName As String, ByVal admEditFileName As String, ByVal admConfirmFileName As String, _
        ByVal admFinalEditFileName As String, ByVal admFinalConfirmFileName As String, _
        ByVal incFolderName As String, ByVal incFuncFileName As String, _
        ByVal incEditTableFileName As String, ByVal incConfirmTableFileName As String, _
        ByVal IMSTypeNum As Integer, _
        ByVal useSearchAccessory As Boolean, ByVal useCalendarAccessory As Boolean, _
        ByVal useAccessAccessory As Boolean, ByVal useWordAccessory As Boolean, _
        ByVal allowNonMemView As Boolean, _
        ByVal allowMemView As Boolean, ByVal allowMemEdit As Boolean, _
        ByVal allowMemApprove As Boolean, ByVal allowMemDisapprove As Boolean, _
        ByVal allowMemDelete As Boolean, _
        ByVal allowAdminView As Boolean, ByVal allowAdminEdit As Boolean, _
        ByVal allowAdminApprove As Boolean, ByVal allowAdminDisapprove As Boolean, _
        ByVal allowAdminDelete As Boolean)

        Me.clsMD.init(metadata)
        Me.dbFieldLenArray = clsMD.dbFieldLenArray
        Me.dbArray = clsMD.dbArray
        Me.sessionArray = clsMD.sessionArray
        Me.formArray = clsMD.formArray

        Me.moduleName = moduleName

        Me.DataBaseTable1Name = DBTable1Name  ' Used by sql file.
        Me.DataBaseTable2Name = DBTable2Name  ' Used by sql file.

        Me.websiteHomeFilePath = globalHomeFilePath
        Me.websiteFuncFilePath = globalFuncFilePath
        Me.loginAdmSecFilePath = loginAdmSecurityFilePath
        Me.loginMemSecFilePath = loginMemSecurityFilePath
        Me.headerFilePath = headerFilePath
        Me.footerFilePath = footerFilePath
        Me.admHomeFilePath = adminHomeFilePath
        Me.memHomeFilePath = memberHomeFilePath

        ' File names
        Me.regEditFile = regEditFileName
        Me.regConfirmFile = regConfirmFileName
        Me.memListFile = memListFileName
        Me.memEditFile = memEditFileName
        Me.memConfirmFile = memConfirmFileName
        Me.admListFile = admListFileName
        Me.admEditFile = admEditFileName
        Me.admConfirmFile = admConfirmFileName
        Me.admFinalEditFile = admFinalEditFileName
        Me.admFinalConfirmFile = admFinalConfirmFileName

        Me.incFolder = incFolderName
        Me.incEditTableFile = incEditTableFileName
        Me.incConfirmTableFile = incConfirmTableFileName
        Me.incFuncFile = incFuncFileName

        ' options
        IMSType = IMSTypeNum
        UseAccessorySearch = useSearchAccessory
        UseAccessoryCalendar = useCalendarAccessory
        UseAccessoryAccess = useAccessAccessory
        UseAccessoryWord = useWordAccessory

        UserRightNonMemView = allowNonMemView
        UserRightMemView = allowMemView
        UserRightMemEdit = allowMemEdit
        UserRightMemApprove = allowMemApprove
        UserRightMemDisapprove = allowMemDisapprove
        UserRightMemDelete = allowMemDelete
        UserRightAdmView = allowAdminView
        UserRightAdmEdit = allowAdminEdit
        UserRightAdmApprove = allowAdminApprove
        UserRightAdmDisapprove = allowAdminDisapprove
        UserRightAdmDelete = allowAdminDelete

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' XML system functions
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function generateXMLFile() As String
        Return Me.clsMD.generateXML
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SQL system functions
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Function generateDBTableSQL(ByVal tableName As String, _
        ByVal imsType As Integer, Optional ByVal tableType As String = "") As String
        ' Given dbFieldLenArray, dbArray
        Dim count As Integer
        Dim i As Integer
        Dim hasTextField As Boolean = False

        Dim str As String = ""

        str = str & "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & DataBaseTable1Name & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        str = str & "drop table [dbo].[" & tableName & "]" & vbCrLf
        str = str & "GO" & vbCrLf

        str = str & vbCrLf & vbCrLf

        str = str & "CREATE TABLE [dbo].[" & tableName & "] (" & vbCrLf
        str = str & sqlIndent & "[ID] [int] IDENTITY (1, 1) NOT NULL ," & vbCrLf
        If imsType = 1 Then
            str = str & sqlIndent & "[memID] [int] NULL ," & vbCrLf
        ElseIf imsType = 2 Then
            If LCase(tableType) = "final" Then
                ' Outside key from the first table of 'imsType = 2'
                str = str & sqlIndent & "[regID] [int] NULL ," & vbCrLf
            End If
            str = str & sqlIndent & "[memID] [int] NULL ," & vbCrLf
        End If
        str = str & sqlIndent & "[regDate] [datetime] NULL ," & vbCrLf
        str = str & sqlIndent & "[approval] [varchar] (1) NULL ," & vbCrLf

        count = UBound(dbFieldLenArray) - 1
        For i = 0 To count
            If UCase(dbFieldLenArray(i)) = "TEXT" Then
                str = str & sqlIndent & "[" & dbArray(i) & "] [text] NULL ," & vbCrLf
                hasTextField = True
            Else
                str = str & sqlIndent & "[" & dbArray(i) & "] [varchar] (" & CInt(dbFieldLenArray(i)) & ") NULL ," & vbCrLf
            End If
        Next

        If UCase(dbFieldLenArray(i)) = "TEXT" Then
            str = str & sqlIndent & "[" & dbArray(i) & "] [text] NULL" & vbCrLf
            hasTextField = True
        Else
            Try
                str = str & sqlIndent & "[" & dbArray(i) & "] [varchar] (" & CInt(dbFieldLenArray(i)) & ") NULL" & vbCrLf
            Catch ex As Exception
                MsgBox("clsModuleIMS.generateTable1SQL error: " & ex.Message)
            End Try
        End If

        If hasTextField = True Then
            str = str & ") ON [PRIMARY] TEXTIMAGE_ON  [PRIMARY]" & vbCrLf
        Else
            str = str & ") ON [PRIMARY]" & vbCrLf
        End If
        str = str & "GO" & vbCrLf

        str = str & vbCrLf
        str = str & "ALTER TABLE [dbo].[" & tableName & "] WITH NOCHECK ADD " & vbCrLf
        str = str & sqlIndent & "CONSTRAINT [PK_" & tableName & "] PRIMARY KEY  NONCLUSTERED " & vbCrLf
        str = str & sqlIndent & "(" & vbCrLf
        str = str & sqlIndent & sqlIndent & "[ID]" & vbCrLf
        str = str & sqlIndent & ") ON [PRIMARY]" & vbCrLf
        str = str & "GO" & vbCrLf

        Return str
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' From now on are functions to generate files.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Function generateRegEditFile() As String
        Dim str As String = ""
        str = str & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incFuncFile & Chr(34) & "-->")) & vbCrLf
        str = str & vbCrLf & vbCrLf
        str = str & (("<" & "%")) & vbCrLf
        str = str & (("Dim err")) & vbCrLf
        str = str & (("err = " & Chr(34) & Chr(34))) & vbCrLf
        str = str & vbCrLf
        str = str & ("if request(" & Chr(34) & "btnSubmit" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then" & vbCrLf)
        str = str & (vbIndent & "call saveSession()" & vbCrLf)
        str = str & (vbIndent & "err = checkError()" & vbCrLf)
        str = str & (vbIndent & "if err = " & Chr(34) & Chr(34) & " then" & vbCrLf)
        str = str & vbIndent & vbIndent & "if request(" & Chr(34) & "redirectURL" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & vbIndent & vbIndent & vbIndent & "Response.Redirect " & Chr(34) & Me.regConfirmFile & Chr(34) & " & " & Chr(34) & "?redirectURL=" & Chr(34) & " & Request(" & Chr(34) & "redirectURL" & Chr(34) & ")" & vbCrLf
        str = str & vbIndent & vbIndent & "else" & vbCrLf
        str = str & vbIndent & (vbIndent & vbIndent & "Response.Redirect " & Chr(34) & Me.regConfirmFile & Chr(34) & vbCrLf)
        str = str & vbIndent & vbIndent & "end if" & vbCrLf
        str = str & (vbIndent & "end if" & vbCrLf)
        str = str & ("elseif request(" & Chr(34) & "btnCancel" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then" & vbCrLf)
        str = str & (vbIndent & "call clearSession()" & vbCrLf)
        str = str & vbIndent & "if request(" & Chr(34) & "redirectURL" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
        str = str & vbIndent & vbIndent & "Response.Redirect request(" & Chr(34) & "redirectURL" & Chr(34) & ")" & vbCrLf
        str = str & vbIndent & "else" & vbCrLf
        str = str & vbIndent & (vbIndent & "Response.Redirect " & Chr(34) & Me.websiteHomeFilePath & Chr(34) & vbCrLf)
        str = str & vbIndent & "end if" & vbCrLf
        str = str & ("end if" & vbCrLf)
        str = str & (("%" & ">")) & vbCrLf

        str = str & vbCrLf
        str = str & (("<!--#include file=" & Chr(34) & Me.headerFilePath & Chr(34) & "-->")) & vbCrLf

        str = str & vbCrLf
        str = str & (("<!--enter the specific registration information here-->")) & vbCrLf
        str = str & (("<h3>Registration Information</h3>")) & vbCrLf        str = str & (("<!--end the specific registration information-->")) & vbCrLf

        str = str & vbCrLf
        str = str & (("<" & "%")) & vbCrLf
        str = str & ("if err <> " & Chr(34) & Chr(34) & " then" & vbCrLf)
        str = str & (vbIndent & ("Response.Write " & Chr(34) & "<font color=red>Registration Error: <br>" & Chr(34) & " & err & " & Chr(34) & "</font>" & Chr(34))) & vbCrLf
        str = str & ("end if" & vbCrLf)
        str = str & (("%" & ">"))

        str = str & vbCrLf
        str = str & (("<form method=post name=RegisterForm id=RegisterForm>")) & vbCrLf
        str = str & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incEditTableFile & Chr(34) & "-->")) & vbCrLf
        str = str & vbCrLf
        str = str & (("<br>")) & vbCrLf
        str = str & "<INPUT TYpe=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "redirectURL" & Chr(34) & " value=" & Chr(34) & "<%=Request(" & Chr(34) & "redirectURL" & Chr(34) & ")%>" & Chr(34) & ">" & vbCrLf
        str = str & (("<table><tr><td align=left>")) & vbCrLf
        str = str & (("<input type=submit name=btnSubmit value=submit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">")) & vbCrLf
        str = str & (("<input type=submit name=btnCancel value=cancel style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">")) & vbCrLf        str = str & (("</td></tr></table>")) & vbCrLf
        str = str & (("</form>")) & vbCrLf
        str = str & vbCrLf
        str = str & (("<!--#include file=" & Chr(34) & Me.footerFilePath & Chr(34) & "-->")) & vbCrLf

        Return str
    End Function


    Public Function generateRegConfirmFile() As String
        Dim str As String = ""

        str = str & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incFuncFile & Chr(34) & "-->")) & vbCrLf
        str = str & vbCrLf & vbCrLf
        str = str & (("<" & "%")) & vbCrLf
        str = str & ("if request(" & Chr(34) & "btnEdit" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then" & vbCrLf)
        str = str & (vbIndent & "Response.Redirect " & Chr(34) & Me.regEditFile & Chr(34) & vbCrLf)
        str = str & ("end if" & vbCrLf)
        str = str & (("%" & ">")) & vbCrLf

        str = str & vbCrLf
        str = str & (("<!--#include file=" & Chr(34) & Me.headerFilePath & Chr(34) & "-->")) & vbCrLf
        str = str & vbCrLf
        str = str & (("<" & "% if request(" & Chr(34) & "btnSubmit" & Chr(34) & ") = " & Chr(34) & Chr(34) & " then%" & ">")) & vbCrLf
        str = str & (("<" & "%else%" & ">")) & vbCrLf        str = str & (("<!--enter the specific confirmation information here-->")) & vbCrLf
        str = str & (("<p><b><font color=red>Thank you for submitting the registration.")) & vbCrLf        str = str & (("The following is your registration information.<br> Please print a copy for your record.</font></b></p>")) & vbCrLf
        str = str & vbIndent & "<% if request(" & Chr(34) & "redirectURL" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " then %>" & vbCrLf
        str = str & vbIndent & vbIndent & "<p><a href=" & Chr(34) & "<%=Request(" & Chr(34) & "redirectURL" & Chr(34) & ")%>" & Chr(34) & ">Back</a></p>" & vbCrLf
        str = str & vbIndent & "<% else %>" & vbCrLf
        str = str & vbIndent & vbIndent & (("<p><a href='" & Me.websiteHomeFilePath & "'>Back to Homepage</a></p>")) & vbCrLf
        str = str & vbIndent & "<% end if %>" & vbCrLf
        str = str & (("<!--end the specific confirmation information-->")) & vbCrLf
        str = str & (("<" & "%end if%" & ">")) & vbCrLf
        str = str & vbCrLf

        str = str & (("<!--enter the specific registration information here-->")) & vbCrLf
        str = str & (("<h3>Registration Information</h3>")) & vbCrLf        str = str & (("<!--end the specific registration information-->")) & vbCrLf

        str = str & vbCrLf
        str = str & (("<form  method=post name=RegisterConfirmForm id=RegisterConfirmForm>")) & vbCrLf
        str = str & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incConfirmTableFile & Chr(34) & "-->")) & vbCrLf
        str = str & (("<br>")) & vbCrLf

        str = str & vbCrLf
        str = str & (("<" & "% if request(" & Chr(34) & "btnSubmit" & Chr(34) & ") = " & Chr(34) & Chr(34) & " then%" & ">")) & vbCrLf
        str = str & "<INPUT Type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "redirectURL" & Chr(34) & " value=" & Chr(34) & "<%=Request(" & Chr(34) & "redirectURL" & Chr(34) & ")%>" & Chr(34) & ">" & vbCrLf
        str = str & (("<table><tr><td align=left>")) & vbCrLf
        str = str & (("<input type=submit name=btnEdit value=Edit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">")) & vbCrLf        str = str & (("<input type=submit name=btnSubmit value=submit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">")) & vbCrLf
        str = str & (("</td></tr></table>")) & vbCrLf
        str = str & (("<" & "%else")) & vbCrLf        str = str & (vbIndent & ("insertData()")) & vbCrLf
        str = str & (vbIndent & ("clearSession()")) & vbCrLf
        str = str & (("end if%" & ">")) & vbCrLf
        str = str & vbCrLf
        str = str & (("</form>")) & vbCrLf        str = str & vbCrLf
        str = str & (("<!--#include file=" & Chr(34) & Me.footerFilePath & Chr(34) & "-->")) & vbCrLf

        Return str
    End Function


    Public Function generateMemListFile() As String
        Dim str As String = ""

        str = str & (("<!--#include file=" & Chr(34) & Me.loginMemSecFilePath & Chr(34) & "-->")) & vbCrLf
        str = str & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incFuncFile & Chr(34) & "-->")) & vbCrLf

        str = str & vbCrLf & vbCrLf

        str = str & (("<!--#include file=" & Chr(34) & Me.headerFilePath & Chr(34) & "-->")) & vbCrLf

        str = str & vbCrLf
        str = str & (("<center>")) & vbCrLf

        str = str & vbCrLf
        str = str & (("<!--enter the specific registration information here-->")) & vbCrLf
        str = str & (("<h2>Registration Information for " & DataBaseTable1Name & " </h2>")) & vbCrLf
        str = str & (("<a href='" & Me.memHomeFilePath & "'>Back to Member Home</a>")) & vbCrLf        str = str & (("<!--end the specific registration information-->")) & vbCrLf
        str = str & vbCrLf
        str = str & "<p><a href=" & Chr(34) & Me.regEditFile & "?redirectURL=memRegList.asp" & Chr(34) & "><b>New " & Me.moduleName & " Registration</b></a></p>" & vbCrLf
        str = str & vbCrLf
        str = str & (("<" & "% call getPendingListMem() %" & ">")) & vbCrLf
        str = str & (("<" & "% call getApprovedListMem() %" & ">")) & vbCrLf
        str = str & vbCrLf

        str = str & (("<P align=center>&nbsp;</P>")) & vbCrLf
        str = str & (("</center>")) & vbCrLf
        str = str & vbCrLf
        str = str & (("<!--#include file=" & Chr(34) & Me.footerFilePath & Chr(34) & "-->")) & vbCrLf

        Return str
    End Function


    Public Function generateMemEditFile() As String
        Dim str As String = ""

        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.loginMemSecFilePath & Chr(34) & "-->"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incFuncFile & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf & vbCrLf

        str = str & vbCrLf & (("<" & "%"))
        str = str & vbCrLf & (("Dim ID, SRC, err"))
        str = str & vbCrLf & (("err = " & Chr(34) & Chr(34)))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("ID = Request.QueryString(" & Chr(34) & "id" & Chr(34) & ")"))
        str = str & vbCrLf & (("SRC = Request.QueryString(" & Chr(34) & "src" & Chr(34) & ")"))
        str = str & vbCrLf & vbCrLf
        str = str & "If getRegMemID(ID) <> session(" & Chr(34) & "_isMem" & Chr(34) & ") then" & vbCrLf
        str = str & "	Response.Redirect " & Chr(34) & Me.memListFile & Chr(34) & vbCrLf
        str = str & "End If" & vbCrLf

        str = str & vbCrLf & ("If request(" & Chr(34) & "btnSubmit" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call saveSession()")
        str = str & vbCrLf & (vbIndent & "err = checkError()")
        str = str & vbCrLf & (vbIndent & "if err = " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & vbIndent & "call updateData(ID)")
        str = str & vbCrLf & (vbIndent & vbIndent & "Response.Redirect " & Chr(34) & Me.memConfirmFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID ")
        str = str & vbCrLf & (vbIndent & "end if")
        str = str & vbCrLf & ("ElseIf request(" & Chr(34) & "btnCancel" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.memConfirmFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID ")
        str = str & vbCrLf & ("End if")
        str = str & vbCrLf & (("%" & ">"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.headerFilePath & Chr(34) & "-->"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--enter the specific registration information here-->"))
        str = str & vbCrLf & (("<h3>Registration Information</h3>"))        str = str & vbCrLf & (("<!--end the specific registration information-->"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<" & "%"))
        str = str & vbCrLf & ("if err <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & ("Response.Write " & Chr(34) & "<font color=red>Registration Error: <br>" & Chr(34) & " & err & " & Chr(34) & "</font>" & Chr(34)))
        str = str & vbCrLf & ("end if")
        str = str & vbCrLf & (("%" & ">"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<form  method=post name=RegisterForm id=RegisterForm>"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incEditTableFile & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<br>"))
        str = str & vbCrLf & (("<table><tr><td align=left>"))
        str = str & vbCrLf & (("<input type=submit name=btnSubmit value=Submit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (("<input type=reset name=btnReset value=Reset style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("<input type=submit name=btnCancel value=Cancel style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("</td></tr></table>"))
        str = str & vbCrLf & (("</form>"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.footerFilePath & Chr(34) & "-->"))

        Return str
    End Function


    Public Function generateMemConfirmFile() As String
        Dim str As String = ""

        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.loginMemSecFilePath & Chr(34) & "-->"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incFuncFile & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & (("<" & "%"))
        str = str & vbCrLf & (("Dim ID, SRC"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("ID = Request.QueryString(" & Chr(34) & "id" & Chr(34) & ")"))
        str = str & vbCrLf & (("SRC = Request.QueryString(" & Chr(34) & "src" & Chr(34) & ")"))
        str = str & vbCrLf & vbCrLf

        str = str & "If getRegMemID(ID) <> session(" & Chr(34) & "_isMem" & Chr(34) & ") then" & vbCrLf
        str = str & "	Response.Redirect " & Chr(34) & Me.memListFile & Chr(34) & vbCrLf
        str = str & "End If" & vbCrLf

        str = str & vbCrLf & (("call retrieveData(ID)"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & ("if request(" & Chr(34) & "btnEdit" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")

        str = str & vbCrLf & ("'To edit")
        str = str & vbCrLf & ("Response.Redirect " & Chr(34) & Me.memEditFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID ")

        str = str & vbCrLf & ("'To go back")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnBack" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.memListFile & Chr(34))

        str = str & vbCrLf & ("'To approve")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnYesApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call approveData(ID)")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.memListFile & Chr(34))
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnNoApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")

        str = str & vbCrLf & ("'To disapprove")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnDisApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnYesDisApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call disapproveData(ID)")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.memListFile & Chr(34))
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnNoDisApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")

        str = str & vbCrLf & ("'To delete")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnDelete" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnYesDelete" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call deleteData(ID)")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.memListFile & Chr(34))
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnNoDelete" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("end if")

        str = str & vbCrLf & ("%" & ">")

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.headerFilePath & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--enter the specific registration information here-->"))
        str = str & vbCrLf & (("<h3>Registration Information</h3>"))        str = str & vbCrLf & (("<!--end the specific registration information-->"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<form  method=post name=RegisterConfirmForm id=RegisterConfirmForm>"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incConfirmTableFile & Chr(34) & "-->"))
        str = str & vbCrLf & (("<br>"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & (("<br>"))
        str = str & vbCrLf & (("<table><tr><td align=left>"))
        str = str & vbCrLf & (("<" & "%if request(" & Chr(34) & "btnDelete" & Chr(34) & ")") & " <> " & Chr(34) & Chr(34) & (" then%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<font color=red><b>Are you sure to delete?</b></font>"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnYesDelete value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnNoReset value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("<" & "%elseif request(" & Chr(34) & "btnApprove" & Chr(34) & ")") & " <> " & Chr(34) & Chr(34) & (" then%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<font color=red><b>Are you sure to appove?</b></font>"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnYesApprove value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnNoApprove value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (("<" & "%elseif request(" & Chr(34) & "btnDisapprove" & Chr(34) & ")") & " <> " & Chr(34) & Chr(34) & (" then%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<font color=red><b>Are you sure to disappove?</b></font>"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnYesDisapprove value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnNoDisapprove value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (("<" & "%else%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnDelete value=Delete style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<" & "%if SRC=" & Chr(34) & "pending" & Chr(34) & "then%" & ">"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnApprove value=Approve style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (vbIndent & ("<" & "%elseif SRC=" & Chr(34) & "approved" & Chr(34) & "then%" & ">"))        str = str & vbCrLf & (vbIndent & vbIndent & ("<input type=submit name=btnDisapprove value=Disapprove style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<" & "%end if%" & ">"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnEdit value=Edit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnBack value=Back style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("<" & "%end if%" & ">"))        str = str & vbCrLf & (("</td></tr></table>"))
        str = str & vbCrLf & (("</form>"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.footerFilePath & Chr(34) & "-->"))

        Return str
    End Function



    Public Function generateAdmListFile() As String
        Dim str As String = ""

        str = str & (("<!--#include file=" & Chr(34) & Me.loginAdmSecFilePath & Chr(34) & "-->")) & vbCrLf
        str = str & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incFuncFile & Chr(34) & "-->")) & vbCrLf

        str = str & vbCrLf & vbCrLf

        str = str & (("<!--#include file=" & Chr(34) & Me.headerFilePath & Chr(34) & "-->")) & vbCrLf

        str = str & vbCrLf
        str = str & (("<center>")) & vbCrLf

        str = str & vbCrLf
        str = str & (("<!--enter the specific registration information here-->")) & vbCrLf
        str = str & (("<h2>Registration Information for " & DataBaseTable1Name & " </h2>")) & vbCrLf
        str = str & (("<a href='" & Me.admHomeFilePath & "'>Back to Admin Home</a>")) & vbCrLf        str = str & (("<!--end the specific registration information-->")) & vbCrLf
        str = str & vbCrLf
        str = str & (("<" & "% call getPendingListAdm() %" & ">")) & vbCrLf
        str = str & (("<" & "% call getApprovedListAdm() %" & ">")) & vbCrLf
        If Me.IMSType = 2 Then
            str = str & (("<" & "% call getPendingFinalListAdm() %" & ">")) & vbCrLf
            str = str & (("<" & "% call getApprovedFinalListAdm() %" & ">")) & vbCrLf
        End If
        str = str & vbCrLf

        str = str & (("<P align=center>&nbsp;</P>")) & vbCrLf
        str = str & (("</center>")) & vbCrLf
        str = str & vbCrLf
        str = str & (("<!--#include file=" & Chr(34) & Me.footerFilePath & Chr(34) & "-->")) & vbCrLf

        Return str
    End Function


    Public Function generateAdmEditFile() As String
        Dim str As String = ""

        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.loginAdmSecFilePath & Chr(34) & "-->"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incFuncFile & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf & vbCrLf

        str = str & vbCrLf & (("<" & "%"))
        str = str & vbCrLf & (("Dim ID, SRC, err"))
        str = str & vbCrLf & (("err = " & Chr(34) & Chr(34)))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("ID = Request.QueryString(" & Chr(34) & "id" & Chr(34) & ")"))
        str = str & vbCrLf & (("SRC = Request.QueryString(" & Chr(34) & "src" & Chr(34) & ")"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & ("If request(" & Chr(34) & "btnSubmit" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call saveSession()")
        str = str & vbCrLf & (vbIndent & "err = checkError()")
        str = str & vbCrLf & (vbIndent & "if err = " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & vbIndent & "call updateData(ID)")
        str = str & vbCrLf & (vbIndent & vbIndent & "Response.Redirect " & Chr(34) & Me.admConfirmFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID ")
        str = str & vbCrLf & (vbIndent & "end if")
        str = str & vbCrLf & ("ElseIf request(" & Chr(34) & "btnCancel" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.admConfirmFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID ")
        str = str & vbCrLf & ("End if")
        str = str & vbCrLf & (("%" & ">"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.headerFilePath & Chr(34) & "-->"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--enter the specific registration information here-->"))
        str = str & vbCrLf & (("<h3>Registration Information</h3>"))        str = str & vbCrLf & (("<!--end the specific registration information-->"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<" & "%"))
        str = str & vbCrLf & ("if err <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & ("Response.Write " & Chr(34) & "<font color=red>Registration Error: <br>" & Chr(34) & " & err & " & Chr(34) & "</font>" & Chr(34)))
        str = str & vbCrLf & ("end if")
        str = str & vbCrLf & (("%" & ">"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<form  method=post name=RegisterForm id=RegisterForm>"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incEditTableFile & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<br>"))
        str = str & vbCrLf & (("<table><tr><td align=left>"))
        str = str & vbCrLf & (("<input type=submit name=btnSubmit value=Submit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (("<input type=reset name=btnReset value=Reset style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("<input type=submit name=btnCancel value=Cancel style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("</td></tr></table>"))
        str = str & vbCrLf & (("</form>"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.footerFilePath & Chr(34) & "-->"))

        Return str
    End Function


    Public Function generateAdmConfirmFile() As String
        Dim str As String = ""

        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.loginAdmSecFilePath & Chr(34) & "-->"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incFuncFile & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & (("<" & "%"))
        str = str & vbCrLf & (("Dim ID, SRC"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("ID = Request.QueryString(" & Chr(34) & "id" & Chr(34) & ")"))
        str = str & vbCrLf & (("SRC = Request.QueryString(" & Chr(34) & "src" & Chr(34) & ")"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & (("call retrieveData(ID)"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & ("if request(" & Chr(34) & "btnEdit" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")

        str = str & vbCrLf & ("'To edit")
        str = str & vbCrLf & ("Response.Redirect " & Chr(34) & Me.admEditFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID ")

        str = str & vbCrLf & ("'To go back")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnBack" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.admListFile & Chr(34))

        str = str & vbCrLf & ("'To approve")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnYesApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call approveData(ID)")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.admListFile & Chr(34))
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnNoApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")

        str = str & vbCrLf & ("'To disapprove")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnDisApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnYesDisApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call disapproveData(ID)")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.admListFile & Chr(34))
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnNoDisApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")

        str = str & vbCrLf & ("'To delete")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnDelete" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnYesDelete" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call deleteData(ID)")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.admListFile & Chr(34))
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnNoDelete" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("end if")

        str = str & vbCrLf & ("%" & ">")

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.headerFilePath & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--enter the specific registration information here-->"))
        str = str & vbCrLf & (("<h3>Registration Information</h3>"))        str = str & vbCrLf & (("<!--end the specific registration information-->"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<form  method=post name=RegisterConfirmForm id=RegisterConfirmForm>"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incConfirmTableFile & Chr(34) & "-->"))
        str = str & vbCrLf & (("<br>"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & (("<br>"))
        str = str & vbCrLf & (("<table><tr><td align=left>"))
        str = str & vbCrLf & (("<" & "%if request(" & Chr(34) & "btnDelete" & Chr(34) & ")") & " <> " & Chr(34) & Chr(34) & (" then%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<font color=red><b>Are you sure to delete?</b></font>"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnYesDelete value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnNoReset value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("<" & "%elseif request(" & Chr(34) & "btnApprove" & Chr(34) & ")") & " <> " & Chr(34) & Chr(34) & (" then%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<font color=red><b>Are you sure to appove?</b></font>"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnYesApprove value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnNoApprove value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (("<" & "%elseif request(" & Chr(34) & "btnDisapprove" & Chr(34) & ")") & " <> " & Chr(34) & Chr(34) & (" then%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<font color=red><b>Are you sure to disappove?</b></font>"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnYesDisapprove value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnNoDisapprove value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (("<" & "%else%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnDelete value=Delete style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<" & "%if SRC=" & Chr(34) & "pending" & Chr(34) & "then%" & ">"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnApprove value=Approve style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (vbIndent & ("<" & "%elseif SRC=" & Chr(34) & "approved" & Chr(34) & "then%" & ">"))        str = str & vbCrLf & (vbIndent & vbIndent & ("<input type=submit name=btnDisapprove value=Disapprove style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<" & "%end if%" & ">"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnEdit value=Edit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnBack value=Back style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("<" & "%end if%" & ">"))        str = str & vbCrLf & (("</td></tr></table>"))
        str = str & vbCrLf & (("</form>"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.footerFilePath & Chr(34) & "-->"))

        Return str
    End Function


    ''''''''''''''''''''''''''''''''''''''''''''
    ' Regitration include/regFuncs.asp
    ''''''''''''''''''''''''''''''''''''''''''''

    Public Function generateIncFuncFile() As String
        Dim i As Integer

        Dim str As String = ""

        str = str & ("<!--#include file=" & Chr(34) & Me.websiteFuncFilePath & Chr(34) & "-->" & vbCrLf)

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & ("<%")

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function saveSession()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Saves session variables.")
        str = str & ("sub saveSession()" & vbCrLf)
        For i = 0 To UBound(formArray)
            str = str & (vbIndent & "session(" & Chr(34) & sessionArray(i) & Chr(34) & ") = trim(request(" & Chr(34) & formArray(i) & Chr(34) & "))" & vbCrLf)
        Next
        str = str & ("end sub") & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function clearSession()	
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Clears session variables.")
        str = str & ("sub clearSession()" & vbCrLf)
        For i = 0 To UBound(formArray)
            str = str & (vbIndent & "session(" & Chr(34) & sessionArray(i) & Chr(34) & ") = " & Chr(34) & Chr(34) & vbCrLf)
        Next
        str = str & ("end sub") & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function insertData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Insert data into database table " & DataBaseTable1Name & ".")
        str = str & ("Sub insertData()" & vbCrLf)
        str = str & (vbIndent & "Dim db, sql, rs") & vbCrLf
        str = str & vbCrLf

        str = str & (vbIndent & "sql = " & Chr(34) & "INSERT INTO " & DataBaseTable1Name & " (" & Chr(34) & vbCrLf)
        str = str & (vbIndent & "sql = sql & " & Chr(34) & "memID, " & Chr(34) & vbCrLf)
        str = str & (vbIndent & "sql = sql & " & Chr(34) & "regDate, " & Chr(34) & vbCrLf)
        str = str & (vbIndent & "sql = sql & " & Chr(34) & "approval, " & Chr(34) & vbCrLf)
        For i = 0 To UBound(formArray) - 1
            str = str & (vbIndent & "sql = sql & " & Chr(34) & dbArray(i) & ", " & Chr(34) & vbCrLf)
        Next
        str = str & (vbIndent & "sql = sql & " & Chr(34) & dbArray(i) & Chr(34) & vbCrLf)

        str = str & vbCrLf
        str = str & (vbIndent & "sql = sql & " & Chr(34) & ") VALUES (" & Chr(34) & vbCrLf)
        str = str & vbCrLf

        If Me.IMSType = 1 Or Me.IMSType = 2 Then
            str = str & vbCrLf
            str = str & vbIndent & "if session(" & Chr(34) & "_isMem" & Chr(34) & ") = " & Chr(34) & "" & Chr(34) & " then" & vbCrLf
            str = str & vbIndent & vbIndent & "sql = sql & " & Chr(34) & "-1, " & Chr(34) & vbCrLf
            str = str & vbIndent & "else" & vbCrLf
            str = str & vbIndent & vbIndent & "sql = sql & session(" & Chr(34) & "_isMem" & Chr(34) & ") & " & Chr(34) & ", " & Chr(34) & vbCrLf
            str = str & vbIndent & "end if" & vbCrLf
            str = str & vbCrLf
        End If

        '        If Me.IMSType = 1 Or Me.IMSType = 2 Then
        '       str = str & vbIndent & "sql = sql & session(" & Chr(34) & "_isMem" & Chr(34) & ") & " & Chr(34) & ", " & Chr(34) & vbCrLf
        '      End If
        str = str & (vbIndent & "sql = sql & sqlEncode(date()) & " & Chr(34) & ", " & Chr(34) & vbCrLf)
        ''''''''''''''''''''''''' insert approval
        str = str & (vbIndent & "sql = sql & " & Chr(34) & "0" & Chr(34) & " & " & Chr(34) & ", " & Chr(34) & vbCrLf)
        'str = str &  vbIndent & "sql = sql & 0 & " & chr(34) & ", " & chr(34) & vbcrlf
        For i = 0 To UBound(dbArray) - 1
            str = str & (vbIndent & "sql = sql & sqlEncode(session(" & Chr(34) & sessionArray(i) & Chr(34) & "))" & " & " & Chr(34) & ", " & Chr(34) & vbCrLf)
        Next
        str = str & (vbIndent & "sql = sql & sqlEncode(session(" & Chr(34) & sessionArray(i) & Chr(34) & "))" & " & " & Chr(34) & ") " & Chr(34) & vbCrLf)

        str = str & vbCrLf
        str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
        str = str & (vbIndent & "call dbExecute(db, sql)" & vbCrLf)
        str = str & vbCrLf
        str = str & (vbIndent & "call db.close()" & vbCrLf)
        str = str & (vbIndent & "set db = nothing" & vbCrLf)

        str = str & ("end sub") & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function updateData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Update data in database table " & DataBaseTable1Name & ".")
        str = str & ("sub updateData(id)" & vbCrLf)
        str = str & (vbIndent & "Dim db, sql" & vbCrLf)
        str = str & vbCrLf

        str = str & (vbIndent & "sql = " & Chr(34) & "UPDATE " & DataBaseTable1Name & " SET " & Chr(34) & vbCrLf)
        For i = 0 To UBound(dbArray) - 1
            str = str & (vbIndent & "sql = sql & " & Chr(34) & dbArray(i) & " = " & Chr(34) & " & sqlEncode(session(" & Chr(34) & sessionArray(i) & Chr(34) & ")) & " & Chr(34) & ", " & Chr(34) & vbCrLf)
        Next
        str = str & (vbIndent & "sql = sql & " & Chr(34) & dbArray(i) & " = " & Chr(34) & " & sqlEncode(session(" & Chr(34) & sessionArray(i) & Chr(34) & "))" & vbCrLf)
        str = str & (vbIndent & "sql = sql & " & Chr(34) & " WHERE ID = " & Chr(34) & " & id " & vbCrLf)

        str = str & vbCrLf
        str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
        str = str & (vbIndent & "call dbExecute(db, sql)" & vbCrLf)
        str = str & vbCrLf
        str = str & (vbIndent & "call db.close()" & vbCrLf)
        str = str & (vbIndent & "set db = nothing" & vbCrLf)
        str = str & ("end sub") & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function deleteData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Delete data in database table " & DataBaseTable1Name & ".")
        str = str & ("sub deleteData(id)" & vbCrLf)
        str = str & (vbIndent & "Dim db, sql" & vbCrLf)

        str = str & (vbIndent & "sql = " & Chr(34) & "DELETE FROM " & DataBaseTable1Name & " WHERE ID = " & Chr(34) & " & id" & vbCrLf)
        str = str & vbCrLf

        str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
        str = str & (vbIndent & "call dbExecute(db, sql)" & vbCrLf)
        str = str & vbCrLf
        str = str & (vbIndent & "call db.close()" & vbCrLf)
        str = str & (vbIndent & "set db = nothing" & vbCrLf)
        str = str & ("end sub") & vbCrLf
        str = str & vbCrLf & vbCrLf


        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function retrieveData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Retrieve data from database table " & DataBaseTable1Name & ".")
        str = str & ("sub retrieveData(id)" & vbCrLf)
        str = str & (vbIndent & "Dim db, sql, sr, count" & vbCrLf)

        str = str & (vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & DataBaseTable1Name & " WHERE ID = " & Chr(34) & " & id" & vbCrLf)
        str = str & vbCrLf
        str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
        str = str & (vbIndent & "set rs = dbExecuteRS(db, sql)" & vbCrLf)
        str = str & vbCrLf

        For i = 0 To UBound(sessionArray)
            str = str & (vbIndent & "session(" & Chr(34) & sessionArray(i) & Chr(34) & ") = trim(rs(" & Chr(34) & dbArray(i) & Chr(34) & "))" & vbCrLf)
        Next

        str = str & vbCrLf
        str = str & (vbIndent & "rs.close()" & vbCrLf)
        str = str & (vbIndent & "set rs = nothing" & vbCrLf)
        str = str & vbCrLf
        str = str & (vbIndent & "call db.close()" & vbCrLf)
        str = str & (vbIndent & "set db = nothing" & vbCrLf)
        str = str & ("end sub") & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function approveData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Approve registration in database table " & Me.DataBaseTable1Name & ".")
        str = str & ("sub approveData(id)" & vbCrLf)
        str = str & (vbIndent & "Dim db, sql" & vbCrLf)

        str = str & (vbIndent & "sql = " & Chr(34) & "UPDATE " & DataBaseTable1Name & " SET approval = 1 WHERE ID = " & Chr(34) & " & id" & vbCrLf)

        str = str & vbCrLf
        str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
        str = str & (vbIndent & "call dbExecute(db, sql)" & vbCrLf)
        str = str & vbCrLf
        str = str & (vbIndent & "call db.close()" & vbCrLf)
        str = str & (vbIndent & "set db = nothing" & vbCrLf)

        If Me.IMSType = 2 Then
            str = str & vbCrLf
            str = str & vbIndent & "Call copyDataToFinal(id)" & vbCrLf
            str = str & vbCrLf
        End If

        str = str & ("end sub") & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function disapproveData()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Disapprove registration in database table " & Me.DataBaseTable1Name & ".")
        str = str & ("sub disApproveData(id)" & vbCrLf)
        str = str & (vbIndent & "Dim db, sql" & vbCrLf)

        str = str & (vbIndent & "sql = " & Chr(34) & "UPDATE " & DataBaseTable1Name & " SET approval = 0 WHERE ID = " & Chr(34) & " & id" & vbCrLf)

        str = str & vbCrLf
        str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
        str = str & (vbIndent & "call dbExecute(db, sql)" & vbCrLf)
        str = str & vbCrLf
        str = str & (vbIndent & "call db.close()" & vbCrLf)
        str = str & (vbIndent & "set db = nothing" & vbCrLf)
        str = str & ("end sub") & vbCrLf
        str = str & vbCrLf & vbCrLf

        If Me.IMSType = 1 Or Me.IMSType = 2 Then
            ' Start - not for IMS type 0
            str = str & writeFunctionComment("Get 'memID' field given 'ID' of the record. Used to prevent a member changing another member's data.")
            str = str & "Function getRegMemID(id)" & vbCrLf
            str = str & "    Dim db, sql, sr, count" & vbCrLf
            str = str & "    sql = " & Chr(34) & "SELECT memID FROM " & Me.DataBaseTable1Name & " WHERE ID = " & Chr(34) & " & id" & vbCrLf
            str = str & vbCrLf
            str = str & "    set db = dbConnect()" & vbCrLf
            str = str & "    set rs = dbExecuteRS(db, sql)" & vbCrLf
            str = str & vbCrLf
            str = str & "    getRegMemID = rs(" & Chr(34) & "memID" & Chr(34) & ")" & vbCrLf
            str = str & vbCrLf
            str = str & "    call dbRSClose(db, rs)" & vbCrLf
            str = str & "End Function" & vbCrLf
            str = str & vbCrLf & vbCrLf

            '''''''''''''''''''''''''''''''''''''''''''''
            ' Create function getPendingListMem()
            '''''''''''''''''''''''''''''''''''''''''''''
            str = str & writeFunctionComment("Get pending registration in database table " & DataBaseTable1Name & ".")
            str = str & ("sub getPendingListMem()" & vbCrLf)
            str = str & (vbIndent & "Dim db, sql, count" & vbCrLf)

            str = str & (vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & DataBaseTable1Name & " WHERE approval = 0 AND memID = " & Chr(34) & " & " & "session(" & Chr(34) & "_isMem" & Chr(34) & ")")

            str = str & vbCrLf
            str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
            str = str & (vbIndent & "set rs =  dbExecuteRS(db, sql)" & vbCrLf)
            str = str & vbCrLf

            str = str & (vbIndent & "if rs.recordcount > 0 then") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<H2>List of Pending Registrants</H2>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<table width=600 border=1>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td colspan=3>There are " & Chr(34) & " & rs.recordcount & " & Chr(34) & " pending registrants.</td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td width=150><b>Register Date</b></td><td width=150><b>Name</b></td><td><b>Detail Information</b></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "do while not rs.eof") & vbCrLf
            'str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>" & Chr(34) & " & rs(" & Chr(34) & "lastName" & Chr(34) & ") & " & Chr(34) & ", " & Chr(34) & " & rs(" & Chr(34) & "firstName" & Chr(34) & ") & " & Chr(34) & "</td><td><a href=" & Chr(39) & Me.admConfirmFile & "?src=pending&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>Dummy Field</td><td><a href=" & Chr(39) & Me.memConfirmFile & "?src=pending&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & vbIndent & "rs.MoveNext()") & vbCrLf
            str = str & (vbIndent & vbIndent & "loop") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "</table>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & "else") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<h2>Currently there is no pending registrant.</h2>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & "end if") & vbCrLf

            str = str & vbCrLf
            str = str & (vbIndent & "rs.close()" & vbCrLf)
            str = str & (vbIndent & "set rs = nothing" & vbCrLf)
            str = str & vbCrLf
            str = str & (vbIndent & "call db.close()" & vbCrLf)
            str = str & (vbIndent & "set db = nothing" & vbCrLf)
            str = str & ("end sub") & vbCrLf
            str = str & vbCrLf & vbCrLf

            '''''''''''''''''''''''''''''''''''''''''''''
            ' Create function getApprovedListMem()
            '''''''''''''''''''''''''''''''''''''''''''''
            str = str & writeFunctionComment("Get approved registration in database table " & DataBaseTable1Name & ".")
            str = str & ("sub getApprovedListMem()" & vbCrLf)
            str = str & (vbIndent & "Dim db, sql, count" & vbCrLf)

            'str = str & (vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & DataBaseTable1Name & " WHERE approval = 1" & Chr(34))
            str = str & (vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & DataBaseTable1Name & " WHERE approval = 1 AND memID = " & Chr(34) & " & " & "session(" & Chr(34) & "_isMem" & Chr(34) & ")")

            str = str & vbCrLf
            str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
            str = str & (vbIndent & "set rs =  dbExecuteRS(db, sql)" & vbCrLf)
            str = str & vbCrLf

            str = str & (vbIndent & "if rs.recordcount > 0 then") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<H2>List of Approved Registrants</H2>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<table width=600 border=1>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td colspan=3>There are " & Chr(34) & " & rs.recordcount & " & Chr(34) & " approved registrants.</td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td width=150><b>Register Date</b></td><td width=150><b>Name</b></td><td><b>Detail Information</b></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "do while not rs.eof") & vbCrLf
            'str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>" & Chr(34) & " & rs(" & Chr(34) & "lastName" & Chr(34) & ") & " & Chr(34) & ", " & Chr(34) & " & rs(" & Chr(34) & "firstName" & Chr(34) & ") & " & Chr(34) & "</td><td><a href=" & Chr(39) & Me.admConfirmFile & "?src=approved&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>Dummy Field</td><td><a href=" & Chr(39) & Me.memConfirmFile & "?src=approved&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & vbIndent & "rs.MoveNext()") & vbCrLf
            str = str & (vbIndent & vbIndent & "loop") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "</table>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & "else") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<h2>Currently there is no approved registrant.</h2>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & "end if") & vbCrLf

            str = str & vbCrLf
            str = str & (vbIndent & "rs.close()" & vbCrLf)
            str = str & (vbIndent & "set rs = nothing" & vbCrLf)
            str = str & vbCrLf
            str = str & (vbIndent & "call db.close()" & vbCrLf)
            str = str & (vbIndent & "set db = nothing" & vbCrLf)
            str = str & ("end sub") & vbCrLf
            str = str & vbCrLf & vbCrLf

            ' End - not for IMS type 0 
        End If

        If Me.IMSType = 2 Then
            str = str & writeFunctionComment("Copy date from " & Me.DataBaseTable1Name & " to database table " & Me.DataBaseTable2Name & ".")
            str = str & vbCrLf
            str = str & "sub copyDataToFinal(regId)" & vbCrLf
            str = str & "    Dim db, rs, sql" & vbCrLf
            str = str & vbCrLf
            str = str & "    sql = " & Chr(34) & "SELECT * FROM " & Me.DataBaseTable2Name & " WHERE regID = " & Chr(34) & " & id & " & Chr(34) & vbCrLf
            str = str & vbCrLf
            str = str & "    set db = dbConnect()" & vbCrLf
            str = str & "    set rs = dbExecuteFO(db, sql)" & vbCrLf
            str = str & vbCrLf
            str = str & "    if rs.eof then" & vbCrLf
            str = str & "        call insertData(" & Chr(34) & Me.DataBaseTable2Name & Chr(34) & ")" & vbCrLf
            str = str & "    else" & vbCrLf
            str = str & "        call updateData(" & Chr(34) & Me.DataBaseTable2Name & Chr(34) & ", id, memID)" & vbCrLf
            str = str & "    end if" & vbCrLf
            str = str & vbCrLf
            str = str & "    call db.close()" & vbCrLf
            str = str & "    set db = nothing" & vbCrLf
            str = str & "end sub" & vbCrLf
            str = str & vbCrLf & vbCrLf


            '''''''''''''''''''''''''''''''''''''''''''''
            ' Create function getPendingFinalListAdm()
            '''''''''''''''''''''''''''''''''''''''''''''
            str = str & writeFunctionComment("Get pending final registration in database table " & DataBaseTable2Name & ".")
            str = str & ("sub getPendingFinalListAdm()" & vbCrLf)
            str = str & (vbIndent & "Dim db, sql, count" & vbCrLf)

            str = str & (vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & DataBaseTable2Name & " WHERE approval = 0" & Chr(34))

            str = str & vbCrLf
            str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
            str = str & (vbIndent & "set rs =  dbExecuteRS(db, sql)" & vbCrLf)
            str = str & vbCrLf

            str = str & (vbIndent & "if rs.recordcount > 0 then") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<H2>List of Pending Registrants</H2>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<table width=600 border=1>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td colspan=3>There are " & Chr(34) & " & rs.recordcount & " & Chr(34) & " pending registrants.</td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td width=150><b>Register Date</b></td><td width=150><b>Name</b></td><td><b>Detail Information</b></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "do while not rs.eof") & vbCrLf
            'str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>" & Chr(34) & " & rs(" & Chr(34) & "lastName" & Chr(34) & ") & " & Chr(34) & ", " & Chr(34) & " & rs(" & Chr(34) & "firstName" & Chr(34) & ") & " & Chr(34) & "</td><td><a href=" & Chr(39) & Me.admConfirmFile & "?src=pending&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>Dummy Field</td><td><a href=" & Chr(39) & Me.admConfirmFile & "?src=pending&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & vbIndent & "rs.MoveNext()") & vbCrLf
            str = str & (vbIndent & vbIndent & "loop") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "</table>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & "else") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<h2>Currently there is no pending registrant.</h2>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & "end if") & vbCrLf

            str = str & vbCrLf
            str = str & (vbIndent & "rs.close()" & vbCrLf)
            str = str & (vbIndent & "set rs = nothing" & vbCrLf)
            str = str & vbCrLf
            str = str & (vbIndent & "call db.close()" & vbCrLf)
            str = str & (vbIndent & "set db = nothing" & vbCrLf)
            str = str & ("end sub") & vbCrLf
            str = str & vbCrLf & vbCrLf

            '''''''''''''''''''''''''''''''''''''''''''''
            ' Create function getApprovedFinalListAdm()
            '''''''''''''''''''''''''''''''''''''''''''''
            str = str & writeFunctionComment("Get approved registration in database table " & DataBaseTable2Name & ".")
            str = str & ("sub getApprovedFinalListAdm()" & vbCrLf)
            str = str & (vbIndent & "Dim db, sql, count" & vbCrLf)

            str = str & (vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & DataBaseTable2Name & " WHERE approval = 1" & Chr(34))

            str = str & vbCrLf
            str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
            str = str & (vbIndent & "set rs =  dbExecuteRS(db, sql)" & vbCrLf)
            str = str & vbCrLf

            str = str & (vbIndent & "if rs.recordcount > 0 then") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<H2>List of Approved Registrants</H2>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<table width=600 border=1>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td colspan=3>There are " & Chr(34) & " & rs.recordcount & " & Chr(34) & " approved registrants.</td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td width=150><b>Register Date</b></td><td width=150><b>Name</b></td><td><b>Detail Information</b></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & "do while not rs.eof") & vbCrLf
            'str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>" & Chr(34) & " & rs(" & Chr(34) & "lastName" & Chr(34) & ") & " & Chr(34) & ", " & Chr(34) & " & rs(" & Chr(34) & "firstName" & Chr(34) & ") & " & Chr(34) & "</td><td><a href=" & Chr(39) & Me.admConfirmFile & "?src=approved&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>Dummy Field</td><td><a href=" & Chr(39) & Me.admConfirmFile & "?src=approved&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & vbIndent & vbIndent & "rs.MoveNext()") & vbCrLf
            str = str & (vbIndent & vbIndent & "loop") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "</table>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & "else") & vbCrLf
            str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<h2>Currently there is no approved registrant.</h2>" & Chr(34)) & vbCrLf
            str = str & (vbIndent & "end if") & vbCrLf

            str = str & vbCrLf
            str = str & (vbIndent & "rs.close()" & vbCrLf)
            str = str & (vbIndent & "set rs = nothing" & vbCrLf)
            str = str & vbCrLf
            str = str & (vbIndent & "call db.close()" & vbCrLf)
            str = str & (vbIndent & "set db = nothing" & vbCrLf)
            str = str & ("end sub") & vbCrLf
            str = str & vbCrLf & vbCrLf
        End If

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function getPendingListAdm()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Get pending registration in database table " & DataBaseTable1Name & ".")
        str = str & ("sub getPendingListAdm()" & vbCrLf)
        str = str & (vbIndent & "Dim db, sql, count" & vbCrLf)

        str = str & (vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & DataBaseTable1Name & " WHERE approval = 0" & Chr(34))

        str = str & vbCrLf
        str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
        str = str & (vbIndent & "set rs =  dbExecuteRS(db, sql)" & vbCrLf)
        str = str & vbCrLf

        str = str & (vbIndent & "if rs.recordcount > 0 then") & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<H2>List of Pending Registrants</H2>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<table width=600 border=1>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td colspan=3>There are " & Chr(34) & " & rs.recordcount & " & Chr(34) & " pending registrants.</td></tr>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td width=150><b>Register Date</b></td><td width=150><b>Name</b></td><td><b>Detail Information</b></td></tr>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & "do while not rs.eof") & vbCrLf
        'str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>" & Chr(34) & " & rs(" & Chr(34) & "lastName" & Chr(34) & ") & " & Chr(34) & ", " & Chr(34) & " & rs(" & Chr(34) & "firstName" & Chr(34) & ") & " & Chr(34) & "</td><td><a href=" & Chr(39) & Me.admConfirmFile & "?src=pending&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>Dummy Field</td><td><a href=" & Chr(39) & Me.admConfirmFile & "?src=pending&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & vbIndent & "rs.MoveNext()") & vbCrLf
        str = str & (vbIndent & vbIndent & "loop") & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "</table>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & "else") & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<h2>Currently there is no pending registrant.</h2>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & "end if") & vbCrLf

        str = str & vbCrLf
        str = str & (vbIndent & "rs.close()" & vbCrLf)
        str = str & (vbIndent & "set rs = nothing" & vbCrLf)
        str = str & vbCrLf
        str = str & (vbIndent & "call db.close()" & vbCrLf)
        str = str & (vbIndent & "set db = nothing" & vbCrLf)
        str = str & ("end sub") & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function getApprovedListAdm()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Get approved registration in database table " & DataBaseTable1Name & ".")
        str = str & ("sub getApprovedListAdm()" & vbCrLf)
        str = str & (vbIndent & "Dim db, sql, count" & vbCrLf)

        str = str & (vbIndent & "sql = " & Chr(34) & "SELECT * FROM " & DataBaseTable1Name & " WHERE approval = 1" & Chr(34))

        str = str & vbCrLf
        str = str & (vbIndent & "set db = dbConnect()" & vbCrLf)
        str = str & (vbIndent & "set rs =  dbExecuteRS(db, sql)" & vbCrLf)
        str = str & vbCrLf

        str = str & (vbIndent & "if rs.recordcount > 0 then") & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<H2>List of Approved Registrants</H2>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<table width=600 border=1>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td colspan=3>There are " & Chr(34) & " & rs.recordcount & " & Chr(34) & " approved registrants.</td></tr>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td width=150><b>Register Date</b></td><td width=150><b>Name</b></td><td><b>Detail Information</b></td></tr>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & "do while not rs.eof") & vbCrLf
        'str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>" & Chr(34) & " & rs(" & Chr(34) & "lastName" & Chr(34) & ") & " & Chr(34) & ", " & Chr(34) & " & rs(" & Chr(34) & "firstName" & Chr(34) & ") & " & Chr(34) & "</td><td><a href=" & Chr(39) & Me.admConfirmFile & "?src=approved&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & vbIndent & "Response.Write " & Chr(34) & "<tr><td>" & Chr(34) & " & rs(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td><td>Dummy Field</td><td><a href=" & Chr(39) & Me.admConfirmFile & "?src=approved&id=" & Chr(34) & " & rs(" & Chr(34) & "ID" & Chr(34) & ") & " & Chr(34) & "" & Chr(39) & ">Detail</a></td></tr>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & vbIndent & vbIndent & "rs.MoveNext()") & vbCrLf
        str = str & (vbIndent & vbIndent & "loop") & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "</table>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & "else") & vbCrLf
        str = str & (vbIndent & vbIndent & "Response.Write " & Chr(34) & "<h2>Currently there is no approved registrant.</h2>" & Chr(34)) & vbCrLf
        str = str & (vbIndent & "end if") & vbCrLf

        str = str & vbCrLf
        str = str & (vbIndent & "rs.close()" & vbCrLf)
        str = str & (vbIndent & "set rs = nothing" & vbCrLf)
        str = str & vbCrLf
        str = str & (vbIndent & "call db.close()" & vbCrLf)
        str = str & (vbIndent & "set db = nothing" & vbCrLf)
        str = str & ("end sub") & vbCrLf
        str = str & vbCrLf & vbCrLf

        '''''''''''''''''''''''''''''''''''''''''''''
        ' Create function checkError()
        '''''''''''''''''''''''''''''''''''''''''''''
        str = str & writeFunctionComment("Check error in form data.")
        str = str & ("function checkError()") & vbCrLf
        str = str & (vbIndent & "checkError = " & Chr(34) & Chr(34) & vbCrLf)
        str = str & ("end function") & vbCrLf
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & "%>"

        Return str
    End Function


    ' Include Edit/Confirm files

    Public Function generateIncEditTableFile(Optional ByVal tblAttr _
    As clsHTMLTableAttributes = Nothing) As String
        Dim i, j, count As Integer
        Dim metadataArray() As String = Me.clsMD.metadataArray

        Dim tableStart, firstColWidth, secondColWidth As String
        xcUtil.getTableAttributes(tblAttr, tableStart, firstColWidth, secondColWidth)

        Dim str As String = tableStart & vbCrLf & vbCrLf

        str = str & "<%If session(" & Chr(34) & "regDate" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " Then" & vbCrLf
        str = str & "	Response.Write " & Chr(34) & "<tr><td" & firstColWidth & ">Register Date </td><td" & secondColWidth & ">" & Chr(34) & " & session(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td></tr>" & Chr(34) & vbCrLf
        str = str & "End if%>" & vbCrLf & vbCrLf

        count = UBound(metadataArray) - 1
        For i = 0 To count
            If LCase(metadataArray(i + 1)) = "text" Or _
               LCase(metadataArray(i + 1)) = "password" Or _
               LCase(metadataArray(i + 1)) = "hidden" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf ' label
                str = str & ("</td><td valign=top>") & vbCrLf
                str = str & ("<input type=" & metadataArray(i + 1) & " name=" & metadataArray(i + 2) & " size=" & metadataArray(i + 3) & " maxlength=" & metadataArray(i + 4) & " value=" & Chr(34) & "<%=TextHTMLEncode( session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ") )%>" & Chr(34) & ">") & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 5
            ElseIf LCase(metadataArray(i + 1)) = "select" Then
                str = str & ("<tr><td valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td valign=top>") & vbCrLf
                str = str & ("<select name=" & metadataArray(i + 2) & " style=" & Chr(34) & "WIDTH: " & metadataArray(i + 3) & Chr(34) & " size=" & metadataArray(i + 4) & ">") & vbCrLf
                count = CInt(metadataArray(i + 5))
                For j = 1 To count
                    str = str & ("<option value=" & j & " <" & "%if session(" & Chr(34) & metadataArray(i + 6) & Chr(34) & ") = " & Chr(34) & j & Chr(34) & " then%" & "> selected <" & "%end if%" & ">" & ">" & metadataArray(i + 6 + j)) & vbCrLf
                Next
                str = str & ("</select>") & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 5 + count
            ElseIf LCase(metadataArray(i + 1)) = "textarea" Then
                str = str & ("<tr><td valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td valign=top>") & vbCrLf
                str = str & ("<TEXTAREA NAME=" & metadataArray(i + 2) & " ROWS=" & CInt(metadataArray(i + 3)) & " COLS=" & CInt(metadataArray(i + 4)) & " " & metadataArray(i + 5) & ">" & "<%=TextAreaEditHTMLEncode( session(" & Chr(34) & metadataArray(i + 6) & Chr(34) & ") )%>" & "</TEXTAREA>") & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 6
            ElseIf LCase(metadataArray(i + 1)) = "radio" Then
                str = str & ("<tr><td valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td valign=top>") & vbCrLf
                count = CInt(metadataArray(i + 3))
                For j = 1 To count
                    str = str & (metadataArray(i + 4 + j) & "<input type=radio name=" & metadataArray(i + 2) & " value=" & j & " <" & "%if session(" & Chr(34) & metadataArray(i + 4) & Chr(34) & ")=" & Chr(34) & j & Chr(34) & " then%" & "> checked <" & "%end if%" & ">>") & vbCrLf
                Next
                str = str & ("</td></tr>") & vbCrLf
                i = i + 3 + count
            ElseIf LCase(metadataArray(i + 1)) = "checkbox" Then
                str = str & ("<tr><td valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td valign=top>") & vbCrLf
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
                    str = str & "<tr><td" & firstColWidth & "><br></td>" & "<td><br></td></tr>" & vbCrLf
                End If
                i = i + 1
            End If
            str = str & vbCrLf
        Next

        str = str & ("</table>") & vbCrLf

        Return str
    End Function


    ' nodeIncConfirmTable.Text = "confirTable.asp"
    Public Function generateIncConfirmTableFile(Optional ByVal tblAttr _
        As clsHTMLTableAttributes = Nothing) As String
        Dim metadataArray() As String = Me.clsMD.metadataArray
        Dim i, j, count As Integer

        Dim tableStart, firstColWidth, secondColWidth As String
        xcUtil.getTableAttributes(tblAttr, tableStart, firstColWidth, secondColWidth)

        Dim str As String = tableStart & vbCrLf & vbCrLf

        str = str & "<%If session(" & Chr(34) & "regDate" & Chr(34) & ") <> " & Chr(34) & "" & Chr(34) & " Then" & vbCrLf
        str = str & "	Response.Write " & Chr(34) & "<tr><td" & firstColWidth & ">Register Date </td><td" & secondColWidth & ">" & Chr(34) & " & session(" & Chr(34) & "regDate" & Chr(34) & ") & " & Chr(34) & "</td></tr>" & Chr(34) & vbCrLf
        str = str & "End if%>" & vbCrLf & vbCrLf

        count = UBound(metadataArray) - 1
        For i = 0 To count
            If LCase(metadataArray(i + 1)) = "text" Then
                str = str & ("<tr><td" & firstColWidth & " valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf ' label
                str = str & ("</td><td" & secondColWidth & " valign=top><b>") & vbCrLf
                str = str & ("<%=session(" & Chr(34) & metadataArray(i + 5) & Chr(34) & ")%>" & "&nbsp;") & vbCrLf
                str = str & ("</b></td></tr>") & vbCrLf
                i = i + 5
            ElseIf LCase(metadataArray(i + 1)) = "password" Then
                str = str & ("<tr><td valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf ' label
                str = str & ("</td><td valign=top>") & vbCrLf
                str = str & ("<b>********</b>") & vbCrLf
                str = str & ("</td></tr>") & vbCrLf
                i = i + 5
            ElseIf LCase(metadataArray(i + 1)) = "hidden" Then
                i = i + 5
            ElseIf LCase(metadataArray(i + 1)) = "select" Then
                str = str & ("<tr><td valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td valign=top>") & vbCrLf
                count = CInt(metadataArray(i + 5))
                For j = 1 To count
                    str = str & ("<" & "%if session(" & Chr(34) & metadataArray(i + 6) & Chr(34) & ") = " & Chr(34) & j & Chr(34) & " then%" & "><b>" & metadataArray(i + 6 + j) & "</b><" & "%end if%" & ">") & vbCrLf
                Next
                str = str & ("&nbsp;</td></tr>") & vbCrLf
                i = i + 5 + count
            ElseIf LCase(metadataArray(i + 1)) = "textarea" Then
                str = str & ("<tr><td valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td valign=top>") & vbCrLf

                str = str & ("<" & "%if session(" & Chr(34) & metadataArray(i + 6) & Chr(34) & ") <> " & _
                Chr(34) & Chr(34) & " then%" & "><b>" & "<%=TextAreaViewHTMLEncode( session(" & Chr(34) & metadataArray(i + 6) & Chr(34) & ") )%>" & "</b><" & _
                "%else%" & "> <br> <" & "%end if%" & ">") & vbCrLf

                str = str & ("</td></tr>") & vbCrLf
                i = i + 6
            ElseIf LCase(metadataArray(i + 1)) = "radio" Then
                str = str & ("<tr><td valign=top>") & vbCrLf
                str = str & (metadataArray(i)) & vbCrLf
                str = str & ("</td><td valign=top>") & vbCrLf
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
                  "<tr><td valign=top>" & metadataArray(i) & "</td><td valign=top><b>" & _
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
                    str = str & "<tr><td" & firstColWidth & "><br></td>" & "<td><br></td></tr>" & vbCrLf
                End If
                i = i + 1
            End If
            str = str & vbCrLf
        Next

        str = str & ("</table>")
        Return str
    End Function


    ' filePath - Path of the file to insert link.
    ' taskListStart - Start mark of the link lists in the file. Not really used.
    ' taskListEnd - End mark of the link lists in the file
    ' link - the link to insert
    Public Sub insertLinksToUI(ByVal filePath As String, _
        ByVal taskListStart As String, ByVal taskListEnd As String, _
        ByVal link As String)

        Dim moduleLinkStart As String = "<!--" & Me.moduleName & ".start-->"
        Dim moduleLinkEnd As String = "<!--" & Me.moduleName & ".end-->"
        Dim hyperLink As String = "<li><a href='" & link & "'>" & Me.moduleName & "</a>"
        Dim startPos As Integer
        Dim endPos As Integer

        Dim content As String
        Dim newContent As String

        If IOManager.fileExists(filePath) Then
            Try
                content = IOManager.GetFileContents(filePath)
                startPos = InStr(LCase(content), LCase(moduleLinkStart)) - 1
                endPos = InStr(LCase(content), LCase(moduleLinkEnd)) - 1
                If startPos > 0 And endPos > 0 Then
                    ' substitute old link for update.
                    newContent = content.Substring(0, startPos) & _
                    moduleLinkStart & hyperLink & moduleLinkEnd & _
                    content.Substring(endPos + Len(moduleLinkEnd))
                    IOManager.SaveTextToFile(newContent, filePath)
                Else
                    ' Insertion new link
                    startPos = InStr(LCase(content), LCase(taskListEnd)) - 1
                    If startPos = 0 Then Exit Sub
                    newContent = content.Substring(0, startPos) & _
                    moduleLinkStart & hyperLink & moduleLinkEnd & vbCrLf & _
                    content.Substring(startPos)
                    IOManager.SaveTextToFile(newContent, filePath)
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

    End Sub


    ' Type 2 IMS functions
    Public Function generateAdmEditFinalFile() As String
        Dim str As String = ""

        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.loginAdmSecFilePath & Chr(34) & "-->"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incFuncFile & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf & vbCrLf

        str = str & vbCrLf & (("<" & "%"))
        str = str & vbCrLf & (("Dim ID, SRC, err"))
        str = str & vbCrLf & (("err = " & Chr(34) & Chr(34)))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("ID = Request.QueryString(" & Chr(34) & "id" & Chr(34) & ")"))
        str = str & vbCrLf & (("SRC = Request.QueryString(" & Chr(34) & "src" & Chr(34) & ")"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & ("If request(" & Chr(34) & "btnSubmit" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call saveSession()")
        str = str & vbCrLf & (vbIndent & "err = checkError()")
        str = str & vbCrLf & (vbIndent & "if err = " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & vbIndent & "call updateData(ID)")
        str = str & vbCrLf & (vbIndent & vbIndent & "Response.Redirect " & Chr(34) & Me.admConfirmFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID ")
        str = str & vbCrLf & (vbIndent & "end if")
        str = str & vbCrLf & ("ElseIf request(" & Chr(34) & "btnCancel" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.admConfirmFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID ")
        str = str & vbCrLf & ("End if")
        str = str & vbCrLf & (("%" & ">"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.headerFilePath & Chr(34) & "-->"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--enter the specific registration information here-->"))
        str = str & vbCrLf & (("<h3>Registration Information</h3>"))        str = str & vbCrLf & (("<!--end the specific registration information-->"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<" & "%"))
        str = str & vbCrLf & ("if err <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & ("Response.Write " & Chr(34) & "<font color=red>Registration Error: <br>" & Chr(34) & " & err & " & Chr(34) & "</font>" & Chr(34)))
        str = str & vbCrLf & ("end if")
        str = str & vbCrLf & (("%" & ">"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<form  method=post name=RegisterForm id=RegisterForm>"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incEditTableFile & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<br>"))
        str = str & vbCrLf & (("<table><tr><td align=left>"))
        str = str & vbCrLf & (("<input type=submit name=btnSubmit value=Submit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (("<input type=reset name=btnReset value=Reset style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("<input type=submit name=btnCancel value=Cancel style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("</td></tr></table>"))
        str = str & vbCrLf & (("</form>"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.footerFilePath & Chr(34) & "-->"))

        Return str
    End Function


    Public Function generateAdmConfirmFinalFile() As String
        Dim str As String = ""

        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.loginAdmSecFilePath & Chr(34) & "-->"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incFuncFile & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & (("<" & "%"))
        str = str & vbCrLf & (("Dim ID, SRC"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("ID = Request.QueryString(" & Chr(34) & "id" & Chr(34) & ")"))
        str = str & vbCrLf & (("SRC = Request.QueryString(" & Chr(34) & "src" & Chr(34) & ")"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & (("call retrieveData(ID)"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & ("if request(" & Chr(34) & "btnEdit" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")

        str = str & vbCrLf & ("'To edit")
        str = str & vbCrLf & ("Response.Redirect " & Chr(34) & Me.admEditFile & "?src=" & Chr(34) & " & SRC & " & Chr(34) & "&id=" & Chr(34) & " & ID ")

        str = str & vbCrLf & ("'To go back")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnBack" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.admListFile & Chr(34))

        str = str & vbCrLf & ("'To approve")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnYesApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call approveData(ID)")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.admListFile & Chr(34))
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnNoApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")

        str = str & vbCrLf & ("'To disapprove")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnDisApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnYesDisApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call disapproveData(ID)")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.admListFile & Chr(34))
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnNoDisApprove" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")

        str = str & vbCrLf & ("'To delete")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnDelete" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnYesDelete" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & (vbIndent & "call deleteData(ID)")
        str = str & vbCrLf & (vbIndent & "call clearSession()")
        str = str & vbCrLf & (vbIndent & "Response.Redirect " & Chr(34) & Me.admListFile & Chr(34))
        str = str & vbCrLf & ("elseif request(" & Chr(34) & "btnNoDelete" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " then")
        str = str & vbCrLf & ("end if")

        str = str & vbCrLf & ("%" & ">")

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.headerFilePath & Chr(34) & "-->"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--enter the specific registration information here-->"))
        str = str & vbCrLf & (("<h3>Registration Information</h3>"))        str = str & vbCrLf & (("<!--end the specific registration information-->"))

        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<form  method=post name=RegisterConfirmForm id=RegisterConfirmForm>"))
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.incFolder & "/" & Me.incConfirmTableFile & Chr(34) & "-->"))
        str = str & vbCrLf & (("<br>"))
        str = str & vbCrLf & vbCrLf

        str = str & vbCrLf & (("<br>"))
        str = str & vbCrLf & (("<table><tr><td align=left>"))
        str = str & vbCrLf & (("<" & "%if request(" & Chr(34) & "btnDelete" & Chr(34) & ")") & " <> " & Chr(34) & Chr(34) & (" then%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<font color=red><b>Are you sure to delete?</b></font>"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnYesDelete value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnNoReset value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("<" & "%elseif request(" & Chr(34) & "btnApprove" & Chr(34) & ")") & " <> " & Chr(34) & Chr(34) & (" then%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<font color=red><b>Are you sure to appove?</b></font>"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnYesApprove value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnNoApprove value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (("<" & "%elseif request(" & Chr(34) & "btnDisapprove" & Chr(34) & ")") & " <> " & Chr(34) & Chr(34) & (" then%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<font color=red><b>Are you sure to disappove?</b></font>"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnYesDisapprove value=Yes style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnNoDisapprove value=No style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (("<" & "%else%" & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnDelete value=Delete style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<" & "%if SRC=" & Chr(34) & "pending" & Chr(34) & "then%" & ">"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnApprove value=Approve style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (vbIndent & ("<" & "%elseif SRC=" & Chr(34) & "approved" & Chr(34) & "then%" & ">"))        str = str & vbCrLf & (vbIndent & vbIndent & ("<input type=submit name=btnDisapprove value=Disapprove style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<" & "%end if%" & ">"))        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnEdit value=Edit style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))
        str = str & vbCrLf & (vbIndent & ("<input type=submit name=btnBack value=Back style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & ">"))        str = str & vbCrLf & (("<" & "%end if%" & ">"))        str = str & vbCrLf & (("</td></tr></table>"))
        str = str & vbCrLf & (("</form>"))
        str = str & vbCrLf & vbCrLf
        str = str & vbCrLf & (("<!--#include file=" & Chr(34) & Me.footerFilePath & Chr(34) & "-->"))

        Return str
    End Function

End Class

