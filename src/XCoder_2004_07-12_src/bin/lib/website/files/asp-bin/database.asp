<!--#include file="adovbs.inc"-->
<!--#include file="database_config.asp"-->
<!--#include file="utilFuncsLib.asp"-->

<%

Function GetDSN()
	GetDSN = "Provider=SQLOLEDB;Data Source=" & db_db & ";User ID=" & db_login & ";Password=" & db_password & ";"
End Function

'********
' Connect to database.
'********
Function dbConnect()
	Dim db
	Set db = Server.CreateObject("ADODB.Connection")
	on error resume next
	db.ConnectionTimeout = 1200
	db.Open db_db,db_login,db_password
	if not err.number=0 then
		response.redirect "/asp-bin/error.asp?error=" & Server.URLEncode(err.Description)
	end if
	Set dbConnect = db
End Function

'********
' Execute an SQL query and return the results.
'********
Function dbExecute(db, sql_command)
	Dim record
	Set record = Server.CreateObject("ADODB.Recordset")
'Response.Write "sql="&sql_command&"<br>"
	record.Open sql_command, db, adOpenKeyset, adLockOptimistic, adCmdText
	'record.Open sql_command, db, 1 , 3, &H0001
	'record.Open sql_command,db
	rem record.Open sql_command, db, 2, 3, &H0001
	Set dbExecute = record
End Function


Function dbExecuteFO(db, sql_command)
'Response.Write sql_command & "<BR>"
	Dim record
	Set record = Server.CreateObject("ADODB.Recordset")

	'record.Open sql_command, db, adOpenForwardonly, adLockReadOnly, adCmdText
	 record.Open sql_command, db, adOpenStatic, adLockReadOnly, adCmdText
	Set dbExecuteFO = record
End Function


' Check if there are any error in executing.
Function dbExecuteQ(db, sql_command)
'Response.Write sql_command & "<BR>"
	db.execute sql_command, , adCmdText
	on error resume next
	if not err.number=0 then
		response.redirect "error.asp?error=" & Server.URLEncode(err.Description)
	end if

End Function



Function dbDynamicRecord(db, sql)
	'Response.Write sql 
	Dim record
	Set record = Server.CreateObject("ADODB.Recordset")
	record.Open sql, db, adOpenKeyset, adLockBatchOptimistic, adCmdText
	rem record.Open sql_command, db, 2, 3, &H0001
	Set dbDynamicRecord = record
End Function


'********
' Escape an SQL parameter (single quotes).
'********
Function numEncode(sql_string)
	if IsEmpty(sql_string) or IsNull(sql_string) or trim(sql_string)="" then
		numEncode = "null"
	else
		numEncode = sql_string
	end if
End Function

Function dbEncode(sql_string)
	if not IsNull(sql_string) then
		dbEncode = Replace(sql_string, "'", "''")
		dbEncode = Replace(dbEncode, "‘", "''")
		dbEncode = Replace(dbEncode, "’", "''")
	end if
End Function


Function SqlEncode (sql_string)
	if trim(sql_string)="" then
		SqlEncode = "null"
	else
		SqlEncode = "'" & dbEncode (trim(sql_string)) & "'"
	end if
End Function

Function SqlEncodeWild (sql_string)
	if trim(sql_string)="" then
		SqlEncodeWild = "null"
	else
		SqlEncodeWild = sqlEncode ("%" & trim(sql_string) & "%")
	end if
End Function


' change "http://..." into hyperlink.
' First divide the string into lines by chr(34), 
' then divide each line into words by " ", change those starts with "http://"
' Right now support image format: .JPG, .GIF
Function HyperlinkEncode(html_string)
	Dim inputStr, outputStr
	Dim line, maxLine, maxI, i, tmpStr, tmpResult, tmpLinesArray, tmpArry
	if html_string <> NULL or html_string <> "" then
		inputStr = html_string
		
		if instr(1, inputStr, "http://") > 0 then
			outputStr = ""
			tmpResult = replace(inputStr, chr(13), "")
			tmpLinesArray = split(tmpResult, chr(10))
			
			maxLine = ubound(tmpLinesArray)
			for line = 0 to maxLine
			
				tmpArray = split(tmpLinesArray(line), " ")
				maxI = ubound(tmpArray)
			
				for i = 0 to maxI
					tmpstr = tmpArray(i)
					if instr(1, lcase(tmpstr), "http://") = 1 then
						if lcase( right(tmpStr, 4) ) = ".jpg" or _
						   lcase( right(tmpStr, 4) ) = ".gif" then
							tmpstr = "<img src='" & tmpStr & "'>" '& "<br>[<a href='" & tmpStr & "'>" & tmpStr & "</a>]"
						else
							tmpstr = "<a href='" & tmpStr & "' target=_blank>" & tmpStr & "</a>"
						end if
					end if
					outputStr = outputStr & tmpStr & " "
				next
				outputStr = outputStr & chr(10) & chr(13)
			next
			
			HyperLinkEncode = outputStr
		else
			HyperLinkEncode = inputStr
		end if
	else
		HyperLinkEncode = ""
	end if
end function


' For use by Text field in View and Edit mode
Function HTMLEncode(html_string)
	Dim result
	If html_string <> Null OR html_string <> "" Then
		result = html_string
		result = replace(result, "&", "&amp;") ' to keep escape codes.
  
		result = Replace(result, "<", "&lt;")
		result = Replace(result, ">", "&gt;")
		 
		result = Replace(result, """", "&quot;") ' Useful for text field in Edit mode only
		result = Replace(result, Chr(10), "<BR>")	' Useful for TextArea field in View mode only
		result = Replace(result, Chr(13), "")		' Useful for TextArea field in View mode only
		HTMLEncode = result
	else
		HTMLEncode = ""
	End If
End Function


' For use by TextArea field in Edit mode
Function TextAreaHTMLEncode(html_string)
	Dim result
	If html_string <> Null OR html_string <> "" Then
		result = html_string
		result = replace(result, "&", "&amp;") ' to keep escape codes.
  
		result = Replace(result, "<", "&lt;")
		result = Replace(result, ">", "&gt;")
  
		TextAreaHTMLEncode = result
	else
		TextareaHTMLEncode = ""
	End If
End Function


' For use by TextArea field in View mode
Function TextAreaViewHTMLEncode(html_string)
	Dim result
	If html_string <> Null OR html_string <> "" Then
		result = html_string
		result = replace(result, "&", "&amp;") ' to keep escape codes.
  
		result = Replace(result, "<", "&lt;")
		result = Replace(result, ">", "&gt;")
		
		result = HyperlinkEncode(result)
  
		result = Replace(result, Chr(10), "<BR>")	' Useful for TextArea field in View mode only
		result = Replace(result, Chr(13), "")		' Useful for TextArea field in View mode only
		TextAreaViewHTMLEncode = result
	else
		TextAreaViewHTMLEncode = ""
	End If
end function


'********
' Create recordset and initialize from filter
'********
function dbTableFilter(db, tablename, filter)
	dim rst, sql
	sql = "select * from " & tablename & " where " & filter
	'response.write "dbTableFilter: " & sql & "<p>"
	set rst = dbExecuteFO(db, sql)
	set dbTableFilter = rst
end function

'********
' Return a dictionary with contents of record
'********
function dbGet(tablename, filter, dict)
	dim db, rst, field
	on error resume next
	set db = dbConnect
	set rst = dbTableFilter(db, tablename, filter)
	if rst.EOF then
		dbGet = "dbGet: Record not found"
		exit function
	end if
	for each field in rst.Fields
		if not IsNull(field) then
			dict(field.Name) = field.Value
		end if
	next 
	call dbclose(db,rst)
end function

'********
'Execute SQL StoredProcedure using ADO Command Object Utils
'********
Sub collectParams(cmd, argparams)
    Dim params 
    Dim I, V
	For I = LBound(argparams) To UBound(argparams)
       cmd.Parameters.Append cmd.CreateParameter(argparams(I)(0), argparams(I)(1), adParamInput, argparams(I)(2), argparams(I)(3))
    Next 
End Sub

'********
'Execute a StoredProcedure to Return a disconnected Dynamic recordset
'********
Function RunSPWithRS_RW(strSP, params)
    ' Set up Command and Connection objects
    Dim rs 
    Dim cmd 
    Set rs = Server.CreateObject ("ADODB.Recordset")
    Set cmd = Server.CreateObject ("ADODB.Command")
       
    'Run the procedure
    cmd.ActiveConnection = GetDSN()
    cmd.CommandText = strSP
    cmd.CommandType = adCmdStoredProc
    
    collectParams cmd, params
    rs.CursorLocation = adUseClient
    On error resume next
    rs.Open cmd, , adOpenDynamic, adLockBatchOptimistic
    On error goto 0

	Set cmd.ActiveConnection = Noting
    Set cmd = Nothing
    Set rs.ActiveConnection = Nothing
    
 	if not err.number=0 then
		setRunSPWithRS_RW = nothing
		Response.Clear 
		response.redirect "/asp-bin/error.asp?module=" & Server.URLEncode ("database.runspwithrs_rw") & "&error=" & Server.URLEncode(err.Description) 	
	else
		Set RunSPWithRS_RW = rs
	end if    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Close a recordset and db.
' Note:
' The parameter passing is ByRef in default, so this function
' actually closes the rs and db and sets them to nothing!
' Test can be done by checking the error message generated by this:
' 	if rs.eof then
'	end if
' after calling dbClose, and compare the two cases using 
' sub dbClose(ByVal db, ByVal rs) and sub dbClose(ByRef db, ByRef rs).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub dbClose (db, rs)
	rs.close()
	set rs=nothing
	db.close()
	set db=nothing
end sub


' see if a number is odd
function IsOdd(byVal hInput)
	' remove anything after the decimal 
	' point and the decimal point if found
	' because we need a whole number to
	' use Mod
	If InStr(hInput, ".") <> 0 Then hInput = Left(hInput, InStrRev(hInput, ".") - 1)

	' If not numeric, exit with a Null value
	If Not IsNumeric(hInput) Then IsOdd = Null: Exit Function

	' use the Mod function to get the remainder
	' of dividing the input by 2. If it's 0, the
	' number is even. If it's 1, the number is odd.
	' Forcing a 1 into a boolean value will make it True
	' and a 0 becomes False
	IsOdd = CBool( hInput Mod 2 )
end function


function getField(sql, field)
	Dim db, rs, result
	set db = dbConnect
	set rs = dbExecute(db, sql)
	result = rs(field)
	call dbClose(db, rs)
	
	getField = result
end function


function getMemFullName(memID)
	getMemFullName = getFieldByID("members", "firstName", "ID", memID) & " " & _
				  	 getFieldByID("members", "lastName", "ID", memID)
end function


function getFields(sql, field1, field2, delimiter)
	Dim db, rs, result
	set db = dbConnect
	set rs = dbExecute(db, sql)
	result = rs(field1) & delimiter & rs(field2)
	getFields = result
	call dbClose(db, rs)
end function


' returns the first field satisfying the SQL requirement
function getFieldByID(table, fieldName, ID, IDValue)
	Dim db, rs, sql
	sql = "SELECT " & fieldName & " FROM " & table & " WHERE " & ID & " = '" & IDValue & "'"
	set db = dbConnect()
	set rs = dbExecute(db, sql)
	if not rs.EOF then
		getFieldByID = rs(fieldName)
	else
		getFieldByID = ""
	end if
	call dbClose(db, rs)
end function


function YYYYMMDD(fullDate)
	if isdate(fullDate) then
		YYYYMMDD = year(fullDate) & "/" & month(fulldate) & "/" & day(fulldate)
	else
		YYYYMMDD = fullDate
	end if
end function


function ShortDateTime(fullDate)
	if isdate(fullDate) then
		ShortDateTime = formatdatetime(fullDate, vbshortdate) & "  "  & _
						formatdatetime(fullDate, vbshorttime)
	else
		ShortDateTime = fullDate
	end if
end function


function writeError(preLabel, errInfo)
	Response.Write "<font color=red>" & preLabel & errInfo & "</font><br>"
end function


' generate a random number between 1 and intHighestNumber
Function RandomNumber(intHighestNumber)
	Randomize
	RandomNumber = Int(Rnd * intHighestNumber) + 1
End Function


' Same as mid, return the substring of "str" starting at position "pos"
Function xcMid(str, pos)
	Dim tmpStr 
	tmpStr = str
	xcMid = replace(tmpStr, left(tmpStr, pos), "", 1, 1)
end function
%>