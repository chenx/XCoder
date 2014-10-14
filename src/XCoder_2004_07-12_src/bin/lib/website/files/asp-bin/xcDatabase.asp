<%
''''''''''''''''''''''''''''''''''''''''''''''''''''
' 1. Basic database functions.
' 2. Extended database functions.
' 3. String Encode Functions.
' 4. String functions.
' 5. Math functions.
' 6. Date functions.
''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''
' 1. Basic database functions.
''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetDSN()
	GetDSN = "Provider=SQLOLEDB;Data Source=" & db_db & ";User ID=" & db_login & ";Password=" & db_password & ";"
End Function


Function dbConnect()
	Dim db
	Set db = Server.CreateObject("ADODB.Connection")
	db.ConnectionTimeout = 1200
	db.Open db_db, db_login, db_password
	Set dbConnect = db
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' Execute an SQL query and return recordset object.
''''''''''''''''''''''''''''''''''''''''''''''''''''
Function dbExecuteRS(byRef db, sql_command)
	Dim record
	Set record = Server.CreateObject("ADODB.Recordset")
	 record.Open sql_command, db, adOpenStatic, adLockReadOnly, adCmdText
	Set dbExecuteRS = record
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' Execute an SQL non-query.
''''''''''''''''''''''''''''''''''''''''''''''''''''
Function dbExecute(byRef db, sql_command)
'	on error resume next
	db.execute sql_command, , adCmdText
'	if not err.number=0 then
'		response.redirect "error.asp?error=" & Server.URLEncode(err.Description)
'	end if
End Function


Function dbRSClose (byRef db, byRef rs)
	rs.close()
	set rs = nothing
	db.close()
	set db = nothing
End Function


Function dbClose(byRef db)
	db.Close()
	set db = nothing
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2. Extended database functions.
''''''''''''''''''''''''''''''''''''''''''''''''''''


function getFieldBySQL(sql, field)
	Dim db, rs
	set db = dbConnect()
	set rs = dbExecute(db, sql)
	getField = rs(field)
	call closeDBRS(db, rs)
end function


function getFieldsBySQL(sql, field1, field2, delimiter)
	Dim db, rs
	set db = dbConnect()
	set rs = dbExecute(db, sql)
	getFields = rs(field1) & delimiter & rs(field2)
	call closeDBRS(db, rs)
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
	call closeDBRS(db, rs)
end function


'''
' Update a table field
'''
sub updateTableByID(table, fieldName, value, ID)
	dim db, rs, sql
	sql = "update " & table & " set " & fieldName & " = '" & value & "' where ID = " & ID

	set db = dbConnect()
	call dbExecute(db, sql)
	
	call db.close()
	set db = nothing
end sub


' Check if a field already exist in a table.
function fieldExist(table, fieldName, fieldValue)
	dim db, rs, sql

	sql = "SELECT * FROM " & table & " WHERE " & fieldName & " = '" & fieldValue  & "'"
	set db = dbConnect()
	set rs = dbExecuteRS(db, sql)
	
	if rs.eof then
		fieldExist = false
	else
		fieldExist = true
	end if	
	
	call dbRSClose(db, rs)
end function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' 3. String Encode Functions.
''''''''''''''''''''''''''''''''''''''''''''''''''''


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


''''''''''''''''''''''''''''''''''''''''''''''''''''
' change "http://..." into hyperlink.
' First divide the string into lines, then divide line into words, 
' and change those words starts with "http://" into hyperlink
' Right now support image format: .JPG, .GIF
''''''''''''''''''''''''''''''''''''''''''''''''''''
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


''''''''''''''''''''''''''''''''''''''''''''''''''''
' For use by Text field in View and Edit mode
''''''''''''''''''''''''''''''''''''''''''''''''''''
Function TextHTMLEncode(html_string)
	Dim result
	If html_string <> Null OR html_string <> "" Then
		result = html_string
		result = replace(result, "&", "&amp;") 
  
		result = Replace(result, "<", "&lt;")
		result = Replace(result, ">", "&gt;")
		 
		result = Replace(result, """", "&quot;") ' Useful for text field in Edit mode only
		TextHTMLEncode = result
	else
		TextHTMLEncode = ""
	End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' For use by TextArea field in Edit mode
''''''''''''''''''''''''''''''''''''''''''''''''''''
Function TextAreaEditHTMLEncode(html_string)
	Dim result
	If html_string <> Null OR html_string <> "" Then
		result = html_string
		result = replace(result, "&", "&amp;") 
  
		result = Replace(result, "<", "&lt;")
		result = Replace(result, ">", "&gt;")
  
		TextAreaEditHTMLEncode = result
	else
		TextareaEditHTMLEncode = ""
	End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' For use by TextArea field in View mode
''''''''''''''''''''''''''''''''''''''''''''''''''''
Function TextAreaViewHTMLEncode(html_string)
	Dim result
	If html_string <> Null OR html_string <> "" Then
		result = html_string
		result = replace(result, "&", "&amp;") 
  
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


''''''''''''''''''''''''''''''''''''''''''''''''''''
' 4. String functions.
''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Same as mid, return the substring of "str" starting at position "pos"
''''''''''''''''''''''''''''''''''''''''''''''''''''
Function xcMid(str, pos)
	Dim tmpStr 
	tmpStr = str
	xcMid = replace(tmpStr, left(tmpStr, pos), "", 1, 1)
end function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' Compares two strings and return true only if
' they are case sensitively equal to each other.
''''''''''''''''''''''''''''''''''''''''''''''''''''
function caseSensitiveEqual(str1, str2)
	if len(str1) = len(str2) then
		length = len(str1)
		i = 1
		do while i <= length
			if asc( Mid(str1, i, 1) ) <> asc( Mid(str2, i, 1) ) then
				'Response.Write "dif pos = " & i & "<br>"
				caseSensitiveEqual = false
				exit function
			end if
			i = i + 1
		loop
		caseSensitiveEqual = true
	else
		caseSensitiveEqual = false
	end if
end function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' 5. Math functions.
''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''
' generate a random number between 1 and intHighestNumber
''''''''''''''''''''''''''''''''''''''''''''''''''''
Function RandomNumber(intHighestNumber)
	Randomize
	RandomNumber = Int(Rnd * intHighestNumber) + 1
End Function


function IsOdd(byVal hInput)
	' remove anything after the decimal point and the decimal point if found
	' because we need a whole number to use Mod
	If InStr(hInput, ".") <> 0 Then hInput = Left(hInput, InStrRev(hInput, ".") - 1)

	' If not numeric, exit with a Null value
	If Not IsNumeric(hInput) Then IsOdd = Null: Exit Function

	' use the Mod function to get the remainder of dividing the input by 2. 
	' Forcing a 1 into a boolean value will make it True and a 0 becomes False
	IsOdd = CBool( hInput Mod 2 )
end function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' 6. Date functions.
''''''''''''''''''''''''''''''''''''''''''''''''''''


function ShortDateTime(fullDate)
	if isdate(fullDate) then
		ShortDateTime = formatdatetime(fullDate, vbshortdate) & "  "  & _
						formatdatetime(fullDate, vbshorttime)
	else
		ShortDateTime = fullDate
	end if
end function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' 7. Miscellaneous functions.
''''''''''''''''''''''''''''''''''''''''''''''''''''


function writeError(errInfo)
	Response.Write "<font color=red>" & errInfo & "</font><br>"
end function

%>