<%

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function library
' Author: Xin Chen
' Date: Tuesday. March 16, 2004
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''
' Update a table field
'''
sub updateTableByID(tableName, fieldName, value, ID)
	dim db, rs, sql
	sql = "update " & tableName & " set " & fieldName & " = '" & value & "' where ID = " & ID

	set db = dbConnect()
	call dbExecuteQ(db, sql)
	
	call db.close()
	set db = nothing
end sub


' Check if a field already exist in a table.
function fieldExist(table, fieldName, fieldValue)
	dim db, rs, sql

	sql = "SELECT * FROM " & table & " WHERE " & fieldName & " = '" & fieldValue  & "'"
	set db = dbConnect()
	set rs = dbExecuteFO(db, sql)
	
	if rs.eof then
		fieldExist = false
	else
		fieldExist = true
	end if	
	
	call dbClose(db, rs)
end function


' Function getDropdownList(sql, listName, valueFld, ItemFldArray, delimiter)
' 3-15-2004
' This function constructs a dropdown list using data retrieved from "sql" arg.
' The name of the dropdown list is "listName"
' The value of each item is "valueFld"
' The displayed text of each item is combination of fields from "itemArray" 
'   separated by "delimiter"
' THomas. 3-15-04
' e.g.
' sql_ = "SELECT * FROM members where approval = 1 and isAdmin = 0"
' dim itemArray(2)
' itemArray(1) = "lastName"
' itemArray(2) = "firstName"
' Response.Write  getDropdownList(sql_, "testList", "ID", itemArray, ", ")
function getDropdownList(sql, listName, valueFld, ItemFldArray, delimiter)
	dim db, rs, listStr
	listStr = ""
	
	set db = dbConnect()
	set rs = dbExecuteFO(db, sql)
	
	if not rs.eof then
		listStr = chr(13) & "<select name='" & listName & "'>" & chr(13) 
		
		do while not rs.eof
			listStr = listStr & "<option value=" & rs(valueFld) & ">"
			if isArray(ItemFldArray) then
				count_ = ubound(ItemFldArray)
				for i = 1 to count_ 
					listStr = listStr & rs(ItemFldArray(i)) & delimiter
				next
			else
				listStr = listStr & rs(ItemFldArray)
			end if
			listStr = listStr & chr(13)
			rs.moveNext()
		loop
		listStr = listStr & "</select>" & chr(13)
	end if

	call dbClose(db, rs)

	getDropdownList = listStr
end function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function compares two strings and return true only if
' they are case sensitively equal to each other.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function case_sensitive_equal(str1, str2)
	'Response.Write "str1/str2 = " & str1 & "/" & str2 & "<br>"
	if len(str1) = len(str2) then
		length = len(str1)
		i = 1
		do while i <= length
			'Response.Write mid(str1, i, 1)
			if asc( Mid(str1, i, 1) ) <> asc( Mid(str2, i, 1) ) then
				'Response.Write "dif pos = " & i & "<br>"
				case_sensitive_equal = false
				exit function
			end if
			i = i + 1
		loop
		case_sensitive_equal = true
	else
		case_sensitive_equal = false
	end if
end function

%>