Public Class xcUtil
    Public Shared WebsiteFrameworkModuleName As String = "Website Framework"

    Public Shared vbIndent As String = "    "
    Public Shared sqlIndent As String = "    "
    Public Shared xmlIndent As String = "    "

    Public Shared admTaskListStart As String = "<!--admTaskList.Start-->"
    Public Shared admTaskListEnd As String = "<!--admTaskList.End-->"
    Public Shared memTaskListStart As String = "<!--memTaskList.Start-->"
    Public Shared memTaskListEnd As String = "<!--memTaskList.End-->"

    ' Used to mark a Edit/Confirm field that cannot be empty.
    ' E.g. login/password/passwordRpt fields, or other non-empty fields.
    Public Shared notEmptyMark As String = "&nbsp;<font color=red>*</font>&nbsp;"

    Public Shared Function writeFunctionComment(ByVal comment As String)
        Dim str As String = ""
        str = str & vbCrLf & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & "' " & comment & vbCrLf
        str = str & "'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf
        str = str & vbCrLf
        Return str
    End Function


    Public Shared Function getTableAttributes( _
        ByRef tblAttr As clsHTMLTableAttributes, ByRef tableStart As String, _
        ByRef firstColWidth As String, ByRef secondColWidth As String)

        tableStart = "<table"
        If Not tblAttr Is Nothing Then
            If tblAttr.border <> "" Then tableStart &= " border=" & tblAttr.border
            If tblAttr.width <> "" Then tableStart &= " width=" & tblAttr.width
            If tblAttr.bgColor <> "" Then tableStart &= " bgcolor=" & tblAttr.bgColor
            If tblAttr.cellPadding <> "" Then tableStart &= " cellpadding=" & tblAttr.cellPadding
            If tblAttr.cellSpacing <> "" Then tableStart &= " cellspacing=" & tblAttr.cellSpacing
            If tblAttr.firstColWidth <> "" Then firstColWidth = " width=" & tblAttr.firstColWidth
            If tblAttr.width <> "" And tblAttr.firstColWidth <> "" Then
                Try
                    Dim _tableWidth As Integer = CType(tblAttr.width, Integer)
                    Dim _firstColWidth As Integer = CType(tblAttr.firstColWidth, Integer)
                    If _tableWidth <= _firstColWidth Then
                        secondColWidth = ""
                    Else
                        secondColWidth = " width=" & CStr(_tableWidth - _firstColWidth)
                    End If
                Catch ex As Exception
                End Try
            End If
        End If
        tableStart &= ">"
    End Function

End Class
