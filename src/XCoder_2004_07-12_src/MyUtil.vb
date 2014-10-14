Public Class MyUtil
    Private myIO As New IOManager()

    Public Shared Function chop(ByVal str As String)
        Dim notfinish = True
        Dim tmpStr As String = str
        ' remove head non-printable characters
        Do While tmpStr.Length > 0 And notfinish
            If (Asc(tmpStr.Substring(0)) < 33 Or Asc(tmpStr.Substring(0)) > 126) Then
                tmpStr = tmpStr.Remove(0, 1)
            Else
                notfinish = False
            End If
        Loop
        ' remove tail non-printable characters
        notfinish = True
        Do While tmpStr.Length > 0 And notfinish
            If (Asc(tmpStr.Substring(tmpStr.Length - 1)) < 33 Or Asc(tmpStr.Substring(tmpStr.Length - 1)) > 126) Then
                tmpStr = tmpStr.Remove(tmpStr.Length - 1, 1)
            Else
                notfinish = False
            End If
        Loop
        Return tmpStr
    End Function


    Public Shared Function getStrAfterDelimiter(ByVal str As String, ByVal delimiter As String)
        Return chop(str.Substring(InStr(str, delimiter)))
    End Function


    ' Return a string which repeats str num times.
    Public Shared Function repeatStr(ByVal str As String, ByVal num As Integer)
        Dim i As Integer
        Dim tmp As String = ""
        For i = 1 To num
            tmp = tmp & str
        Next
        Return tmp
    End Function


    ' Create a string of num that is left filled by zeros, the total length
    ' of the string is len.
    Public Shared Function get0LeftFilledStr(ByVal num As Integer, ByVal len As Integer) As String
        Dim numStr As String = num.ToString
        Dim fillLen = len - numStr.Length
        Dim i As Integer
        For i = 1 To fillLen
            numStr = "0" & numStr
        Next
        Return numStr
    End Function


    ' format a string so it has a fixed length of len, and right appended with
    ' empty string " "
    Public Shared Function leftAlignFormat(ByVal str As String, ByVal len As Integer)
        If str.Length > len Then
            str = str.Substring(0, len)
        Else
            str = str & repeatStr(" ", len - str.Length)
        End If
        Return str
    End Function


    Public Shared Function appendNewLine(ByVal str As String, ByVal newLine As String)
        If newLine = "" Then
            Return str
        ElseIf str = "" Then
            Return newLine
        Else
            Return str & vbCrLf & newLine
        End If
    End Function


    Public Shared Function doubleLine(ByVal len As Integer)
        Return New String("=", len)
    End Function


    Public Shared Function singleLine(ByVal len As Integer)
        Return New String("-", len)
    End Function


    ' Convert an arrayList to string, delimited by delimiter.
    Public Shared Function arrayListToString(ByVal al As ArrayList, ByVal delimiter As String)
        If al Is Nothing Or al.Count = 0 Then
            Return ""
        End If

        Dim i As Integer
        Dim maxIndex As Integer = al.Count - 1
        Dim str As String = ""
        For i = 0 To maxIndex
            If str = "" Then
                str = al(i)
            Else
                str = str & delimiter & al(i)
            End If
        Next
        Return str
    End Function


    Public Shared Function arrayToString(ByVal ar As Array, ByVal delimiter As String)
        If ar Is Nothing Or ar.Length = 0 Then
            Return ""
        End If

        Dim i As Integer
        Dim maxIndex As Integer = ar.Length - 1
        Dim str As String = ""
        For i = 0 To maxIndex
            If str = "" Then
                str = ar(i)
            Else
                str = str & delimiter & ar(i)
            End If
        Next
        Return str
    End Function


    Public Shared Function rmEmptyStrings(ByVal str As String)
        Return Replace(str, " ", "")
    End Function


    Public Shared Function appendToEnd(ByVal str As String, ByVal endStr As String)
        If Not LCase(str).EndsWith(LCase(endStr)) Then
            str = str & endStr
        End If
        Return str
    End Function


    Public Shared Function removeEndString(ByVal str As String, ByVal endStr As String)
        If LCase(str).EndsWith(LCase(endStr)) Then
            str = StrReverse(StrReverse(str).Substring(Len(endStr)))
        End If
        Return str
    End Function


    Public Shared Function stringArrayToArrayList(ByVal ar As String()) As ArrayList
        Dim al As New ArrayList()
        If ar Is Nothing Then Return al

        Dim i, maxI As Integer
        maxI = ar.Length - 1
        For i = 0 To maxI
            al.Add(ar(i))
        Next
        Return al
    End Function


    ' Determine if an arraylist contains a string.
    ' returns index of the string in the arraylist, or -1 if not found
    Public Shared Function arrayListContainsString(ByVal al As ArrayList, _
           ByVal str As String, ByVal caseSensitive As Boolean) As Integer
        Dim i As Integer
        Dim maxI As Integer = al.Count - 1
        For i = 0 To maxI
            If caseSensitive Then
                If CType(al(i), String) = str Then Return i
            Else
                If UCase(CType(al(i), String)) = UCase(str) Then Return i
            End If
        Next
        Return -1
    End Function


    ' http://www.codeproject.com/useritems/getControlFromName.asp
    Public Shared Function getControlFromName(ByRef containerObj As Object, _
                          ByVal name As String) As Control
        Try
            Dim tempCtrl As Control
            For Each tempCtrl In containerObj.Controls
                If tempCtrl.Name.ToUpper.Trim = name.ToUpper.Trim Then
                    Return tempCtrl
                End If
            Next tempCtrl
        Catch ex As Exception
        End Try
    End Function


    Public Shared Function arrayToArrayList(ByRef array As String()) As ArrayList
        Dim i As Integer
        Dim maxI As Integer = array.Length - 1 'UBound(array) - 1
        Dim al As New ArrayList()
        For i = 0 To maxI
            al.Add(array(i))
        Next
        Return al
    End Function
End Class
