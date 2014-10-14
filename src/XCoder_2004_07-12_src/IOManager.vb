Imports System.IO

Public Class IOManager

    ' Get file contents.
    Public Shared Function GetFileContents(ByVal FullPath As String, _
           Optional ByRef ErrInfo As String = "") As String

        Dim strContents As String
        Dim objReader As StreamReader
        Try
            objReader = New StreamReader(FullPath)
            strContents = objReader.ReadToEnd()
            objReader.Close()
            Return strContents
        Catch Ex As Exception
            ErrInfo = Ex.Message
        End Try
    End Function


    ' Create a new file and write to it.
    Public Shared Function SaveTextToFile(ByVal strData As String, _
        ByVal FullPath As String, _
        Optional ByRef ErrInfo As String = "") As Boolean

        Dim bAns As Boolean = False
        Dim objWriter As StreamWriter
        Try
            objWriter = New StreamWriter(FullPath)
            objWriter.Write(strData)
            objWriter.Close()
            bAns = True
        Catch Ex As Exception
            ErrInfo = Ex.Message
        End Try
        Return bAns
    End Function


    ' Move a file.
    Public Shared Function Move(ByVal srcFullPath As String, ByVal destFullPath As String, _
        Optional ByRef ErrInfo As String = "") As Boolean

        Dim bAns As Boolean = False
        Try
            System.IO.File.Move(srcFullPath, destFullPath)
            bAns = True
        Catch Ex As Exception
            ErrInfo = Ex.Message
        End Try
        Return bAns
    End Function


    ' Delete a file
    Public Shared Function DeleteFile(ByVal FullPath As String, _
        Optional ByRef ErrInfo As String = "") As Boolean

        Dim bAns As Boolean = False
        Try
            System.IO.File.Delete(FullPath)
            bAns = True
        Catch Ex As Exception
            ErrInfo = Ex.Message
        End Try
        Return bAns
    End Function


    ' Append text to the end of a file
    Public Shared Function appendTextToFile(ByVal FullPath As String, _
        ByVal strData As String, Optional ByRef ErrInfo As String = "") As Boolean
        Dim bAns As Boolean = False
        Try
            If Not System.IO.File.Exists(FullPath) Then
                SaveTextToFile("", FullPath)
            End If
            Dim objWriter As StreamWriter
            objWriter = System.IO.File.AppendText(FullPath)
            objWriter.Write(strData)
            objWriter.Close()
            bAns = True
        Catch ex As Exception
            ErrInfo = ex.Message
        End Try
        Return bAns
    End Function


    ' Get file length
    Public Shared Function GetFileLength(ByVal FullPath As String, _
           Optional ByRef ErrInfo As String = "") As Long
        Dim strLength As Long
        Dim objReader As StreamReader
        Try
            objReader = New StreamReader(FullPath)
            strLength = objReader.ReadToEnd().Length
            objReader.Close()
            Return strLength
        Catch Ex As Exception
            ErrInfo = Ex.Message
        End Try
    End Function


    Public Shared Function CreateFolder(ByVal FullPath As String, _
           Optional ByRef ErrInfo As String = "") As Boolean
        Dim bAns As Boolean = False
        Try
            System.IO.Directory.CreateDirectory(FullPath)
            bAns = True
        Catch ex As Exception
            ErrInfo = ex.Message
        End Try
        Return bAns
    End Function


    Public Shared Function DeleteFolder(ByVal FullPath As String, _
           Optional ByVal recursive As Boolean = True, _
           Optional ByRef ErrInfo As String = "") As Boolean
        Dim bAns As Boolean = False
        Try
            System.IO.Directory.Delete(FullPath, recursive)
            bAns = True
        Catch ex As Exception
            ErrInfo = ex.Message
        End Try
        Return bAns
    End Function


    Public Shared Function folderExists(ByVal FullPath As String, _
            Optional ByRef ErrInfo As String = "") As Boolean
        Dim bAns As Boolean = False
        Try
            If System.IO.Directory.Exists(FullPath) Then
                bAns = True
            End If
        Catch ex As Exception
            ErrInfo = ex.Message
        End Try
        Return bAns
    End Function


    Public Shared Function fileExists(ByVal FullPath As String, _
            Optional ByRef ErrInfo As String = "") As Boolean
        Dim bAns As Boolean = False
        Try
            If System.IO.File.Exists(FullPath) Then
                bAns = True
            End If
        Catch ex As Exception
            ErrInfo = ex.Message
        End Try
        Return bAns
    End Function


    Public Shared Function getFileAsLines(ByVal filePath As String, _
        ByRef lines As String(), Optional ByRef errInfo As String = "") As Boolean

        Dim str As String = IOManager.GetFileContents(filePath, errInfo)

        If errInfo <> "" Then
            Return False
        End If

        lines = Split(str, Chr(10))
        Dim i As Integer
        Dim maxI As Integer = lines.Length - 1
        For i = 0 To maxI
            lines(i) = MyUtil.chop(lines(i))
        Next

        Return True
    End Function


    Public Shared Function isValidFileName(ByVal str As String, ByVal showWarning As Boolean) As Boolean
        If InStr(str, "\") > 0 Or InStr(str, "/") > 0 Or _
           InStr(str, ":") > 0 Or InStr(str, "*") > 0 Or _
           InStr(str, "?") > 0 Or InStr(str, Chr(34)) > 0 Or _
           InStr(str, "<") > 0 Or InStr(str, ">") > 0 Or _
           InStr(str, "|") > 0 Then
            If showWarning Then MessageBox.Show("A file name cannot contain any of the following characters: \/:*?" & Chr(34) & "<>|")
            Return False
        End If
        Return True
    End Function


    ' Given full path C:xx\xx\...\y.z, returns the folder part
    Public Shared Function getFolder(ByVal fileFullPath As String) As String
        Dim strRev As String = StrReverse(fileFullPath)
        Dim folder As String = StrReverse(strRev.Substring(InStr(strRev, "\")))
        If Not folder.EndsWith("\") Then
            folder += "\"
        End If
        Return folder
    End Function


    Public Shared Function getFile(ByVal fileFullPath As String) As String
        Dim folder As String = getFolder(fileFullPath)
        Dim file As String = Replace(fileFullPath, folder, "", 1, 1)
        Return file
    End Function


    Public Shared Function formatFileName(ByVal fileName As String, _
        ByVal fileExtension As String, ByVal withExtension As Boolean) As String

        If withExtension Then
            If LCase(fileName).EndsWith(LCase(fileExtension)) Then Return fileName
            Return fileName & fileExtension
        Else
            If Not LCase(fileName).EndsWith(LCase(fileExtension)) Then Return fileName
            Return StrReverse(StrReverse(fileName).Substring(Len(fileExtension)))
        End If
    End Function

End Class
