Public Class clsWebsiteFrameworkParameters
    ' This class contains global parameters for a website 

    Private _homeFile As String
    Private _cssFile As String
    Private _rootIncFolder As String
    Private _rootIncHeaderFile As String
    Private _rootIncFooterFile As String
    Private _admFolder As String
    Private _admHomeFile As String
    Private _memFolder As String
    Private _memHomeFile As String
    Private _loginFolder As String
    Private _loginAdmSecurityFile As String
    Private _loginMemSecurityFile As String
    Private _aspBinFolder As String
    Private _aspFuncLibFile As String
    Private _imageFolder As String


    Public Property homeFile() As String
        Get
            Return Me._homeFile
        End Get
        Set(ByVal Value As String)
            Me._homeFile = Value
        End Set
    End Property


    Public Property cssFile() As String
        Get
            Return Me._cssFile
        End Get
        Set(ByVal Value As String)
            Me._cssFile = Value
        End Set
    End Property


    Public Property rootIncFolder() As String
        Get
            Return Me._rootIncFolder
        End Get
        Set(ByVal Value As String)
            Me._rootIncFolder = Value
        End Set
    End Property


    Public Property rootIncHeaderFile() As String
        Get
            Return Me._rootIncHeaderFile
        End Get
        Set(ByVal Value As String)
            Me._rootIncHeaderFile = Value
        End Set
    End Property


    Public Property rootIncFooterFile() As String
        Get
            Return Me._rootIncFooterFile
        End Get
        Set(ByVal Value As String)
            Me._rootIncFooterFile = Value
        End Set
    End Property


    Public Property admFolder() As String
        Get
            Return Me._admFolder
        End Get
        Set(ByVal Value As String)
            Me._admFolder = Value
        End Set
    End Property


    Public Property admHomeFile() As String
        Get
            Return Me._admHomeFile
        End Get
        Set(ByVal Value As String)
            Me._admHomeFile = Value
        End Set
    End Property


    Public Property memFolder() As String
        Get
            Return Me._memFolder
        End Get
        Set(ByVal Value As String)
            Me._memFolder = Value
        End Set
    End Property


    Public Property memHomeFile() As String
        Get
            Return Me._memHomeFile
        End Get
        Set(ByVal Value As String)
            Me._memHomeFile = Value
        End Set
    End Property


    Public Property loginFolder() As String
        Get
            Return Me._loginFolder
        End Get
        Set(ByVal Value As String)
            Me._loginFolder = Value
        End Set
    End Property


    Public Property loginAdmSecurityFile() As String
        Get
            Return Me._loginAdmSecurityFile
        End Get
        Set(ByVal Value As String)
            Me._loginAdmSecurityFile = Value
        End Set
    End Property


    Public Property loginMemSecurityFile() As String
        Get
            Return Me._loginMemSecurityFile
        End Get
        Set(ByVal Value As String)
            Me._loginMemSecurityFile = Value
        End Set
    End Property


    Public Property aspBinFolder() As String
        Get
            Return Me._aspBinFolder
        End Get
        Set(ByVal Value As String)
            Me._aspBinFolder = Value
        End Set
    End Property


    Public Property aspFuncLibFile() As String
        Get
            Return Me._aspFuncLibFile
        End Get
        Set(ByVal Value As String)
            Me._aspFuncLibFile = Value
        End Set
    End Property


    Public Property imageFolder() As String
        Get
            Return Me._imageFolder
        End Get
        Set(ByVal Value As String)
            Me._imageFolder = Value
        End Set
    End Property


    ' Prefix is used for customization by calling procedure.
    Public Function writeParams(ByVal prefix As String) As String
        Dim str As String = ""
        str = str & prefix & "FrameworkHomeFile=" & Me.homeFile & vbCrLf
        str = str & prefix & "FrameworkCssFile=" & Me.cssFile & vbCrLf
        str = str & prefix & "FrameworkRootIncFolder=" & Me.rootIncFolder & vbCrLf
        str = str & prefix & "FrameworkRootIncHeaderFile=" & Me.rootIncHeaderFile & vbCrLf
        str = str & prefix & "FrameworkRootIncFooterFile=" & Me.rootIncFooterFile & vbCrLf
        str = str & prefix & "FrameworkAdmFolder=" & Me.admFolder & vbCrLf
        str = str & prefix & "FrameworkAdmHomeFile=" & Me.admHomeFile & vbCrLf
        str = str & prefix & "FrameworkMemFolder=" & Me.memFolder & vbCrLf
        str = str & prefix & "FrameworkMemHomeFile=" & Me.memHomeFile & vbCrLf
        str = str & prefix & "FrameworkLoginFolder=" & Me.loginFolder & vbCrLf
        str = str & prefix & "FrameworkLoginAdmSecurityFile=" & Me.loginAdmSecurityFile & vbCrLf
        str = str & prefix & "FrameworkLoginMemSecurityFile=" & Me.loginMemSecurityFile & vbCrLf
        str = str & prefix & "FrameworkAspBinFolder=" & Me.aspBinFolder & vbCrLf
        str = str & prefix & "FrameworkAspFuncLibFile=" & Me.aspFuncLibFile & vbCrLf
        str = str & prefix & "FrameworkImageFolder=" & Me.imageFolder & vbCrLf
        Return str
    End Function


    Public Sub readParams(ByVal lines As String())
        Dim i As Integer
        Dim maxI As Integer = lines.Length - 1

        Dim line, lcase_line As String
        For i = 0 To maxI
            line = lines(i)
            lcase_line = LCase(line)
            If lcase_line.StartsWith("projectframeworkhomefile") Then
                Me.homeFile = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkcssfile") Then
                Me.cssFile = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkrootincfolder") Then
                Me.rootIncFolder = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkrootincheaderfile") Then
                Me.rootIncHeaderFile = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkrootincfooterfile") Then
                Me.rootIncFooterFile = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkadmfolder") Then
                Me.admFolder = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkadmhomefile") Then
                Me.admHomeFile = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkmemfolder") Then
                Me.memFolder = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkmemhomefile") Then
                Me.memHomeFile = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkloginfolder") Then
                Me.loginFolder = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkloginadmsecurityfile") Then
                Me.loginAdmSecurityFile = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkloginmemsecurityfile") Then
                Me.loginMemSecurityFile = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkaspbinfolder") Then
                Me.aspBinFolder = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkaspfunclibfile") Then
                Me.aspFuncLibFile = MyUtil.getStrAfterDelimiter(line, "=")
            ElseIf lcase_line.StartsWith("projectframeworkimagefolder") Then
                Me.imageFolder = MyUtil.getStrAfterDelimiter(line, "=")
            End If
        Next
        'MsgBox(Me.writeParams("Debug-"))
    End Sub

End Class
