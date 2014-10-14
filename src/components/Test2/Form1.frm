VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4932
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   7896
   LinkTopic       =   "Form1"
   ScaleHeight     =   4932
   ScaleWidth      =   7896
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListFields 
      Height          =   2160
      Left            =   5400
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ListBox ListTables 
      Height          =   2160
      Left            =   3120
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ListBox ListDatabases 
      Height          =   2160
      Left            =   720
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "SallySalt"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "sa"
      Top             =   720
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lbMsg 
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   4320
      TabIndex        =   13
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label S 
      Caption         =   "SQL Servers"
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Fields"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Tables"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label label3 
      Caption         =   "Databases"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Login:"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.metaprosystems.com/TipsTricks30.htm
' http://www.codeproject.com/useritems/listsqlservers.asp?df=100&forumid=3432&exp=0&select=509443
'http://www.codeproject.com/cs/database/LocatingSql.asp
'http://livedocs.macromedia.com/coldfusion/5.0/Installing_and_Configuring_ColdFusion_Server/datasources7.htm
'http://delphi.about.com/od/sqlservermsdeaccess/l/aa090704a.htm
'http://www.codeproject.com/cs/database/CSharp_Wrapper.asp
'http://www.informit.com/articles/article.asp?p=27199&seqNum=3

' current connection information
    Dim data_src As String
    Dim uid As String
    Dim db As String
    Dim pass As String


Private Sub Command1_Click()
    Dim con As New ADODB.Connection
    
    On Error GoTo Error
    Err.Clear
  
    data_src = Me.Combo1.Text
    If data_src = "" Then Exit Sub
    
    uid = Me.Text1.Text
    pass = Me.Text2.Text
    
    Dim connStr As String
    ' http://www.able-consulting.com/MDAC/ADO/Connection/OLEDB_Providers.htm#OLEDBProviderForSQLServer
    connStr = "Provider=SQLOLEDB;Data Source=" & data_src & ";Initial Catalog=;User Id=" & uid & ";Password=" & pass
    'MsgBox connStr
    
    con.Open connStr
    'con.Execute "USE master"
    
    'Dim rs As ADODB.Recordset
    'con.Execute sp_databases
    'Set rs = con.OpenSchema(adSchemaTables)
    getAllDatabases data_src, uid, pass
       
    setMsg data_src & " is connected."
    con.Close
    
exit_sub:
    Exit Sub
    
Error:
    If Err.Number <> 0 Then
        setMsg "Cannot connect to " & data_src ' & Err.Description
        Call clearListboxes
        GoTo exit_sub
    End If
End Sub


Private Sub setMsg(ByVal msg As String)
    Me.lbMsg.Caption = msg
End Sub


Private Function getAllDatabases(ByVal db As String, ByVal login As String, ByVal pass As String)
    Dim sv As New GetSQLServers
    Dim c As Collection
    Set c = sv.NonSystemTables(db, login, pass)
    
    fillListBox Me.ListDatabases, c
End Function


Private Sub fillListBox(ByRef lb As ListBox, ByRef c As Collection)
    If IsNull(c) Or c.Count = 0 Then Exit Sub
    
    lb.Clear
    Dim i As Integer
    For i = 1 To c.Count
        lb.AddItem (c(i))
    Next
End Sub


Private Sub clearListBox(ByRef lb As ListBox)
    lb.Clear
End Sub


Private Sub Form_Load()
    Dim sv As New GetSQLServers
        
    Dim c As Collection
    Set c = sv.GetLocalServerNames
    
    Dim i As Integer
    For i = 1 To c.Count
        Me.Combo1.AddItem (c(i))
    Next
    
    If c.Count >= 1 Then
        'Me.Combo1.ListIndex = 0
    End If
End Sub


Private Sub ListDatabases_Click()
    Dim dbName As String
    dbName = Me.ListDatabases.Text
    'MsgBox dbName
    
    On Error GoTo err_handler
    Dim con As New ADODB.Connection

    Dim connStr As String
    ' http://www.able-consulting.com/MDAC/ADO/Connection/OLEDB_Providers.htm#OLEDBProviderForSQLServer
    connStr = "Provider=SQLOLEDB;Data Source=" & data_src & ";Initial Catalog=" & dbName & ";User Id=" & uid & ";Password=" & pass
    'MsgBox connStr
    
    con.Open connStr
    
    'Dim sql As String
    Dim rs As ADODB.Recordset
    
    ' http://www.psacake.com/web/hk.asp
    con.Execute ("Use " & dbName)
    'Set rs = con.Execute("EXECUTE sp_tables null, null, '" & dbName & "', null")
    Set rs = con.Execute("EXECUTE sp_tables")
    
    Dim c As New Collection
    Do While Not rs.EOF And Not rs.BOF
        If CStr(rs("table_type")) = "TABLE" And CStr(rs("Table_name")) <> "dtproperties" Then
          c.Add (rs("table_name"))
        End If
        rs.MoveNext
    Loop
    
    Call fillListBox(Me.ListTables, c)

    con.Close
exit_sub:
    Exit Sub
    
err_handler:
    If Err.Number <> 0 Then
        setMsg "Cannot list tables in " & dbName & ": " & Err.Description
        Call clearListboxes
        GoTo exit_sub
    End If
End Sub


Private Sub clearListboxes()
    clearListBox ListDatabases
    clearListBox ListTables
    clearListBox ListFields
End Sub



Private Sub ListTables_Click()
    Dim dbName As String
    dbName = Me.ListDatabases.Text
    
    Dim tblName As String
    tblName = Me.ListTables.Text
    
    On Error GoTo err_handler
    Dim con As New ADODB.Connection

    Dim connStr As String
    connStr = "Provider=SQLOLEDB;Data Source=" & data_src & ";Initial Catalog=" & dbName & ";User Id=" & uid & ";Password=" & pass
    
    con.Open connStr
    
    Dim rs As ADODB.Recordset
    
    con.Execute ("Use " & dbName)
    Set rs = con.Execute("EXECUTE sp_columns " & tblName)
    
    Dim c As New Collection
    Do While Not rs.EOF And Not rs.BOF
        c.Add (rs("column_name"))
        rs.MoveNext
    Loop
    
    Call fillListBox(Me.ListFields, c)

    con.Close
exit_sub:
    Exit Sub
    
err_handler:
    If Err.Number <> 0 Then
        setMsg "Cannot list tables in " & dbName & ": " & Err.Description
        Call clearListboxes
        GoTo exit_sub
    End If
End Sub
