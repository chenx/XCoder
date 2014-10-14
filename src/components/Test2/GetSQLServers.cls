VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetSQLServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetLocalServerNames() As Collection
    Dim ServerNames As New Collection
   
    Dim SQLServer As SQLDMO.Application
    Dim ServerList As SQLDMO.NameList
    Dim iItem As Integer
    
    Set SQLServer = New SQLDMO.Application
    Set ServerList = SQLServer.ListAvailableSQLServers
    
    For iItem = 1 To ServerList.Count
        ServerNames.Add (ServerList.Item(iItem))
    Next
    
    Set ServerList = Nothing
    Set SQLServer = Nothing
    
    Set GetLocalServerNames = ServerNames

End Function


Public Function NonSystemTables(srvName As String, srvUserName _
As String, srvUserPass As String) As Collection
    ' http://www.freevbcode.com/ShowCode.asp?ID=989
    'Input: srvName - SQL server Name
    '       srvUserName - User Name for connect to server
    '       srvUserPass - User Password
    
    'Returns: Collection of the names
    'of non-system tables in that database
    'or Nothing if there is an error
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    Dim oSQLServer As New SQLServer
    Dim colTables As Collection
    
    Set colTables = New Collection
    oSQLServer.Connect srvName, srvUserName, srvUserPass
        For i = 1 To oSQLServer.Databases.Count
            With oSQLServer.Databases(i)
                If Not .SystemObject Then
                    colTables.Add .Name
                End If
            End With
        Next i
    
    Set NonSystemTables = colTables
    
    Exit Function
    
ErrHandler:
    On Error Resume Next
    
    Set NonSystemTables = Nothing

End Function


 




