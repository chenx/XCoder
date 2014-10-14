''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Similar in functionality to FolderBrowser.vb
' a little less choices are provided.
'
' @ This file is NOT used in this application @
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System.Windows.Forms.Design

Public Class clsFolderBrowserSimple
    Inherits FolderNameEditor

    Dim mo_FB As New FolderBrowser()

    Public Sub New()
        MyBase.new()
        'mo_FB.Style = FolderNameEditor.FolderBrowserStyles.RestrictToFilesystem
        mo_FB.Style = FolderNameEditor.FolderBrowserStyles.RestrictToFilesystem
        mo_FB.StartLocation = FolderBrowserFolder.Desktop
    End Sub

    Public Sub ShowDialog()
        mo_FB.ShowDialog()
    End Sub

    Public ReadOnly Property DirectoryPath() As String
        Get
            Return mo_FB.DirectoryPath
        End Get
    End Property

    Public WriteOnly Property Description() As String
        Set(ByVal Value As String)
            mo_FB.Description = Value
        End Set
    End Property

End Class