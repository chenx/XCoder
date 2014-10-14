''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' A dialog class which displays a tree view of folders and allows 
' the user to select a folder path
'
' Similar in functionality to clsFolderBrowserSimple.vb
' a little more choices are provided.
'
' @ Not used in this application @
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Imports System.Windows.Forms.Design


Public Class clsFolderBrowser
    Inherits FolderNameEditor

    'Some folders the starting location of the FolderBrowser can be set to
    Public Enum Folder
        Desktop
        Favorites
        MyComputer
        MyDocuments
        MyPictures
        NetAndDialUpConnections
        NetworkNeighborhood
        Printers
        Recent
        SendTo
        StartMenu
        Templates
    End Enum

    'Styles effecting the look and functionality of the FolderBrowser dialog
    Public Enum BrowserStyles
        BrowseForComputer
        BrowseForEverything
        BrowseForPrinter
        RestrictToDomain
        RestrictToFileSystem
        RestrictToSubfolders
        ShowTextBox
    End Enum

    Private _browser As FolderBrowser 'the protected dialog

    Public Sub New()
        _browser = New FolderBrowser()

        ' Default values
        _browser.Style = FolderNameEditor.FolderBrowserStyles.RestrictToFilesystem
        _browser.StartLocation = FolderNameEditor.FolderBrowserFolder.Desktop
    End Sub


    Public Function ShowDialog() As DialogResult
        Return _browser.ShowDialog()
    End Function


    Public Function ShowDialog(ByVal owner As System.Windows.Forms.IWin32Window) As DialogResult
        Return _browser.ShowDialog(owner)
    End Function


    'The Path the user has chosen
    Public ReadOnly Property DirectoryPath() As String
        Get
            Return _browser.DirectoryPath
        End Get
    End Property


    'Puts a text string at the top of the dialog
    Public Property Description() As String
        Get
            Return _browser.Description
        End Get
        Set(ByVal Value As String)
            _browser.Description = Value
        End Set
    End Property


    Public Property StartLocation() As Folder
        Get
            Select Case _browser.StartLocation
                Case FolderBrowserFolder.Desktop
                    Return Folder.Desktop
                Case FolderBrowserFolder.Favorites
                    Return Folder.Favorites
                Case FolderBrowserFolder.MyComputer
                    Return Folder.MyComputer
                Case FolderBrowserFolder.MyDocuments
                    Return Folder.MyDocuments
                Case FolderBrowserFolder.MyPictures
                    Return Folder.MyPictures
                Case FolderBrowserFolder.NetAndDialUpConnections
                    Return Folder.NetAndDialUpConnections
                Case FolderBrowserFolder.NetworkNeighborhood
                    Return Folder.NetworkNeighborhood
                Case FolderBrowserFolder.Printers
                    Return Folder.Printers
                Case FolderBrowserFolder.Recent
                    Return Folder.Recent
                Case FolderBrowserFolder.SendTo
                    Return Folder.SendTo
                Case FolderBrowserFolder.StartMenu
                    Return Folder.StartMenu
                Case FolderBrowserFolder.Templates
                    Return Folder.Templates
            End Select
        End Get
        Set(ByVal Value As Folder)
            Select Case Value
                Case Folder.Desktop
                    _browser.StartLocation = FolderBrowserFolder.Desktop
                Case Folder.Favorites
                    _browser.StartLocation = FolderBrowserFolder.Favorites
                Case Folder.MyComputer
                    _browser.StartLocation = FolderBrowserFolder.MyComputer
                Case Folder.MyDocuments
                    _browser.StartLocation = FolderBrowserFolder.MyDocuments
                Case Folder.MyPictures
                    _browser.StartLocation = FolderBrowserFolder.MyPictures
                Case Folder.NetAndDialUpConnections
                    _browser.StartLocation = FolderBrowserFolder.NetAndDialUpConnections
                Case Folder.NetworkNeighborhood
                    _browser.StartLocation = FolderBrowserFolder.NetworkNeighborhood
                Case Folder.Printers
                    _browser.StartLocation = FolderBrowserFolder.Printers
                Case FolderBrowserFolder.Recent
                    _browser.StartLocation = FolderBrowserFolder.Recent
                Case Folder.SendTo
                    _browser.StartLocation = FolderBrowserFolder.SendTo
                Case Folder.StartMenu
                    _browser.StartLocation = FolderBrowserFolder.StartMenu
                Case Folder.Templates
                    _browser.StartLocation = FolderBrowserFolder.Templates
            End Select
        End Set
    End Property


    Public Property Style() As BrowserStyles
        Get
            Select Case _browser.Style
                Case FolderBrowserStyles.BrowseForComputer
                    Return BrowserStyles.BrowseForComputer
                Case FolderBrowserStyles.BrowseForEverything
                    Return BrowserStyles.BrowseForEverything
                Case FolderBrowserStyles.BrowseForPrinter
                    Return BrowserStyles.BrowseForPrinter
                Case FolderBrowserStyles.RestrictToDomain
                    Return BrowserStyles.RestrictToDomain
                Case FolderBrowserStyles.RestrictToFilesystem
                    Return BrowserStyles.RestrictToFileSystem
                Case FolderBrowserStyles.RestrictToSubfolders
                    Return BrowserStyles.RestrictToSubfolders
                Case FolderBrowserStyles.ShowTextBox
                    Return BrowserStyles.ShowTextBox
            End Select
        End Get
        Set(ByVal Value As BrowserStyles)
            Select Case Value
                Case BrowserStyles.BrowseForComputer
                    _browser.Style = FolderBrowserStyles.BrowseForComputer
                Case BrowserStyles.BrowseForEverything
                    _browser.Style = FolderBrowserStyles.BrowseForEverything
                Case BrowserStyles.BrowseForPrinter
                    _browser.Style = FolderBrowserStyles.BrowseForPrinter
                Case BrowserStyles.RestrictToDomain
                    _browser.Style = FolderBrowserStyles.RestrictToDomain
                Case BrowserStyles.RestrictToFileSystem
                    _browser.Style = FolderBrowserStyles.RestrictToFilesystem
                Case BrowserStyles.RestrictToSubfolders
                    _browser.Style = FolderBrowserStyles.RestrictToSubfolders
                Case BrowserStyles.ShowTextBox
                    _browser.Style = FolderBrowserStyles.ShowTextBox
            End Select
        End Set
    End Property

End Class
