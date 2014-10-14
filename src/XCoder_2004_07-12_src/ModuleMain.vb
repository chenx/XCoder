Module ModuleMain
    Dim frmMain As New FormMain()

    Public Sub main()
        Application.Run(frmMain)
    End Sub

    Public Enum xcProjectType
        xcGeneralIMSWebsiteProject
    End Enum

    Public Enum xcModuleType
        xcTemplateBase
        xcWebSiteFramework
        xcUpload
        xcDiscBoard
        xcCalendar
        xcSearch
        xcIMS
        xcLogin
    End Enum

    Public Enum xcValidateNewFileStatus
        noError
        isOpening
        alreadyExists
    End Enum

End Module
