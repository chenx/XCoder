Public Class SoundManager
    ' To play sound
    Declare Auto Function sndPlaySound Lib "WINMM.DLL" (ByVal FileName As String, ByVal Options As Int32) As Int32
    Public Shared Sub myPlaySound(ByVal sFileName As String)
        If sFileName <> "" And IO.File.Exists(sFileName) = True And _
            SetupManager.USE_SOUND = True Then
            sndPlaySound(sFileName, &H0)
        End If
    End Sub

    Public Shared Sub sdDing()
        myPlaySound(SetupManager.SOUND_PATH & "ding.wav")
    End Sub

    Public Shared Sub sdComplete()
        myPlaySound(SetupManager.SOUND_PATH & "complete.wav")
    End Sub

    Public Shared Sub sdConnect()
        myPlaySound(SetupManager.SOUND_PATH & "connect.wav")
    End Sub

    Public Shared Sub sdError()
        myPlaySound(SetupManager.SOUND_PATH & "error.wav")
    End Sub

    Public Shared Sub sdType()
        myPlaySound(SetupManager.SOUND_PATH & "type.wav")
    End Sub


End Class
