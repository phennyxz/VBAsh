Option Explicit

Private mTextStream As Object

Public Sub OpenForWrite(ByVal path As String, Optional ByVal append As Boolean = False)
    Dim fileSystem As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If append Then
        Set mTextStream = fileSystem.OpenTextFile(path, 8, True)
    Else
        Set mTextStream = fileSystem.CreateTextFile(path, True)
    End If
End Sub

Public Function WriteText(ByVal text As String) As Long
    If mTextStream Is Nothing Then Err.Raise vbObjectError + 1030, "clsFileStream", "Stream não está aberto."
    mTextStream.Write text
    WriteText = Len(text)
End Function

Public Sub CloseStream()
    If Not mTextStream Is Nothing Then
        mTextStream.Close
        Set mTextStream = Nothing
    End If
End Sub

Private Sub Class_Terminate()
    CloseStream
End Sub
