Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim fileSystem As Object
    Dim target As String
    Dim path As String
    Dim showPath As Boolean

    If args.Count > 1 Then Err.Raise vbObjectError + 1012, "clsBuiltinCd", "Uso: cd [diretório]"

    If args.Count = 0 Then
        path = session.GetEnv("HOME")
        If Len(path) = 0 Then Err.Raise vbObjectError + 1013, "clsBuiltinCd", "HOME não está definido."
    Else
        path = CStr(args(1))
    End If

    If path = "-" Then
        target = session.PreviousDirectory
        If Len(target) = 0 Then Err.Raise vbObjectError + 1014, "clsBuiltinCd", "Diretório anterior não está definido."
        showPath = True
    Else
        If Left$(path, 1) = "~" Then
            If Len(session.GetEnv("HOME")) = 0 Then Err.Raise vbObjectError + 1015, "clsBuiltinCd", "HOME não está definido."
            If Len(path) = 1 Then
                path = session.GetEnv("HOME")
            ElseIf Mid$(path, 2, 1) = "\" Or Mid$(path, 2, 1) = "/" Then
                path = session.GetEnv("HOME") & Mid$(path, 2)
            End If
        End If

        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        If Len(path) >= 2 And Mid$(path, 2, 1) = ":" Then
            target = path
        ElseIf Left$(path, 2) = "\\" Then
            target = path
        ElseIf Left$(path, 1) = "\" Or Left$(path, 1) = "/" Then
            target = Left$(session.CurrentDirectory, 2) & path
        Else
            target = fileSystem.BuildPath(session.CurrentDirectory, path)
        End If
        target = fileSystem.GetAbsolutePathName(target)
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(target) Then
        Err.Raise vbObjectError + 1016, "clsBuiltinCd", "Diretório não encontrado: " & path
    End If

    session.ChangeDirectory target
    If showPath Then
        IBuiltin_Execute = session.CurrentDirectory
        session.HasOutput = True
    End If
End Function
