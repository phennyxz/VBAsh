Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim fileSystem As Object
    Dim path As String

    If args.Count <> 1 Then Err.Raise vbObjectError + 1052, "clsBuiltinMkdir", "Uso: mkdir caminho"

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    path = ResolvePath(session, CStr(args(1)), fileSystem)
    If fileSystem.FolderExists(path) Or fileSystem.FileExists(path) Then
        Err.Raise vbObjectError + 1053, "clsBuiltinMkdir", "O caminho já existe: " & path
    End If

    fileSystem.CreateFolder path
    session.HasOutput = False
End Function

Private Function ResolvePath(ByVal session As clsSession, ByVal path As String, _
                             ByVal fileSystem As Object) As String
    If Len(path) >= 2 And Mid$(path, 2, 1) = ":" Then
        ResolvePath = fileSystem.GetAbsolutePathName(path)
    ElseIf Left$(path, 2) = "\\" Then
        ResolvePath = fileSystem.GetAbsolutePathName(path)
    ElseIf Left$(path, 1) = "\" Or Left$(path, 1) = "/" Then
        ResolvePath = fileSystem.GetAbsolutePathName(Left$(session.CurrentDirectory, 2) & path)
    Else
        ResolvePath = fileSystem.GetAbsolutePathName( _
            fileSystem.BuildPath(session.CurrentDirectory, path))
    End If
End Function
