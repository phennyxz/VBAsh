Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim fileSystem As Object
    Dim folder As Object
    Dim entry As Object
    Dim path As String

    If args.Count > 1 Then Err.Raise vbObjectError + 1050, "clsBuiltinLs", "Uso: ls [caminho]"

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If args.Count = 0 Then
        path = session.CurrentDirectory
    Else
        path = ResolvePath(session, CStr(args(1)), fileSystem)
    End If

    If fileSystem.FileExists(path) Then
        IBuiltin_Execute = fileSystem.GetFileName(path)
    ElseIf fileSystem.FolderExists(path) Then
        Set folder = fileSystem.GetFolder(path)
        For Each entry In folder.SubFolders
            AppendEntry IBuiltin_Execute, entry.Name
        Next entry
        For Each entry In folder.Files
            AppendEntry IBuiltin_Execute, entry.Name
        Next entry
    Else
        Err.Raise vbObjectError + 1051, "clsBuiltinLs", "Caminho não encontrado: " & path
    End If

    session.HasOutput = True
End Function

Private Sub AppendEntry(ByRef output As String, ByVal name As String)
    If Len(output) > 0 Then output = output & vbCrLf
    output = output & name
End Sub

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
