Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim fileSystem As Object
    Dim testOperator As String
    Dim value As String
    Dim path As String
    Dim isTrue As Boolean

    Select Case args.Count
        Case 0
            isTrue = False
        Case 1
            isTrue = (Len(CStr(args(1))) > 0)
        Case 2
            testOperator = CStr(args(1))
            value = CStr(args(2))
            Select Case testOperator
                Case "-n"
                    isTrue = (Len(value) > 0)
                Case "-z"
                    isTrue = (Len(value) = 0)
                Case "-d", "-f", "-e"
                    Set fileSystem = CreateObject("Scripting.FileSystemObject")
                    path = ResolveTestPath(session, value, fileSystem)
                    Select Case testOperator
                        Case "-d"
                            isTrue = fileSystem.FolderExists(path)
                        Case "-f"
                            isTrue = fileSystem.FileExists(path)
                        Case "-e"
                            isTrue = fileSystem.FolderExists(path) Or fileSystem.FileExists(path)
                    End Select
                Case Else
                    Err.Raise vbObjectError + 1040, "clsBuiltinTest", _
                        "Operador de teste não suportado: " & testOperator
            End Select
        Case 3
            testOperator = CStr(args(2))
            value = CStr(args(1))
            Select Case testOperator
                Case "=", "=="
                    isTrue = (StrComp(value, CStr(args(3)), vbBinaryCompare) = 0)
                Case "!="
                    isTrue = (StrComp(value, CStr(args(3)), vbBinaryCompare) <> 0)
                Case Else
                    Err.Raise vbObjectError + 1041, "clsBuiltinTest", _
                        "Operador de teste não suportado: " & testOperator
            End Select
        Case Else
            Err.Raise vbObjectError + 1042, "clsBuiltinTest", "Uso: test [expressão]"
    End Select

    If isTrue Then
        session.LastExitCode = 0
    Else
        session.LastExitCode = 1
    End If
    session.HasOutput = False
    IBuiltin_Execute = ""
End Function

Private Function ResolveTestPath(ByVal session As clsSession, ByVal path As String, _
                                 ByVal fileSystem As Object) As String
    If Len(path) >= 2 And Mid$(path, 2, 1) = ":" Then
        ResolveTestPath = fileSystem.GetAbsolutePathName(path)
    ElseIf Left$(path, 2) = "\\" Then
        ResolveTestPath = fileSystem.GetAbsolutePathName(path)
    ElseIf Left$(path, 1) = "\" Or Left$(path, 1) = "/" Then
        ResolveTestPath = fileSystem.GetAbsolutePathName(Left$(session.CurrentDirectory, 2) & path)
    Else
        ResolveTestPath = fileSystem.GetAbsolutePathName( _
            fileSystem.BuildPath(session.CurrentDirectory, path))
    End If
End Function
