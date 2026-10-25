Option Explicit

Public CurrentDirectory As String
Public Env As Object
Public History As Collection
Public LastExitCode As Long
Public Prompt As String
Public ShouldExit As Boolean
Public HasOutput As Boolean
Public SuppressNewline As Boolean

Private mPreviousDirectory As String

Private Sub Class_Initialize()
    CurrentDirectory = CurDir$
    Set Env = CreateObject("Scripting.Dictionary")
    Env.CompareMode = vbTextCompare
    Set History = New Collection
    LastExitCode = 0
    Prompt = "vbash> "
    ShouldExit = False

    If Len(Environ$("USERPROFILE")) > 0 Then
        SetEnv "HOME", Environ$("USERPROFILE")
    End If
End Sub

Public Sub SetEnv(ByVal key As String, ByVal value As String)
    If Not IsValidName(key) Then
        Err.Raise vbObjectError + 1010, "clsSession.SetEnv", "Nome de variável inválido: " & key
    End If
    Env(key) = value
End Sub

Public Function GetEnv(ByVal key As String) As String
    If Env.Exists(key) Then GetEnv = CStr(Env(key))
End Function

Public Sub AddHistory(ByVal commandLine As String)
    If Len(Trim$(commandLine)) > 0 Then History.Add commandLine
End Sub

Public Function ProcessInput(ByVal commandLine As String) As String
    Dim trimmed As String
    Dim tokenizer As clsTokenizer
    Dim parser As clsParser
    Dim resolver As clsCommandResolver
    Dim command As clsCommand
    Dim builtinName As String
    Dim separator As Long

    HasOutput = False
    SuppressNewline = False
    trimmed = Trim$(commandLine)
    If Len(trimmed) = 0 Then Exit Function

    AddHistory trimmed
    On Error GoTo CommandError

    Set tokenizer = New clsTokenizer
    Set parser = New clsParser
    Set command = parser.Parse(tokenizer.Tokenize(trimmed), trimmed)

    separator = InStr(1, command.Name, "=", vbBinaryCompare)
    If separator > 1 And command.Args.Count = 0 Then
        SetEnv Left$(command.Name, separator - 1), Mid$(command.Name, separator + 1)
        LastExitCode = 0
        Exit Function
    End If

    Set resolver = New clsCommandResolver
    builtinName = resolver.Resolve(command.Name)
    If Len(builtinName) = 0 Then
        ProcessInput = "Command not found: " & command.Name
        HasOutput = True
        LastExitCode = 127
        Exit Function
    End If

    LastExitCode = 0
    Select Case builtinName
        Case "echo"
            ProcessInput = BuiltinEcho(command.Args)
        Case "pwd"
            ProcessInput = BuiltinPwd(command.Args)
        Case "cd"
            ProcessInput = BuiltinCd(command.Args)
        Case "export", "setenv"
            ProcessInput = BuiltinExport(command.Args, builtinName = "setenv")
        Case "unset"
            ProcessInput = BuiltinUnset(command.Args)
        Case "help"
            ProcessInput = BuiltinHelp(command.Args)
        Case "exit"
            ProcessInput = BuiltinExit(command.Args)
    End Select
    Exit Function

CommandError:
    ProcessInput = "VBAsh: " & Err.Description
    HasOutput = True
    If Err.Number >= vbObjectError + 1000 And Err.Number < vbObjectError + 2000 Then
        LastExitCode = 2
    Else
        LastExitCode = 1
    End If
End Function

Private Function BuiltinEcho(ByVal args As Collection) As String
    Dim firstArg As Long
    Dim i As Long

    firstArg = 1
    If args.Count > 0 Then
        If CStr(args(1)) = "-n" Then
            firstArg = 2
            SuppressNewline = True
        End If
    End If

    For i = firstArg To args.Count
        If i > firstArg Then BuiltinEcho = BuiltinEcho & " "
        BuiltinEcho = BuiltinEcho & CStr(args(i))
    Next i
    HasOutput = True
End Function

Private Function BuiltinPwd(ByVal args As Collection) As String
    If args.Count > 0 Then Err.Raise vbObjectError + 1011, "clsSession.pwd", "Uso: pwd"
    BuiltinPwd = CurrentDirectory
    HasOutput = True
End Function

Private Function BuiltinCd(ByVal args As Collection) As String
    Dim fileSystem As Object
    Dim target As String
    Dim path As String
    Dim showPath As Boolean

    If args.Count > 1 Then Err.Raise vbObjectError + 1012, "clsSession.cd", "Uso: cd [diretório]"

    If args.Count = 0 Then
        path = GetEnv("HOME")
        If Len(path) = 0 Then Err.Raise vbObjectError + 1013, "clsSession.cd", "HOME não está definido."
    Else
        path = CStr(args(1))
    End If

    If path = "-" Then
        If Len(mPreviousDirectory) = 0 Then
            Err.Raise vbObjectError + 1014, "clsSession.cd", "Diretório anterior não está definido."
        End If
        target = mPreviousDirectory
        showPath = True
    Else
        If Left$(path, 1) = "~" Then
            If Len(GetEnv("HOME")) = 0 Then Err.Raise vbObjectError + 1015, "clsSession.cd", "HOME não está definido."
            If Len(path) = 1 Then
                path = GetEnv("HOME")
            ElseIf Mid$(path, 2, 1) = "\" Or Mid$(path, 2, 1) = "/" Then
                path = GetEnv("HOME") & Mid$(path, 2)
            End If
        End If

        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        If Len(path) >= 2 And Mid$(path, 2, 1) = ":" Then
            target = path
        ElseIf Left$(path, 2) = "\\" Then
            target = path
        ElseIf Left$(path, 1) = "\" Or Left$(path, 1) = "/" Then
            target = Left$(CurrentDirectory, 2) & path
        Else
            target = fileSystem.BuildPath(CurrentDirectory, path)
        End If
        target = fileSystem.GetAbsolutePathName(target)
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(target) Then
        Err.Raise vbObjectError + 1016, "clsSession.cd", "Diretório não encontrado: " & path
    End If

    mPreviousDirectory = CurrentDirectory
    CurrentDirectory = target
    Prompt = CurrentDirectory & "> "
    If showPath Then
        BuiltinCd = CurrentDirectory
        HasOutput = True
    End If
End Function

Private Function BuiltinExport(ByVal args As Collection, ByVal setenvMode As Boolean) As String
    Dim i As Long
    Dim key As Variant
    Dim assignment As String
    Dim separator As Long

    If setenvMode Then
        If args.Count <> 2 Then Err.Raise vbObjectError + 1017, "clsSession.setenv", "Uso: setenv NOME valor"
        SetEnv CStr(args(1)), CStr(args(2))
        Exit Function
    End If

    If args.Count = 0 Then
        For Each key In Env.Keys
            If Len(BuiltinExport) > 0 Then BuiltinExport = BuiltinExport & vbCrLf
            BuiltinExport = BuiltinExport & CStr(key) & "=" & CStr(Env(key))
        Next key
        HasOutput = (Len(BuiltinExport) > 0)
        Exit Function
    End If

    For i = 1 To args.Count
        assignment = CStr(args(i))
        separator = InStr(1, assignment, "=", vbBinaryCompare)
        If separator > 0 Then
            SetEnv Left$(assignment, separator - 1), Mid$(assignment, separator + 1)
        ElseIf Not Env.Exists(assignment) Then
            Err.Raise vbObjectError + 1018, "clsSession.export", "Variável não definida: " & assignment
        End If
    Next i
End Function

Private Function BuiltinUnset(ByVal args As Collection) As String
    Dim i As Long

    If args.Count = 0 Then Err.Raise vbObjectError + 1019, "clsSession.unset", "Uso: unset NOME [NOME ...]"
    For i = 1 To args.Count
        If Not IsValidName(CStr(args(i))) Then
            Err.Raise vbObjectError + 1020, "clsSession.unset", "Nome de variável inválido: " & CStr(args(i))
        End If
        If Env.Exists(CStr(args(i))) Then Env.Remove CStr(args(i))
    Next i
End Function

Private Function BuiltinHelp(ByVal args As Collection) As String
    If args.Count > 0 Then Err.Raise vbObjectError + 1021, "clsSession.help", "Uso: help"
    BuiltinHelp = "Builtins: echo [-n], pwd, cd [diretório], export, setenv, unset, help, exit [status]"
    HasOutput = True
End Function

Private Function BuiltinExit(ByVal args As Collection) As String
    Dim exitCode As Long
    Dim numericValue As Double

    If args.Count > 1 Then Err.Raise vbObjectError + 1022, "clsSession.exit", "Uso: exit [status]"
    If args.Count = 1 Then
        If Not IsNumeric(CStr(args(1))) Then Err.Raise vbObjectError + 1023, "clsSession.exit", "Status inválido."
        numericValue = CDbl(args(1))
        If numericValue < 0 Or numericValue > 255 Or numericValue <> Fix(numericValue) Then
            Err.Raise vbObjectError + 1024, "clsSession.exit", "Status deve ser um inteiro entre 0 e 255."
        End If
        exitCode = CLng(numericValue)
    End If

    LastExitCode = exitCode
    ShouldExit = True
End Function

Private Function IsValidName(ByVal name As String) As Boolean
    Dim i As Long
    Dim code As Long

    If Len(name) = 0 Then Exit Function
    For i = 1 To Len(name)
        code = AscW(Mid$(name, i, 1))
        If i = 1 Then
            If Not ((code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or code = 95) Then Exit Function
        ElseIf Not ((code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or _
                    (code >= 48 And code <= 57) Or code = 95) Then
            Exit Function
        End If
    Next i
    IsValidName = True
End Function
