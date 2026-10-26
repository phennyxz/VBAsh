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
Private mBuiltinRegistry As clsCommandResolver

Private Sub Class_Initialize()
    CurrentDirectory = CurDir$
    Set Env = CreateObject("Scripting.Dictionary")
    Env.CompareMode = vbTextCompare
    Set History = New Collection
    LastExitCode = 0
    Prompt = "vbash> "
    ShouldExit = False
    Set mBuiltinRegistry = New clsCommandResolver

    If Len(Environ$("USERPROFILE")) > 0 Then
        SetEnv "HOME", Environ$("USERPROFILE")
    End If
End Sub

Public Sub SetEnv(ByVal key As String, ByVal value As String)
    If Not IsValidEnvironmentName(key) Then
        Err.Raise vbObjectError + 1010, "clsSession.SetEnv", "Nome de variável inválido: " & key
    End If
    Env(key) = value
End Sub

Public Function GetEnv(ByVal key As String) As String
    If Env.Exists(key) Then GetEnv = CStr(Env(key))
End Function

Public Sub RemoveEnv(ByVal key As String)
    If Env.Exists(key) Then Env.Remove key
End Sub

Public Function PreviousDirectory() As String
    PreviousDirectory = mPreviousDirectory
End Function

Public Sub ChangeDirectory(ByVal target As String)
    mPreviousDirectory = CurrentDirectory
    CurrentDirectory = target
    Prompt = CurrentDirectory & "> "
End Sub

Public Sub AddHistory(ByVal commandLine As String)
    If Len(Trim$(commandLine)) > 0 Then History.Add commandLine
End Sub

Public Function ProcessInput(ByVal commandLine As String) As String
    Dim trimmed As String
    Dim tokenizer As clsTokenizer
    Dim tokens As Collection
    Dim parser As clsParser
    Dim command As clsCommand
    Dim builtin As IBuiltin
    Dim separator As Long

    HasOutput = False
    SuppressNewline = False
    trimmed = Trim$(commandLine)
    If Len(trimmed) = 0 Then Exit Function

    AddHistory trimmed
    On Error GoTo CommandError

    Set tokenizer = New clsTokenizer
    Set tokens = tokenizer.Tokenize(trimmed)
    If tokens.Count = 1 Then
        LastExitCode = 0
        Exit Function
    End If
    Set parser = New clsParser
    Set command = parser.Parse(tokens, trimmed)

    separator = InStr(1, command.Name, "=", vbBinaryCompare)
    If separator > 1 And command.Args.Count = 0 Then
        SetEnv Left$(command.Name, separator - 1), Mid$(command.Name, separator + 1)
        LastExitCode = 0
        Exit Function
    End If

    Set builtin = mBuiltinRegistry.Resolve(command.Name)
    If builtin Is Nothing Then
        ProcessInput = "Command not found: " & command.Name
        HasOutput = True
        LastExitCode = 127
        Exit Function
    End If

    LastExitCode = 0
    ProcessInput = builtin.Execute(Me, command.Args)
    If command.Redirections.Count > 0 Then
        ApplyOutputRedirection command, ProcessInput
        ProcessInput = ""
        HasOutput = False
    End If
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

Public Function IsValidEnvironmentName(ByVal name As String) As Boolean
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
    IsValidEnvironmentName = True
End Function

Private Sub ApplyOutputRedirection(ByVal command As clsCommand, ByVal output As String)
    Dim redirection As clsRedirection
    Dim stream As clsFileStream
    Dim text As String

    If command.Redirections.Count > 1 Then
        Err.Raise vbObjectError + 1025, "clsSession", "Apenas um redirecionamento é suportado."
    End If

    Set redirection = command.Redirections(1)
    If redirection.Operator <> "REDIR_OUT" And redirection.Operator <> "REDIR_APPEND" Then
        Err.Raise vbObjectError + 1026, "clsSession", "Apenas redirecionamento de saída é suportado."
    End If
    If Len(redirection.Target) = 0 Then
        Err.Raise vbObjectError + 1027, "clsSession", "Destino de redirecionamento vazio."
    End If

    text = output
    If Not SuppressNewline Then text = text & vbCrLf
    Set stream = New clsFileStream
    stream.OpenForWrite ResolvePath(redirection.Target), redirection.Operator = "REDIR_APPEND"
    stream.WriteText text
    stream.CloseStream
End Sub

Private Function ResolvePath(ByVal path As String) As String
    Dim fileSystem As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Len(path) >= 2 And Mid$(path, 2, 1) = ":" Then
        ResolvePath = fileSystem.GetAbsolutePathName(path)
    ElseIf Left$(path, 2) = "\\" Then
        ResolvePath = fileSystem.GetAbsolutePathName(path)
    ElseIf Left$(path, 1) = "\" Or Left$(path, 1) = "/" Then
        ResolvePath = fileSystem.GetAbsolutePathName(Left$(CurrentDirectory, 2) & path)
    Else
        ResolvePath = fileSystem.GetAbsolutePathName(fileSystem.BuildPath(CurrentDirectory, path))
    End If
End Function
