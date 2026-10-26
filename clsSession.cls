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

    If LCase$(FirstWord(tokens)) = "if" Then
        ProcessInput = ExecuteIf(tokens, trimmed)
    Else
        ProcessInput = ExecuteTokens(tokens, trimmed)
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

Private Function ExecuteTokens(ByVal tokens As Collection, ByVal originalText As String) As String
    Dim parser As clsParser
    Dim command As clsCommand
    Dim builtin As IBuiltin
    Dim separator As Long

    HasOutput = False
    SuppressNewline = False
    Set parser = New clsParser
    Set command = parser.Parse(tokens, originalText)

    separator = InStr(1, command.Name, "=", vbBinaryCompare)
    If separator > 1 And command.Args.Count = 0 Then
        SetEnv Left$(command.Name, separator - 1), Mid$(command.Name, separator + 1)
        LastExitCode = 0
        Exit Function
    End If

    Set builtin = mBuiltinRegistry.Resolve(command.Name)
    If builtin Is Nothing Then
        ExecuteTokens = "Command not found: " & command.Name
        HasOutput = True
        LastExitCode = 127
        Exit Function
    End If

    LastExitCode = 0
    ExecuteTokens = builtin.Execute(Me, command.Args)
    If command.Redirections.Count > 0 Then
        ApplyOutputRedirection command, ExecuteTokens
        ExecuteTokens = ""
        HasOutput = False
    End If
End Function

Private Function ExecuteIf(ByVal tokens As Collection, ByVal originalText As String) As String
    Dim conditionStart As Long
    Dim conditionEnd As Long
    Dim thenStart As Long
    Dim thenEnd As Long
    Dim elseStart As Long
    Dim elseEnd As Long
    Dim firstSeparator As Long
    Dim branchSeparator As Long
    Dim elseSeparator As Long
    Dim markerIndex As Long

    conditionStart = 2
    firstSeparator = FindToken(tokens, "SEMI", conditionStart)
    If firstSeparator = 0 Then
        Err.Raise vbObjectError + 1031, "clsSession.if", _
            "Uso: if test condição; then comando; [else comando;] fi"
    End If

    conditionEnd = firstSeparator - 1
    If conditionEnd < conditionStart Then
        Err.Raise vbObjectError + 1031, "clsSession.if", "Condição ausente no if."
    End If

    markerIndex = firstSeparator + 1
    If Not IsWordAt(tokens, markerIndex, "then") Then
        Err.Raise vbObjectError + 1031, "clsSession.if", "Esperado 'then' após a condição."
    End If
    thenStart = markerIndex + 1
    branchSeparator = FindToken(tokens, "SEMI", thenStart)
    If branchSeparator = 0 Or branchSeparator = thenStart Then
        Err.Raise vbObjectError + 1031, "clsSession.if", "Comando ausente após 'then'."
    End If
    thenEnd = branchSeparator - 1

    markerIndex = branchSeparator + 1
    If IsWordAt(tokens, markerIndex, "else") Then
        elseStart = markerIndex + 1
        elseSeparator = FindToken(tokens, "SEMI", elseStart)
        If elseSeparator = 0 Or elseSeparator = elseStart Then
            Err.Raise vbObjectError + 1031, "clsSession.if", "Comando ausente após 'else'."
        End If
        elseEnd = elseSeparator - 1
        markerIndex = elseSeparator + 1
        If Not IsWordAt(tokens, markerIndex, "fi") Then
            Err.Raise vbObjectError + 1031, "clsSession.if", "Esperado 'fi' ao final do if."
        End If
    ElseIf IsWordAt(tokens, markerIndex, "fi") Then
        elseStart = 0
    Else
        Err.Raise vbObjectError + 1031, "clsSession.if", "Esperado 'else' ou 'fi' após o comando."
    End If

    If markerIndex <> tokens.Count - 1 Then
        Err.Raise vbObjectError + 1031, "clsSession.if", "Texto inesperado após 'fi'."
    End If

    LastExitCode = 0
    ExecuteTokens = ExecuteTokens(TokenRange(tokens, conditionStart, conditionEnd), originalText)
    HasOutput = False
    SuppressNewline = False

    If LastExitCode = 0 Then
        ExecuteIf = ExecuteTokens(TokenRange(tokens, thenStart, thenEnd), originalText)
    ElseIf elseStart > 0 Then
        ExecuteIf = ExecuteTokens(TokenRange(tokens, elseStart, elseEnd), originalText)
    End If
End Function

Private Function FindToken(ByVal tokens As Collection, ByVal tokenType As String, ByVal startIndex As Long) As Long
    Dim i As Long
    Dim token As clsToken

    For i = startIndex To tokens.Count
        Set token = tokens(i)
        If token.TokenType = "EOF" Then Exit Function
        If token.TokenType = tokenType Then
            FindToken = i
            Exit Function
        End If
    Next i
End Function

Private Function IsWordAt(ByVal tokens As Collection, ByVal index As Long, ByVal word As String) As Boolean
    Dim token As clsToken

    If index < 1 Or index >= tokens.Count Then Exit Function
    Set token = tokens(index)
    IsWordAt = (token.TokenType = "WORD" And StrComp(token.Value, word, vbTextCompare) = 0)
End Function

Private Function TokenRange(ByVal tokens As Collection, ByVal firstIndex As Long, ByVal lastIndex As Long) As Collection
    Dim result As New Collection
    Dim token As clsToken
    Dim i As Long

    For i = firstIndex To lastIndex
        Set token = tokens(i)
        result.Add token
    Next i
    Set token = New clsToken
    token.Init "EOF", ""
    result.Add token
    Set TokenRange = result
End Function

Private Function FirstWord(ByVal tokens As Collection) As String
    Dim token As clsToken

    If tokens.Count = 0 Then Exit Function
    Set token = tokens(1)
    If token.TokenType = "WORD" Then FirstWord = token.Value
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
