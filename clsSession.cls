Option Explicit

Public CurrentDirectory As String
Public Env As Object
Public History As Collection
Public LastExitCode As Long
Public Prompt As String
Public ShouldExit As Boolean
Public HasOutput As Boolean
Public ErrorOutput As String
Public HasErrorOutput As Boolean
Public SuppressNewline As Boolean
Public ClearRequested As Boolean

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
    ErrorOutput = ""
    HasErrorOutput = False
    SuppressNewline = False
    ClearRequested = False
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
    HasOutput = False
    ErrorOutput = "VBAsh: " & Err.Description
    HasErrorOutput = True
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
    Dim errorNumber As Long
    Dim errorDescription As String

    HasOutput = False
    HasErrorOutput = False
    ErrorOutput = ""
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
        ErrorOutput = "Command not found: " & command.Name
        HasErrorOutput = True
        LastExitCode = 127
    Else
        LastExitCode = 0
        On Error GoTo BuiltinError
        ExecuteTokens = builtin.Execute(Me, command.Args)
    End If

    If command.Redirections.Count > 0 Then
        ApplyRedirections command, ExecuteTokens, ErrorOutput
    End If
    Exit Function

BuiltinError:
    errorNumber = Err.Number
    errorDescription = Err.Description
    HasOutput = False
    ErrorOutput = "VBAsh: " & errorDescription
    HasErrorOutput = True
    If errorNumber >= vbObjectError + 1000 And errorNumber < vbObjectError + 2000 Then
        LastExitCode = 2
    Else
        LastExitCode = 1
    End If
    Resume Next
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
    Dim conditionStatus As Long
    Dim conditionOutput As String
    Dim conditionHasOutput As Boolean
    Dim conditionSuppressNewline As Boolean
    Dim branchOutput As String
    Dim branchHasOutput As Boolean
    Dim branchSuppressNewline As Boolean
    Dim conditionErrors As String
    Dim conditionHasErrors As Boolean
    Dim branchErrors As String
    Dim branchHasErrors As Boolean

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

    conditionStatus = EvaluateCondition(tokens, conditionStart, conditionEnd, originalText, _
                                        conditionOutput, conditionHasOutput, conditionSuppressNewline, _
                                        conditionErrors, conditionHasErrors)
    LastExitCode = conditionStatus
    HasOutput = False
    SuppressNewline = False

    If conditionStatus = 0 Then
        branchOutput = ExecuteTokens(TokenRange(tokens, thenStart, thenEnd), originalText)
        branchHasOutput = HasOutput
        branchSuppressNewline = SuppressNewline
        branchErrors = ErrorOutput
        branchHasErrors = HasErrorOutput
    ElseIf elseStart > 0 Then
        branchOutput = ExecuteTokens(TokenRange(tokens, elseStart, elseEnd), originalText)
        branchHasOutput = HasOutput
        branchSuppressNewline = SuppressNewline
        branchErrors = ErrorOutput
        branchHasErrors = HasErrorOutput
    Else
        LastExitCode = 0
    End If

    If conditionHasOutput And branchHasOutput And Not conditionSuppressNewline Then
        conditionOutput = conditionOutput & vbCrLf
    End If
    ExecuteIf = conditionOutput & branchOutput
    HasOutput = conditionHasOutput Or branchHasOutput
    If branchHasOutput Then
        SuppressNewline = branchSuppressNewline
    Else
        SuppressNewline = conditionSuppressNewline
    End If
    ErrorOutput = conditionErrors
    If Len(conditionErrors) > 0 And Len(branchErrors) > 0 Then ErrorOutput = ErrorOutput & vbCrLf
    ErrorOutput = ErrorOutput & branchErrors
    HasErrorOutput = conditionHasErrors Or branchHasErrors
End Function

Private Function EvaluateCondition(ByVal tokens As Collection, ByVal firstIndex As Long, _
                                   ByVal lastIndex As Long, ByVal originalText As String, _
                                   ByRef conditionOutput As String, ByRef conditionHasOutput As Boolean, _
                                   ByRef conditionSuppressNewline As Boolean, _
                                   ByRef conditionErrors As String, ByRef conditionHasErrors As Boolean) As Long
    Dim termStart As Long
    Dim termEnd As Long
    Dim operatorIndex As Long
    Dim pendingOperator As String
    Dim currentStatus As Long
    Dim shouldExecute As Boolean
    Dim termTokens As Collection
    Dim token As clsToken
    Dim termOutput As String
    Dim termErrors As String

    termStart = firstIndex
    Do
        operatorIndex = FindConditionOperator(tokens, termStart, lastIndex)
        If operatorIndex = 0 Then
            termEnd = lastIndex
        Else
            termEnd = operatorIndex - 1
        End If
        If termEnd < termStart Then
            Err.Raise vbObjectError + 1032, "clsSession.if", "Condição incompleta perto de '" & pendingOperator & "'."
        End If

        Set termTokens = ConditionTerm(tokens, termStart, termEnd)
        shouldExecute = (termStart = firstIndex)
        If pendingOperator = "AND_IF" And currentStatus = 0 Then shouldExecute = True
        If pendingOperator = "OR_IF" And currentStatus <> 0 Then shouldExecute = True

        If shouldExecute Then
            termOutput = ExecuteConditionTerm(termTokens, originalText)
            currentStatus = LastExitCode
            termErrors = ErrorOutput
            If HasErrorOutput Then
                If conditionHasErrors Then conditionErrors = conditionErrors & vbCrLf
                conditionErrors = conditionErrors & termErrors
                conditionHasErrors = True
            End If
            If HasOutput Then
                If conditionHasOutput And Not conditionSuppressNewline Then
                    conditionOutput = conditionOutput & vbCrLf
                End If
                conditionOutput = conditionOutput & termOutput
                conditionHasOutput = True
                conditionSuppressNewline = SuppressNewline
            End If
        End If
        HasOutput = False
        ErrorOutput = ""
        HasErrorOutput = False
        SuppressNewline = False

        If operatorIndex = 0 Then Exit Do
        Set token = tokens(operatorIndex)
        pendingOperator = token.TokenType
        termStart = operatorIndex + 1
        If termStart > lastIndex Then
            Err.Raise vbObjectError + 1032, "clsSession.if", "Condição incompleta após '" & token.Value & "'."
        End If
    Loop

    EvaluateCondition = currentStatus
End Function

Private Function ExecuteConditionTerm(ByVal tokens As Collection, ByVal originalText As String) As String
    On Error GoTo CommandError
    ExecuteConditionTerm = ExecuteTokens(tokens, originalText)
    Exit Function

CommandError:
    ErrorOutput = "VBAsh: " & Err.Description
    HasErrorOutput = True
    SuppressNewline = False
    If Err.Number >= vbObjectError + 1000 And Err.Number < vbObjectError + 2000 Then
        LastExitCode = 2
    Else
        LastExitCode = 1
    End If
End Function

Private Function FindConditionOperator(ByVal tokens As Collection, ByVal firstIndex As Long, _
                                       ByVal lastIndex As Long) As Long
    Dim i As Long
    Dim token As clsToken

    For i = firstIndex To lastIndex
        Set token = tokens(i)
        If token.TokenType = "AND_IF" Or token.TokenType = "OR_IF" Then
            FindConditionOperator = i
            Exit Function
        End If
    Next i
End Function

Private Function ConditionTerm(ByVal tokens As Collection, ByVal firstIndex As Long, _
                               ByVal lastIndex As Long) As Collection
    Dim firstToken As clsToken
    Dim lastToken As clsToken
    Dim result As New Collection
    Dim i As Long

    Set firstToken = tokens(firstIndex)
    If firstToken.TokenType <> "WORD" Then
        Err.Raise vbObjectError + 1033, "clsSession.if", "Esperado 'test' ou '[ expressão ]' na condição."
    End If

    If firstToken.Value = "[" Then
        Set lastToken = tokens(lastIndex)
        If lastToken.TokenType <> "WORD" Or lastToken.Value <> "]" Then
            Err.Raise vbObjectError + 1034, "clsSession.if", "Esperado ']' ao final da condição."
        End If
        Set firstToken = New clsToken
        firstToken.Init "WORD", "test"
        result.Add firstToken
        For i = firstIndex + 1 To lastIndex - 1
            Set firstToken = tokens(i)
            result.Add firstToken
        Next i
    ElseIf StrComp(firstToken.Value, "test", vbTextCompare) = 0 Then
        For i = firstIndex To lastIndex
            Set firstToken = tokens(i)
            result.Add firstToken
        Next i
    Else
        For i = firstIndex To lastIndex
            Set firstToken = tokens(i)
            result.Add firstToken
        Next i
    End If

    Set firstToken = New clsToken
    firstToken.Init "EOF", ""
    result.Add firstToken
    Set ConditionTerm = result
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

Private Sub ApplyRedirections(ByVal command As clsCommand, ByVal output As String, _
                              ByVal errorText As String)
    Dim redirection As clsRedirection
    Dim stdoutType As String
    Dim stdoutPath As String
    Dim stdoutAppend As Boolean
    Dim stderrType As String
    Dim stderrPath As String
    Dim stderrAppend As Boolean
    Dim redirectPath As String
    Dim mergedText As String
    Dim mergedHasOutput As Boolean
    Dim mergedAppend As Boolean
    Dim mergedSuppressNewline As Boolean

    stdoutType = "CONSOLE"
    stderrType = "CONSOLE"
    For Each redirection In command.Redirections
        Select Case redirection.Operator
            Case "REDIR_OUT", "REDIR_APPEND", "REDIR_ERR", "REDIR_ERR_APPEND"
                redirectPath = redirection.Target
                If Len(redirectPath) = 0 Then
                    Err.Raise vbObjectError + 1027, "clsSession", "Destino de redirecionamento vazio."
                End If
                If StrComp(redirectPath, "/dev/null", vbTextCompare) = 0 Or _
                   StrComp(redirectPath, "NUL", vbTextCompare) = 0 Then
                    If Left$(redirection.Operator, 9) = "REDIR_ERR" Then
                        stderrType = "DISCARD"
                    Else
                        stdoutType = "DISCARD"
                    End If
                Else
                    redirectPath = ResolvePath(redirectPath)
                    If redirection.Operator = "REDIR_OUT" Or redirection.Operator = "REDIR_ERR" Then
                        TruncateOutputFile redirectPath
                    Else
                        EnsureOutputFile redirectPath
                    End If
                    If Left$(redirection.Operator, 9) = "REDIR_ERR" Then
                        stderrType = "FILE"
                        stderrPath = redirectPath
                        stderrAppend = (redirection.Operator = "REDIR_ERR_APPEND")
                    Else
                        stdoutType = "FILE"
                        stdoutPath = redirectPath
                        stdoutAppend = (redirection.Operator = "REDIR_APPEND")
                    End If
                End If
            Case "REDIR_DISCARD"
                stdoutType = "DISCARD"
            Case "REDIR_ERR_DISCARD"
                stderrType = "DISCARD"
            Case "REDIR_ERR_TO_OUT"
                stderrType = stdoutType
                stderrPath = stdoutPath
                stderrAppend = (stdoutType = "FILE")
        End Select
    Next redirection

    If stdoutType = "FILE" And stderrType = "FILE" And _
       StrComp(stdoutPath, stderrPath, vbTextCompare) = 0 Then
        mergedText = output
        mergedHasOutput = HasOutput
        mergedAppend = stdoutAppend
        mergedSuppressNewline = SuppressNewline
        If HasErrorOutput Then
            If mergedHasOutput And Not mergedSuppressNewline Then mergedText = mergedText & vbCrLf
            mergedText = mergedText & errorText
            mergedHasOutput = True
            mergedSuppressNewline = False
        End If
        RouteOutput "FILE", stdoutPath, mergedAppend, mergedText, _
                    mergedHasOutput, mergedSuppressNewline
    Else
        RouteOutput stdoutType, stdoutPath, stdoutAppend, output, HasOutput, SuppressNewline
        RouteOutput stderrType, stderrPath, stderrAppend, errorText, HasErrorOutput, False
    End If
    If stdoutType <> "CONSOLE" Then HasOutput = False
    If stderrType <> "CONSOLE" Then
        HasErrorOutput = False
        ErrorOutput = ""
    End If
End Sub

Private Sub RouteOutput(ByVal destinationType As String, ByVal destinationPath As String, _
                        ByVal append As Boolean, ByVal text As String, ByVal hasText As Boolean, _
                        ByVal suppressNewline As Boolean)
    Dim stream As clsFileStream
    Dim content As String

    If destinationType = "CONSOLE" Or destinationType = "DISCARD" Then Exit Sub
    If Not hasText Then Exit Sub

    content = text
    If Not suppressNewline Then content = content & vbCrLf
    Set stream = New clsFileStream
    stream.OpenForWrite destinationPath, append
    stream.WriteText content
    stream.CloseStream
End Sub

Private Sub TruncateOutputFile(ByVal path As String)
    Dim stream As clsFileStream

    Set stream = New clsFileStream
    stream.OpenForWrite path
    stream.CloseStream
End Sub

Private Sub EnsureOutputFile(ByVal path As String)
    Dim fileSystem As Object
    Dim stream As clsFileStream

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FileExists(path) Then Exit Sub
    Set stream = New clsFileStream
    stream.OpenForWrite path
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
