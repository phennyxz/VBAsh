Option Explicit

Public Function Execute(ByVal session As clsSession, ByVal node As clsCommand) As String
    If session Is Nothing Then
        Err.Raise vbObjectError + 1120, "clsAstExecutor.Execute", "Sessão ausente."
    End If
    If node Is Nothing Then
        Err.Raise vbObjectError + 1121, "clsAstExecutor.Execute", "Nó AST ausente."
    End If

    Execute = ExecuteNode(session, node)
End Function

Private Function ExecuteNode(ByVal session As clsSession, ByVal node As clsCommand) As String
    Select Case node.NodeType
        Case "SIMPLE_COMMAND"
            ExecuteNode = session.ExecuteAstSimple(node)
        Case "SEQUENCE"
            ExecuteNode = ExecuteSequence(session, node)
        Case "AND_IF", "OR_IF"
            ExecuteNode = ExecuteLogical(session, node)
        Case Else
            Err.Raise vbObjectError + 1122, "clsAstExecutor.ExecuteNode", _
                      "Tipo de nó AST não executável: " & node.NodeType
    End Select
End Function

Private Function ExecuteSequence(ByVal session As clsSession, ByVal node As clsCommand) As String
    Dim child As clsCommand
    Dim childOutput As String
    Dim output As String
    Dim errors As String
    Dim hasOutput As Boolean
    Dim hasErrors As Boolean
    Dim suppressNewline As Boolean

    If node.Children.Count = 0 Then
        Err.Raise vbObjectError + 1123, "clsAstExecutor.ExecuteSequence", _
                  "Sequência AST sem comandos."
    End If

    For Each child In node.Children
        childOutput = ExecuteNode(session, child)
        AppendChildResult session, childOutput, output, hasOutput, suppressNewline, _
                          errors, hasErrors
    Next child

    SetSessionResult session, output, hasOutput, suppressNewline, errors, hasErrors
    ExecuteSequence = output
End Function

Private Function ExecuteLogical(ByVal session As clsSession, ByVal node As clsCommand) As String
    Dim leftNode As clsCommand
    Dim rightNode As clsCommand
    Dim leftOutput As String
    Dim rightOutput As String
    Dim output As String
    Dim errors As String
    Dim hasOutput As Boolean
    Dim hasErrors As Boolean
    Dim suppressNewline As Boolean
    Dim shouldExecuteRight As Boolean

    If node.Children.Count <> 2 Then
        Err.Raise vbObjectError + 1124, "clsAstExecutor.ExecuteLogical", _
                  "Operador lógico AST deve ter dois operandos."
    End If

    Set leftNode = node.Children(1)
    Set rightNode = node.Children(2)
    leftOutput = ExecuteNode(session, leftNode)
    AppendChildResult session, leftOutput, output, hasOutput, suppressNewline, _
                      errors, hasErrors

    If node.NodeType = "AND_IF" Then
        shouldExecuteRight = (session.LastExitCode = 0)
    Else
        shouldExecuteRight = (session.LastExitCode <> 0)
    End If

    If shouldExecuteRight Then
        rightOutput = ExecuteNode(session, rightNode)
        AppendChildResult session, rightOutput, output, hasOutput, suppressNewline, _
                          errors, hasErrors
    End If

    SetSessionResult session, output, hasOutput, suppressNewline, errors, hasErrors
    ExecuteLogical = output
End Function

Private Sub AppendChildResult(ByVal session As clsSession, ByVal childOutput As String, _
                              ByRef output As String, ByRef hasOutput As Boolean, _
                              ByRef suppressNewline As Boolean, ByRef errors As String, _
                              ByRef hasErrors As Boolean)
    If session.HasOutput Then
        If hasOutput And Not suppressNewline Then output = output & vbCrLf
        output = output & childOutput
        hasOutput = True
        suppressNewline = session.SuppressNewline
    End If

    If session.HasErrorOutput Then
        If hasErrors Then errors = errors & vbCrLf
        errors = errors & session.ErrorOutput
        hasErrors = True
    End If
End Sub

Private Sub SetSessionResult(ByVal session As clsSession, ByVal output As String, _
                             ByVal hasOutput As Boolean, ByVal suppressNewline As Boolean, _
                             ByVal errors As String, ByVal hasErrors As Boolean)
    session.HasOutput = hasOutput
    session.SuppressNewline = suppressNewline
    session.ErrorOutput = errors
    session.HasErrorOutput = hasErrors
End Sub
