Option Explicit

Private mTokens As Collection
Private mOriginalText As String
Private mIndex As Long

Public Function Parse(ByVal tokens As Collection, ByVal originalText As String) As clsCommand
    Dim root As clsCommand

    Set root = ParseAst(tokens, originalText)
    If root.NodeType <> "SIMPLE_COMMAND" Or HasListOperator(tokens) Then
        Err.Raise vbObjectError + 1001, "clsParser.Parse", _
            "Operador ainda não suportado: " & FirstNonWordOperator(tokens)
    End If
    Set Parse = root
End Function

Private Function HasListOperator(ByVal tokens As Collection) As Boolean
    Dim token As clsToken

    For Each token In tokens
        If token.TokenType = "SEMI" Or IsLogicalOperator(token.TokenType) Then
            HasListOperator = True
            Exit Function
        End If
    Next token
End Function

Public Function ParseAst(ByVal tokens As Collection, ByVal originalText As String) As clsCommand
    If tokens Is Nothing Then
        Err.Raise vbObjectError + 1002, "clsParser.ParseAst", "Comando vazio."
    End If
    If tokens.Count = 0 Then
        Err.Raise vbObjectError + 1002, "clsParser.ParseAst", "Comando vazio."
    End If

    Set mTokens = tokens
    mOriginalText = originalText
    mIndex = 1
    Set ParseAst = ParseSequence()
End Function

Private Function ParseSequence() As clsCommand
    Dim expressions As New Collection
    Dim node As clsCommand
    Dim expression As clsCommand

    If CurrentTokenType() = "EOF" Then
        Err.Raise vbObjectError + 1002, "clsParser.ParseAst", "Comando vazio."
    End If

    expressions.Add ParseAndOr()
    Do While CurrentTokenType() = "SEMI"
        mIndex = mIndex + 1
        If CurrentTokenType() = "EOF" Then Exit Do
        If IsLogicalOperator(CurrentTokenType()) Or CurrentTokenType() = "SEMI" Then
            RaiseUnexpectedOperator
        End If
        expressions.Add ParseAndOr()
    Loop

    If CurrentTokenType() <> "EOF" Then RaiseUnexpectedOperator
    If expressions.Count = 1 Then
        Set ParseSequence = expressions(1)
        Exit Function
    End If
    Set node = New clsCommand
    node.InitAstNode "SEQUENCE", expressions(1).SourceStart, _
                     expressions(expressions.Count).SourceEnd
    node.OriginalText = mOriginalText
    For Each expression In expressions
        node.AddChild expression
    Next expression
    Set ParseSequence = node
End Function

Private Function ParseAndOr() As clsCommand
    Dim leftNode As clsCommand
    Dim rightNode As clsCommand
    Dim parentNode As clsCommand
    Dim operatorType As String

    Set leftNode = ParseSimpleCommand()
    Do While IsLogicalOperator(CurrentTokenType())
        operatorType = CurrentTokenType()
        mIndex = mIndex + 1
        Set rightNode = ParseSimpleCommand()

        Set parentNode = New clsCommand
        parentNode.InitAstNode operatorType, leftNode.SourceStart, rightNode.SourceEnd
        parentNode.OriginalText = mOriginalText
        parentNode.AddChild leftNode
        parentNode.AddChild rightNode
        Set leftNode = parentNode
    Loop
    Set ParseAndOr = leftNode
End Function

Private Function ParseSimpleCommand() As clsCommand
    Dim node As clsCommand
    Dim token As clsToken
    Dim firstWord As clsToken
    Dim redirection As clsRedirection
    Dim expectsTarget As Boolean
    Dim startPosition As Long
    Dim endPosition As Long

    If CurrentTokenType() <> "WORD" Then RaiseUnexpectedOperator
    Set node = New clsCommand
    node.InitAstNode "SIMPLE_COMMAND", 0, 0
    startPosition = CurrentSourceStart()

    Do While CurrentTokenType() <> "EOF" And CurrentTokenType() <> "SEMI" And _
             Not IsLogicalOperator(CurrentTokenType())
        Set token = CurrentToken()
        Select Case token.TokenType
            Case "WORD"
                If expectsTarget Then
                    redirection.Target = token.Value
                    node.AddRedirection redirection
                    expectsTarget = False
                Else
                    node.AddWordToken token
                    If firstWord Is Nothing Then
                        Set firstWord = token
                        node.Name = token.Value
                    Else
                        node.AddArg token.Value
                    End If
                End If
            Case "REDIR_OUT", "REDIR_APPEND", "REDIR_ERR", "REDIR_ERR_APPEND"
                If expectsTarget Then
                    Err.Raise vbObjectError + 1005, "clsParser.ParseAst", _
                        "Destino de redirecionamento ausente."
                End If
                Set redirection = New clsRedirection
                redirection.Operator = token.TokenType
                expectsTarget = True
            Case "REDIR_DISCARD", "REDIR_ERR_DISCARD", "REDIR_ERR_TO_OUT"
                If expectsTarget Then
                    Err.Raise vbObjectError + 1005, "clsParser.ParseAst", _
                        "Destino de redirecionamento ausente."
                End If
                Set redirection = New clsRedirection
                redirection.Operator = token.TokenType
                If token.TokenType = "REDIR_DISCARD" Or token.TokenType = "REDIR_ERR_DISCARD" Then
                    redirection.Target = "-"
                Else
                    redirection.Target = "1"
                End If
                node.AddRedirection redirection
            Case Else
                RaiseUnexpectedOperator
        End Select

        endPosition = token.SourceEnd
        mIndex = mIndex + 1
    Loop

    If expectsTarget Then
        Err.Raise vbObjectError + 1003, "clsParser.ParseAst", "Destino de redirecionamento ausente."
    End If
    If firstWord Is Nothing Then
        Err.Raise vbObjectError + 1002, "clsParser.ParseAst", "Comando vazio."
    End If

    node.OriginalText = mOriginalText
    node.SourceStart = startPosition
    node.SourceEnd = endPosition

    Set ParseSimpleCommand = node
End Function

Private Function CurrentToken() As clsToken
    If mIndex < 1 Or mIndex > mTokens.Count Then Exit Function
    Set CurrentToken = mTokens(mIndex)
End Function

Private Function CurrentTokenType() As String
    Dim token As clsToken

    Set token = CurrentToken()
    If Not token Is Nothing Then CurrentTokenType = token.TokenType
End Function

Private Function CurrentSourceStart() As Long
    Dim token As clsToken

    Set token = CurrentToken()
    If Not token Is Nothing Then CurrentSourceStart = token.SourceStart
End Function

Private Function IsLogicalOperator(ByVal tokenType As String) As Boolean
    IsLogicalOperator = (tokenType = "AND_IF" Or tokenType = "OR_IF")
End Function

Private Sub RaiseUnexpectedOperator()
    Dim token As clsToken

    Set token = CurrentToken()
    If token Is Nothing Then
        Err.Raise vbObjectError + 1083, "clsParser.ParseAst", "Expressão incompleta."
    End If
    If token.TokenType = "EOF" Then
        Err.Raise vbObjectError + 1083, "clsParser.ParseAst", "Expressão incompleta."
    End If
    Err.Raise vbObjectError + 1001, "clsParser.ParseAst", _
        "Operador ainda não suportado: " & token.Value
End Sub

Private Function FirstNonWordOperator(ByVal tokens As Collection) As String
    Dim token As clsToken

    For Each token In tokens
        If token.TokenType <> "WORD" And token.TokenType <> "EOF" Then
            FirstNonWordOperator = token.Value
            Exit Function
        End If
    Next token
End Function
