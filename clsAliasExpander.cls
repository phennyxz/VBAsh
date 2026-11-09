' @VBAshRevision clsAliasExpander-r1
Option Explicit

Private Const MAX_ALIAS_EXPANSION_DEPTH As Long = 32

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsAliasExpander-r1"
End Property

Public Function ExpandTokens(ByVal sourceTokens As Collection, _
                             ByVal registry As clsAliasRegistry) As Collection
    Dim activeAliases As Object
    Dim expanded As Collection
    Dim finalCommandPosition As Boolean
    Dim token As clsToken
    Dim eofToken As clsToken

    If sourceTokens Is Nothing Then
        Err.Raise vbObjectError + 1410, "clsAliasExpander.ExpandTokens", _
                  "A coleção de tokens é obrigatória."
    End If
    If registry Is Nothing Then
        Err.Raise vbObjectError + 1411, "clsAliasExpander.ExpandTokens", _
                  "O registry de aliases é obrigatório."
    End If

    Set activeAliases = CreateObject("Scripting.Dictionary")
    activeAliases.CompareMode = vbBinaryCompare
    Set expanded = ExpandSequence(sourceTokens, registry, activeAliases, 0, _
                                  finalCommandPosition)

    If sourceTokens.Count > 0 Then
        Set token = sourceTokens(sourceTokens.Count)
        If token.TokenType = "EOF" Then Set eofToken = token
    End If
    If eofToken Is Nothing Then
        Set eofToken = New clsToken
        eofToken.Init "EOF", ""
    End If
    expanded.Add eofToken
    Set ExpandTokens = expanded
End Function

Private Function ExpandSequence(ByVal sourceTokens As Collection, _
                                ByVal registry As clsAliasRegistry, _
                                ByVal activeAliases As Object, _
                                ByVal depth As Long, _
                                ByRef finalCommandPosition As Boolean) As Collection
    Dim result As New Collection
    Dim token As clsToken
    Dim replacementToken As clsToken
    Dim replacementTokens As Collection
    Dim expandedReplacement As Collection
    Dim tokenizer As clsTokenizer
    Dim nestedCommandPosition As Boolean
    Dim commandPosition As Boolean
    Dim expectsRedirectionTarget As Boolean
    Dim aliasName As String
    Dim replacementText As String
    Dim i As Long

    commandPosition = True

    For i = 1 To sourceTokens.Count
        Set token = sourceTokens(i)
        If token.TokenType = "EOF" Then Exit For

        If expectsRedirectionTarget Then
            result.Add token
            expectsRedirectionTarget = False
            GoTo ContinueToken
        End If

        If IsRedirectionWithTarget(token.TokenType) Then
            result.Add token
            expectsRedirectionTarget = True
            GoTo ContinueToken
        End If

        If commandPosition And token.TokenType = "WORD" Then
            If IsAssignmentWord(token) Then
                result.Add token
                GoTo ContinueToken
            End If

            If IsUnquotedWord(token) Then
                aliasName = token.Value
                If registry.HasAlias(aliasName) Then
                    If activeAliases.Exists(aliasName) Then
                        ' A currently expanding alias is emitted literally.
                        result.Add token
                        commandPosition = False
                    Else
                        If depth >= MAX_ALIAS_EXPANSION_DEPTH Then
                            Err.Raise vbObjectError + 1412, _
                                      "clsAliasExpander.ExpandTokens", _
                                      "Profundidade máxima de expansão de aliases excedida."
                        End If

                        replacementText = registry.ResolveAlias(aliasName)
                        activeAliases.Add aliasName, True
                        Set tokenizer = New clsTokenizer
                        Set replacementTokens = tokenizer.Tokenize(replacementText)
                        For Each replacementToken In replacementTokens
                            If replacementToken.TokenType <> "EOF" Then
                                replacementToken.SourceStart = token.SourceStart
                                replacementToken.SourceEnd = token.SourceEnd
                            End If
                        Next replacementToken

                        Set expandedReplacement = ExpandSequence( _
                            replacementTokens, registry, activeAliases, depth + 1, _
                            nestedCommandPosition)
                        activeAliases.Remove aliasName

                        For Each replacementToken In expandedReplacement
                            result.Add replacementToken
                        Next replacementToken
                        commandPosition = nestedCommandPosition
                    End If
                Else
                    result.Add token
                    commandPosition = False
                End If
            Else
                result.Add token
                commandPosition = False
            End If
        Else
            result.Add token
            Select Case token.TokenType
                Case "SEMI", "PIPE", "AND_IF", "OR_IF", "LPAREN"
                    commandPosition = True
                Case "RPAREN"
                    commandPosition = False
            End Select
        End If

ContinueToken:
    Next i

    finalCommandPosition = commandPosition
    Set ExpandSequence = result
End Function

Private Function IsRedirectionWithTarget(ByVal tokenType As String) As Boolean
    Select Case tokenType
        Case "REDIR_OUT", "REDIR_APPEND", "REDIR_ERR", _
             "REDIR_ERR_APPEND", "REDIR_IN"
            IsRedirectionWithTarget = True
    End Select
End Function

Private Function IsUnquotedWord(ByVal token As clsToken) As Boolean
    Dim segment As clsWordSegment

    If token.Segments Is Nothing Then
        IsUnquotedWord = True
        Exit Function
    End If
    If token.Segments.Count = 0 Then
        IsUnquotedWord = True
        Exit Function
    End If

    For Each segment In token.Segments
        If segment.QuoteContext <> "UNQUOTED" Or segment.IsEscaped Then Exit Function
    Next segment
    IsUnquotedWord = True
End Function

Private Function IsAssignmentWord(ByVal token As clsToken) As Boolean
    Dim separator As Long
    Dim namePart As String
    Dim i As Long
    Dim ch As String
    Dim segment As clsWordSegment
    Dim checkedCharacters As Long

    separator = InStr(1, token.Value, "=", vbBinaryCompare)
    If separator <= 1 Then Exit Function
    namePart = Left$(token.Value, separator - 1)

    For Each segment In token.Segments
        For i = 1 To Len(segment.Text)
            If checkedCharacters >= Len(namePart) Then Exit For
            If segment.QuoteContext <> "UNQUOTED" Or segment.IsEscaped Then Exit Function
            checkedCharacters = checkedCharacters + 1
        Next i
        If checkedCharacters >= Len(namePart) Then Exit For
    Next segment
    If checkedCharacters <> Len(namePart) Then Exit Function

    For i = 1 To Len(namePart)
        ch = Mid$(namePart, i, 1)
        If i = 1 Then
            If Not (ch = "_" Or (ch >= "A" And ch <= "Z") Or _
                    (ch >= "a" And ch <= "z")) Then Exit Function
        Else
            If Not (ch = "_" Or (ch >= "A" And ch <= "Z") Or _
                    (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9")) Then
                Exit Function
            End If
        End If
    Next i

    IsAssignmentWord = True
End Function
