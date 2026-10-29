Option Explicit

Private mSource As String
Private mIndex As Long
Private mLen As Long

Public Function Tokenize(ByVal sourceText As String) As Collection
    Set Tokenize = New Collection
    mSource = sourceText
    mLen = Len(mSource)
    mIndex = 1

    Do While mIndex <= mLen
        Dim ch As String
        ch = Mid$(mSource, mIndex, 1)

        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            mIndex = mIndex + 1
            GoTo ContinueLoop
        End If

        If ch = "#" Then
            Do While mIndex <= mLen
                If Mid$(mSource, mIndex, 1) = vbCr Or Mid$(mSource, mIndex, 1) = vbLf Then
                    Exit Do
                End If
                mIndex = mIndex + 1
            Loop
            GoTo ContinueLoop
        End If

        If Mid$(mSource, mIndex, 4) = "2>&1" Then
            AddToken Tokenize, "REDIR_ERR_TO_OUT", "2>&1"
            mIndex = mIndex + 4
            GoTo ContinueLoop
        End If

        If Mid$(mSource, mIndex, 4) = "2>&-" Then
            AddToken Tokenize, "REDIR_ERR_DISCARD", "2>&-"
            mIndex = mIndex + 4
            GoTo ContinueLoop
        End If

        If Mid$(mSource, mIndex, 3) = "2>>" Then
            AddToken Tokenize, "REDIR_ERR_APPEND", "2>>"
            mIndex = mIndex + 3
            GoTo ContinueLoop
        End If

        If Mid$(mSource, mIndex, 2) = "2>" Then
            AddToken Tokenize, "REDIR_ERR", "2>"
            mIndex = mIndex + 2
            GoTo ContinueLoop
        End If

        If ch = "|" Then
            If Mid$(mSource, mIndex + 1, 1) = "|" Then
                AddToken Tokenize, "OR_IF", "||"
                mIndex = mIndex + 2
            Else
                AddToken Tokenize, "PIPE", "|"
                mIndex = mIndex + 1
            End If
            GoTo ContinueLoop
        End If

        If ch = "&" Then
            If Mid$(mSource, mIndex + 1, 1) = "&" Then
                AddToken Tokenize, "AND_IF", "&&"
                mIndex = mIndex + 2
                GoTo ContinueLoop
            End If
            Err.Raise vbObjectError + 1001, "clsTokenizer.Tokenize", "Operador '&' ainda não suportado."
        End If

        If ch = ";" Then
            AddToken Tokenize, "SEMI", ";"
            mIndex = mIndex + 1
            GoTo ContinueLoop
        End If

        If ch = ">" Then
            If Mid$(mSource, mIndex + 1, 2) = "&-" Then
                AddToken Tokenize, "REDIR_DISCARD", ">&-"
                mIndex = mIndex + 3
            ElseIf Mid$(mSource, mIndex + 1, 1) = ">" Then
                AddToken Tokenize, "REDIR_APPEND", ">>"
                mIndex = mIndex + 2
            Else
                AddToken Tokenize, "REDIR_OUT", ">"
                mIndex = mIndex + 1
            End If
            GoTo ContinueLoop
        End If

        If ch = "<" Then
            AddToken Tokenize, "REDIR_IN", "<"
            mIndex = mIndex + 1
            GoTo ContinueLoop
        End If

        If ch = "(" Then
            AddToken Tokenize, "LPAREN", "("
            mIndex = mIndex + 1
            GoTo ContinueLoop
        End If

        If ch = ")" Then
            AddToken Tokenize, "RPAREN", ")"
            mIndex = mIndex + 1
            GoTo ContinueLoop
        End If

        Dim buffer As String
        Dim bufferMask As String
        Dim hasWord As Boolean
        Dim braceExpansionAllowed As Boolean
        Dim wordSegments As Collection
        Dim token As clsToken
        Dim wordStart As Long
        Dim segmentText As String
        Dim segmentStart As Long
        Dim segmentEnd As Long
        Dim hasSegment As Boolean
        Dim quotedText As String
        Dim quotedStart As Long
        Dim quoteStart As Long
        Dim quoteEnd As Long
        Dim escapedChar As String
        buffer = ""
        bufferMask = ""
        hasWord = False
        braceExpansionAllowed = True
        Set wordSegments = New Collection
        wordStart = mIndex
        hasSegment = False

        Do While mIndex <= mLen
            ch = Mid$(mSource, mIndex, 1)

            If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or _
               ch = "|" Or ch = "&" Or ch = ";" Or ch = ">" Or ch = "<" Or ch = "(" Or ch = ")" Then
                Exit Do
            End If

            If ch = "'" Or ch = """" Then
                Dim quoteCh As String
                Dim quoteClosed As Boolean
                If hasSegment Then
                    AddWordSegment wordSegments, segmentText, "UNQUOTED", False, segmentStart, segmentEnd
                    hasSegment = False
                    segmentText = ""
                End If
                braceExpansionAllowed = False
                quoteCh = ch
                quoteClosed = False
                hasWord = True
                quoteStart = mIndex
                mIndex = mIndex + 1

                If quoteCh = "'" Then
                    quotedText = ""
                    Do While mIndex <= mLen
                        ch = Mid$(mSource, mIndex, 1)
                        If ch = quoteCh Then
                            quoteClosed = True
                            quoteEnd = mIndex
                            mIndex = mIndex + 1
                            Exit Do
                        End If
                        quotedText = quotedText & ch
                        buffer = buffer & ch
                        bufferMask = bufferMask & "0"
                        mIndex = mIndex + 1
                    Loop
                    If quoteClosed Then
                        If Len(quotedText) = 0 Then
                            AddWordSegment wordSegments, "", "SINGLE", False, quoteStart, quoteEnd
                        Else
                            AddWordSegment wordSegments, quotedText, "SINGLE", False, _
                                           quoteStart + 1, quoteEnd - 1
                        End If
                    End If
                Else
                    quotedText = ""
                    quotedStart = 0
                    Do While mIndex <= mLen
                        ch = Mid$(mSource, mIndex, 1)
                        If ch = quoteCh Then
                            quoteClosed = True
                            quoteEnd = mIndex
                            mIndex = mIndex + 1
                            Exit Do
                        End If

                        If ch = "\" And mIndex < mLen Then
                            escapedChar = Mid$(mSource, mIndex + 1, 1)
                            If escapedChar = """" Or escapedChar = "\" Or _
                               escapedChar = "$" Or escapedChar = "`" Then
                                If Len(quotedText) > 0 Then
                                    AddWordSegment wordSegments, quotedText, "DOUBLE", False, _
                                                   quotedStart, mIndex - 1
                                    quotedText = ""
                                    quotedStart = 0
                                End If
                                buffer = buffer & escapedChar
                                bufferMask = bufferMask & IIf(escapedChar = "$", "0", "1")
                                AddWordSegment wordSegments, escapedChar, "DOUBLE", True, mIndex, mIndex + 1
                                mIndex = mIndex + 2
                                GoTo ContinueQuoted
                            End If
                        End If

                        If quotedStart = 0 Then quotedStart = mIndex
                        quotedText = quotedText & ch
                        buffer = buffer & ch
                        bufferMask = bufferMask & "1"
                        mIndex = mIndex + 1
ContinueQuoted:
                    Loop
                    If quoteClosed Then
                        If Len(quotedText) > 0 Then
                            AddWordSegment wordSegments, quotedText, "DOUBLE", False, _
                                           quotedStart, quoteEnd - 1
                        ElseIf quoteEnd = quoteStart + 1 Then
                            AddWordSegment wordSegments, "", "DOUBLE", False, quoteStart, quoteEnd
                        End If
                    End If
                End If

                If Not quoteClosed Then Err.Raise vbObjectError + 1000, "clsTokenizer.Tokenize", "Aspas não fechadas."
                GoTo ContinueWord
            End If

            If ch = "\" And mIndex < mLen Then
                braceExpansionAllowed = False
                If hasSegment Then
                    AddWordSegment wordSegments, segmentText, "UNQUOTED", False, segmentStart, segmentEnd
                    hasSegment = False
                    segmentText = ""
                End If
                mIndex = mIndex + 1
                escapedChar = Mid$(mSource, mIndex, 1)
                buffer = buffer & escapedChar
                bufferMask = bufferMask & IIf(escapedChar = "$", "0", "1")
                AddWordSegment wordSegments, escapedChar, "UNQUOTED", True, mIndex - 1, mIndex
                hasWord = True
                mIndex = mIndex + 1
                GoTo ContinueWord
            End If

            If Not hasSegment Then
                segmentStart = mIndex
                segmentText = ""
                hasSegment = True
            End If
            segmentEnd = mIndex
            segmentText = segmentText & ch
            buffer = buffer & ch
            bufferMask = bufferMask & "1"
            hasWord = True
            mIndex = mIndex + 1
ContinueWord:
        Loop

        If hasWord Then
            If hasSegment Then
                AddWordSegment wordSegments, segmentText, "UNQUOTED", False, segmentStart, segmentEnd
            End If
            AddToken Tokenize, "WORD", buffer, bufferMask, braceExpansionAllowed
            Set token = Tokenize(Tokenize.Count)
            Set token.Segments = wordSegments
            token.SourceStart = wordStart
            token.SourceEnd = mIndex - 1
        End If

ContinueLoop:
    Loop

    AddToken Tokenize, "EOF", ""
End Function

Private Sub AddToken(ByVal tokenList As Collection, ByVal tokenType As String, _
                     ByVal tokenValue As String, Optional ByVal expansionMask As String = "", _
                     Optional ByVal braceExpansionAllowed As Boolean = True)
    Dim token As clsToken
    Set token = New clsToken
    token.Init tokenType, tokenValue, expansionMask, braceExpansionAllowed
    If tokenType = "EOF" Then
        token.SourceStart = mLen + 1
        token.SourceEnd = mLen + 1
    Else
        token.SourceStart = mIndex
        token.SourceEnd = mIndex + Len(tokenValue) - 1
    End If
    tokenList.Add token
End Sub

Private Sub AddWordSegment(ByVal segments As Collection, ByVal segmentText As String, _
                           ByVal quoteContext As String, ByVal isEscaped As Boolean, _
                           ByVal sourceStart As Long, ByVal sourceEnd As Long)
    Dim segment As clsWordSegment

    Set segment = New clsWordSegment
    segment.Init segmentText, quoteContext, isEscaped, sourceStart, sourceEnd
    segments.Add segment
End Sub