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
        Dim hasWord As Boolean
        buffer = ""
        hasWord = False

        Do While mIndex <= mLen
            ch = Mid$(mSource, mIndex, 1)

            If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or _
               ch = "|" Or ch = "&" Or ch = ";" Or ch = ">" Or ch = "<" Or ch = "(" Or ch = ")" Then
                Exit Do
            End If

            If ch = "'" Or ch = """" Then
                Dim quoteCh As String
                Dim quoteClosed As Boolean
                quoteCh = ch
                quoteClosed = False
                hasWord = True
                mIndex = mIndex + 1

                Do While mIndex <= mLen
                    ch = Mid$(mSource, mIndex, 1)
                    If ch = quoteCh Then
                        quoteClosed = True
                        mIndex = mIndex + 1
                        Exit Do
                    End If

                    If quoteCh = """" And ch = "\" And mIndex < mLen Then
                        Dim escapedChar As String
                        escapedChar = Mid$(mSource, mIndex + 1, 1)
                        If escapedChar = """" Or escapedChar = "\" Or escapedChar = "$" Or escapedChar = "`" Then
                            mIndex = mIndex + 1
                            ch = Mid$(mSource, mIndex, 1)
                        Else
                            buffer = buffer & ch
                            mIndex = mIndex + 1
                            GoTo ContinueQuoted
                        End If
                    End If

                    buffer = buffer & ch
                    mIndex = mIndex + 1
ContinueQuoted:
                Loop

                If Not quoteClosed Then Err.Raise vbObjectError + 1000, "clsTokenizer.Tokenize", "Aspas não fechadas."
                GoTo ContinueWord
            End If

            If ch = "\" And mIndex < mLen Then
                mIndex = mIndex + 1
                buffer = buffer & Mid$(mSource, mIndex, 1)
                hasWord = True
                mIndex = mIndex + 1
                GoTo ContinueWord
            End If

            buffer = buffer & ch
            hasWord = True
            mIndex = mIndex + 1
ContinueWord:
        Loop

        If hasWord Then
            AddToken Tokenize, "WORD", buffer
        End If

ContinueLoop:
    Loop

    AddToken Tokenize, "EOF", ""
End Function

Private Sub AddToken(ByVal tokenList As Collection, ByVal tokenType As String, ByVal tokenValue As String)
    Dim token As clsToken
    Set token = New clsToken
    token.Init tokenType, tokenValue
    tokenList.Add token
End Sub