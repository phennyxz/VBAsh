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
            ' Ignora comentário até final da linha
            Do While mIndex <= mLen
                If Mid$(mSource, mIndex, 1) = vbCr Or Mid$(mSource, mIndex, 1) = vbLf Then
                    Exit Do
                End If
                mIndex = mIndex + 1
            Loop
            GoTo ContinueLoop
        End If

        If ch = "|" Then
            AddToken Tokenize, "PIPE", "|"
            mIndex = mIndex + 1
            GoTo ContinueLoop
        End If

        If ch = ";" Then
            AddToken Tokenize, "SEMI", ";"
            mIndex = mIndex + 1
            GoTo ContinueLoop
        End If

        If ch = ">" Then
            If Mid$(mSource, mIndex + 1, 1) = ">" Then
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

        If ch = "'" Or ch = """" Then
            Dim quoteCh As String
            quoteCh = ch
            Dim word As String
            word = ""
            mIndex = mIndex + 1

            Do While mIndex <= mLen
                ch = Mid$(mSource, mIndex, 1)
                If ch = quoteCh Then
                    mIndex = mIndex + 1
                    Exit Do
                End If
                word = word & ch
                mIndex = mIndex + 1
            Loop

            AddToken Tokenize, "WORD", word
            GoTo ContinueLoop
        End If

        Dim buffer As String
        buffer = ""
        Do While mIndex <= mLen
            ch = Mid$(mSource, mIndex, 1)

            If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or _
               ch = "|" Or ch = ";" Or ch = ">" Or ch = "<" Or ch = "(" Or ch = ")" Or ch = "#" Then
                Exit Do
            End If

            buffer = buffer & ch
            mIndex = mIndex + 1
        Loop

        If Len(buffer) > 0 Then
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